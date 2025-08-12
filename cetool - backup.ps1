<# Centraal Examen tool om een excelbestand van Eduarte om te zetten naar een voor Facet geschikt format.
   Deze tool is gemaakt door Benvindo Neves
#>

$programma = @{
    versie = '0.0.5'
    extralabel = 'alpha.250812' # extra label voor de alpha versie
    mode = 'alpha' # alpha, beta, prerelease of release
    auteur = 'Benvindo Neves'
    github = "https://api.github.com/repos/examencentrumtcr/cetool/contents/latest"
}

# toevoegen .NET framework klassen
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.IO.Compression.FileSystem

# Bestand waarin alle uitgevoerde stappen worden bijgehouden.
$LogFile = "$PSScriptRoot\script_log.txt"

# Bepalen van het bestand met gebruikersinstellingen.
$gebruikersbestand = "$PSScriptRoot\gebruikersinstellingen.json"

# De ingelezen data uit het excelbestand
$Exceldata = [PSCustomObject]@{
        aantal = 0
        inhoud = @()
        geselecteert = ""
        uitvoermap = ""
        uitvoernaam = ""
        brinnummers = ""
        groepsidnrs = ""
        examendagen = ""
        vakcodes = ""
    }

# gebruikersinstellingen worden opgeslagen in een hash tabel
# en kunnen worden aangepast door de gebruiker.
# Deze worden opgeslagen in een json bestand.
# De standaard instellingen worden in de functie gebruikersinstellingen gedeclareerd.
$gebruiker = @{}

function gebruikersinstellingen {
# hier worden de standaardinstellingen gedeclareerd die gebruikers kunnen wijzigen
$Std_inst=@{
    startmap  = ""       # standaard map waar de gebruiker het script start
    uitvoermap = ""      # standaard map waar de gebruiker de uitvoer plaatst
    uitvoernaam = ""     # standaard naam voor de uitvoer
    aantaldagenlogs = 30 # aantal dagen dat de gebeurtenissen in het logboeken worden bewaard
    }
return $Std_inst
} # einde gebruikersinstellingen

Function StandardInOutputDirectory {
# Standaard in- en uitvoer voor de applicatie
[string]$std_inoutput = [Environment]::GetFolderPath("MyDocuments")
return $std_inoutput
}

Function StandardOutputFileName {
# Standaard uitvoernaam voor het omgezette bestand
[string]$std_outputname = "Facetbestand_" + (Get-Date -Format "yyyyMMdd_HHmmss")
return $std_outputname
}

Function Empty_Exceldata {
# Leeg maken van array en standaard waarden geven
$Exceldata.aantal = 0
$Exceldata.inhoud = @()
$Exceldata.geselecteert = ""

if ($gebruiker.uitvoermap -eq "") {
    $Exceldata.uitvoermap = StandardInOutputDirectory # standaard uitvoermap
} else {
    $Exceldata.uitvoermap = $gebruiker.uitvoermap
}
if ($gebruiker.uitvoernaam -eq "") {
    $Exceldata.uitvoernaam = StandardOutputFileName # standaard uitvoernaam
} else {
    $Exceldata.uitvoernaam = $gebruiker.uitvoernaam
}
$Exceldata.brinnummers = ""
$Exceldata.groepsidnrs = ""
$Exceldata.examendagen = ""
$Exceldata.vakcodes = ""
} # einde Empty_Exceldata

Function declareren_standaardvenster {
param (
    [Parameter(Mandatory = $true)] [string]$titel,
    [Parameter(Mandatory = $true)] [string]$size_x,
    [Parameter(Mandatory = $true)] [string]$size_y
)

$StandaardForm                            = New-Object system.Windows.Forms.Form
$StandaardForm.MaximumSize = New-Object System.Drawing.size($size_x,$size_y)
$StandaardForm.MinimumSize = New-Object System.Drawing.size($size_x,$size_y)
$StandaardForm.text                       = $titel
$StandaardForm.TopMost                    = $true
$StandaardForm.StartPosition              = 'CenterScreen'
# $StandaardForm.BackColor = "white"
$StandaardForm.MaximizeBox = $False
$StandaardForm.font                       = Standaardfont
$StandaardForm.Icon                       = [System.Drawing.Icon]::ExtractAssociatedIcon('cetool_icoon.ico')

return $StandaardForm
} # einde declareren_standaardvenster

Function Standaardfont {
    # Deze functie retourneert een standaard font voor de applicatie)
    $standaard = New-Object System.Drawing.Font('Microsoft Sans Serif',11)
    return $standaard
}

function Write-Log {
    param(
        [Parameter(Mandatory = $true)] [string]$Message,
        [switch]$Notimestamp
        )

    # datum en tijd van de log
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

    # bericht naar tijdelijke log
    $tijdelijkelog = 'tijdelijkelog.txt'
    # als -Notimestamp is meegegeven dan een regel schrijven zonder de datum en tijd.
    if ($Notimestamp) {
        Set-Content -Path $tijdelijkelog -Value "$Message"
    } else {
        Set-Content -Path $tijdelijkelog -Value "$timestamp`t$Message"
    }
    

    # oude inhoud inlezen, toevoegen aan tijdelijke bestand en oude bestand verwijderen
    if (Test-Path $LogFile) {
        $inhoudlog = Get-Content -path $LogFile
        Add-Content -Path $tijdelijkelog -Value $inhoudlog
        Remove-Item $LogFile
    } 
    # naam wijzigen naar oude bestand
    Rename-Item -Path $tijdelijkelog -NewName $LogFile
    
} # einde Write-Log

function Get-PNGImage {
    param([string]$path)
    if (Test-Path $path) {
        return [System.Drawing.Image]::FromFile($path)
    } else {
        return $null
    }
} # einde Get-PNGImage

Function ReadSettings {

# standaard instellingen bepalen voor object $init
$init = gebruikersinstellingen

# inlezen initialisatiebestand als deze bestaat en toevoegen of verwijderen waarden van $init
# dit is nodig zodat, als de initialisatiebestand wordt ingelezen, eventuele nieuwe variabelen worden behouden en verwijderde variabelen niet worden toegevoegd. 
if (test-path -path $gebruikersbestand -pathtype leaf) { 
    # inlezen van object als hashwaarden. hier staan de gewijzigde waarden in die je wil behouden.
    $myObject = Get-Content -Path $gebruikersbestand | ConvertFrom-Json

    # waarden overzetten naar $init. nieuwe waarden worden behouden. verwijderde waarden worden niet toegevoegd.
    foreach( $property in $myobject.psobject.properties.name ) {
 
        # alleen toevoegen als deze bij 'init' al bestaat
        if ( $init.$property -ne $null) { 
            $init[$property] = $myObject.$property
            }

        } # einde foreach $property - loop    

    # gekozen is om dit altijd te bewaren zodat je lijst met variabelen up to date is.
    # $init | ConvertTo-Json -depth 1 | Set-Content -Path $gebruikersbestand
    } 

    return $init

} # einde ReadSettings

function SaveSettings {
    param(
        [Parameter(Mandatory = $true)] [hashtable]$init
    )

    <# Onderstaand is door ChatGpt gegenereerd. Nog niet getest.
    
    inlezen van het huidige gebruikersinstellingen bestand
    if (test-path -path $gebruikersbestand -pathtype leaf) { 
        $myObject = Get-Content -Path $gebruikersbestand | ConvertFrom-Json
    } else {
        $myObject = @{}
    }

    # waarden overzetten naar $myObject. nieuwe waarden worden toegevoegd. verwijderde waarden worden niet toegevoegd.
    foreach( $property in $init.psobject.properties.name ) {
        foreach( $subproperty in $init.$property.psobject.properties.name )
        {
            $myObject[$property][$subproperty] = $init.$property.$subproperty 
        } # einde foreach $subproperty - loop 
    } # einde foreach $property - loop    
    #>

    # schrijven naar bestand
    # $myObject | ConvertTo-Json -depth 1 | Set-Content -Path $gebruikersbestand

    $init | ConvertTo-Json -depth 1 | Set-Content -Path $gebruikersbestand

} # einde SaveSettings
function SelectExcelForm {
    # Deze functie toont een dialoogvenster om een Excelbestand te selecteren
    # en retourneert het pad naar het geselecteerde bestand of 'GEEN' als er geen bestand is geselecteerd.
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Filter = "Excel bestanden (*.xlsx)|*.xlsx"
    $dialog.Title = "Kies een Excelbestand"
    if ($gebruiker.startmap -eq "") {
        $dialog.InitialDirectory = StandardInOutputDirectory # standaard zoekmap voor Excelbestanden
    } else {
        $dialog.InitialDirectory = $gebruiker.startmap # standaard map is de map waar het script staat
    }

    if ($dialog.ShowDialog() -eq "OK") {
        $selectedFile = $dialog.FileName
    } else {
        $selectedFile = 'GEEN'
    }
return $selectedFile
} # einde SelectExcelForm

function Show-Logboek {
    
    $logForm = declareren_standaardvenster -titel "Logboek" -size_x 900 -size_y 400

    $txtLog = New-Object System.Windows.Forms.TextBox
    #$txtLog.Dock = "Fill"
    $txtLog.Location = New-Object System.Drawing.Size(1,1) 
    $txtLog.Size = New-Object System.Drawing.Size(880,310)
    $txtLog.Multiline = $true
    $txtLog.ReadOnly = $true
    # Om de tekst niet selecteerbaar te maken, dus in het blauw, moet tabstop = $false
    $txtLog.TabStop = $false
    # om beide scrollbars te zien moet wordwrap = $false en ScrollBars = 'Both'
    $txtLog.WordWrap = $false
    $txtLog.ScrollBars = 'Both'
    $txtLog.BackColor  = 'white'
    # $txtLog.font = Standaardfont
    $logForm.Controls.Add($txtLog)

    # inlezen logbestand
    if (Test-Path $LogFile) {
        $txtLog.Text = Get-Content -Path $LogFile -Raw
        } else {
        $txtLog.Text = "Logboek is leeg."
        }

    # Voeg een sluitknop toe
    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Text = "Terug"
    # $btnClose.Dock = "Bottom"
    $btnClose.location = "50,320" 
    $btnClose.size = "150,30"
    $btnClose.BackColor = "White"
    $btnClose.DialogResult = [System.Windows.Forms.DialogResult]::cancel
    $logForm.Controls.Add($btnClose)
       
    $null = $logForm.ShowDialog()
} # einde Show-Logboek


function Add-Output {
    param([string]$text)

    # Functie om tekst toe te voegen aan de TextBox. Wordt gebruikt bij Import-ExcelFile en Show_ExcelForm
    $outputBox.AppendText("$text`r`n")
    $outputBox.SelectionStart = $outputBox.Text.Length
    $outputBox.ScrollToCaret()
}

function Import-ExcelFile {
    param([string]$ExcelPath)

    # alleen de naam weergeven en niet de volledige pas
    $naamexcel = Split-Path -Path $ExcelPath -Leaf

    # weergeven in tekstbox en in console
    Write-Host "Excelbestand $naamexcel wordt geopend."
    Add-Output "Excelbestand $naamexcel wordt geopend."

    # aanmaken object voor Excel en openen
    $excel = New-Object -ComObject Excel.Application
    Write-Host "Excelbestand $naamexcel wordt ingelezen."
    Add-Output "Excelbestand $naamexcel wordt ingelezen."
    $workbook = $excel.Workbooks.Open($ExcelPath)
    Write-Host "Blad 1 wordt geopend."
    Add-Output "Blad 1 wordt geopend."
    $sheet = $workbook.Sheets.Item(1)
    # tekst aanpassen aan inhoud zodat het goed ingelezen word.
    # $sheet.Columns.Item("F").ColumnWidth = 20 
    $sheet.Columns.Item("A").ColumnWidth = 20 
   
    $range = $sheet.UsedRange
    $data = @()
    $rowmax = $range.Rows.Count

    Write-Host "Rij 1 bevat de header en wordt overgeslagen."
    Add-Output "Rij 1 bevat de header en wordt overgeslagen."
    # $huidetekst wordt gebruikt om de tekst "Rij 2 van 10 wordt ingelezen"op 1 regel te houden.
    # dit wordt gedaan door de huidge waarde te onthouden en dan steeds alleen 1 regel toe te voegen.
    $huidigetekst = $outputBox.text
    for ($row = 2; $row -le $rowmax; $row++) {

        Write-Host "`rRij $row van $rowmax wordt ingelezen." -NoNewline
        # Voor Add-Output kan niet hetzelfde worden gedaan als de regel hiervoor omdat het niet werkt.
        $outputBox.text = $huidigetekst + "`nRij $row van $rowmax wordt ingelezen."

        # lees de cellen in de rij
        $kolom2 = ($range.Cells.Item($row, 2)).Text
        $kolom1 = ($range.Cells.Item($row, 1)).Text
        if (![string]::IsNullOrWhiteSpace($kolom2) -and ![string]::IsNullOrWhiteSpace($kolom1)) {
            $line = [PSCustomObject]@{
                GroepsIdImport = $kolom1
                MapNaam        = $kolom2
                Kolom3        = ($range.Cells.Item($row, 3)).Text
                Kolom4        = ($range.Cells.Item($row, 4)).Text
                Kolom5        = ($range.Cells.Item($row, 5)).Text
                Kolom6        = ($range.Cells.Item($row, 6)).Text
                Kolom7        = ($range.Cells.Item($row, 7)).Text
                Kolom8        = ($range.Cells.Item($row, 8)).Text
                Kolom9        = ($range.Cells.Item($row, 9)).Text
                Kolom10        = ($range.Cells.Item($row, 10)).Text
                Kolom11        = ($range.Cells.Item($row, 11)).Text
                Kolom12        = ($range.Cells.Item($row, 12)).Text
                Kolom13        = ($range.Cells.Item($row, 13)).Text
                Kolom14        = ($range.Cells.Item($row, 14)).Text
                Kolom15        = ($range.Cells.Item($row, 15)).Text
                Kolom16        = ($range.Cells.Item($row, 16)).Text
                Kolom17        = ($range.Cells.Item($row, 17)).Text                
                Kolom18        = ($range.Cells.Item($row, 18)).Text
                Kolom19        = ($range.Cells.Item($row, 19)).Text
                Kolom20        = ($range.Cells.Item($row, 20)).Text
                Kolom21        = ($range.Cells.Item($row, 21)).Text
                Kolom22        = ($range.Cells.Item($row, 22)).Text                
                Kolom23        = ($range.Cells.Item($row, 23)).Text
                Kolom24        = ($range.Cells.Item($row, 24)).Text
                Kolom25        = ($range.Cells.Item($row, 25)).Text
                Kolom26        = ($range.Cells.Item($row, 26)).Text
                Kolom27        = ($range.Cells.Item($row, 27)).Text
                Kolom28        = ($range.Cells.Item($row, 28)).Text
            }
            $data += $line
        }
    }
    
    Write-Host ""  # Nieuwe regel na afloop
    Write-Host "Data ingelezen."
    Write-Host "Sluiten Excelbestand."
    Add-Output "`n"
    Add-Output "Data ingelezen."
    # $outputBox.text = $outputBox.text + "`nData ingelezen."

    $workbook.Close($false)
    $excel.Quit()

    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($sheet) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

    [gc]::Collect()
    [gc]::WaitForPendingFinalizers()

    # Alles opslaan in Exceldata en een aantal gegevens om te controleren.
    # Het controleren gebeurt bij Show-ExcelForm
    
    Empty_Exceldata;

    $Exceldata.geselecteert = $ExcelPath
    $Exceldata.inhoud = $data
    $Exceldata.aantal = $rowmax - 1 # aantal rijen minus de header
    # De gevonden data in de array zetten
    # Met Sort-object -Property Values wordt alles op alfabetisch volgorde gezet
    # met `r`n`t de regels onder elkaar en een tab naar rechts
    $brinnrs = $data | Group-Object -Property MapNaam | Sort-object -Property Values
       foreach ($brin in $brinnrs) {
        $Exceldata.brinnummers = -join ($Exceldata.brinnummers,"`r`n`t",$brin.name)
    }
    # De gevonden groepsIdImport-nummers in de array zetten
    $gr_id_imp_nrs = $data | Group-Object -Property GroepsIdImport | Sort-object -Property Values
    foreach ($gr_id in $gr_id_imp_nrs) {
        # $Exceldata.groepsidnrs = $Exceldata.groepsidnrs + $gr_id.name + " - "
        $Exceldata.groepsidnrs = -join ($Exceldata.groepsidnrs,"`r`n`t",$gr_id.name)
    }
    # De gevonden aantal dagen in de array zetten
    $alledagen = $data | Group-Object -Property kolom6 | Sort-object -Property Values
    foreach ($dag in $alledagen) {
        $Exceldata.examendagen = -join ($Exceldata.examendagen,"`r`n`t",$dag.name)
    }
    # De gevonden vakcodes in de array zetten
    $allevakcodes = $data | Group-Object -Property kolom12 | Sort-object -Property Values
    foreach ($vakcode in $allevakcodes) {
        $Exceldata.vakcodes = -join ($Exceldata.vakcodes,"`r`n`t",$vakcode.name)
    }

} # einde Import-ExcelFile

function Show-ExcelForm {
    param([string]$ExcelPath)

    # Declareren venster met standaard waarden
    $ShowExcelForm = declareren_standaardvenster -titel "De data wordt ingelezen." -size_x 600 -size_y 300

    # TextBox voor statusoutput
    $outputBox = New-Object System.Windows.Forms.TextBox
    $outputBox.Multiline = $true
    $outputBox.ScrollBars = "Vertical"
    $outputBox.Size = New-Object System.Drawing.Size(560, 200)
    $outputBox.Location = New-Object System.Drawing.Point(10, 5)
    $outputBox.ReadOnly = $true
    $ShowExcelForm.Controls.Add($outputBox)

    # Uitvoeren als Form is geladen. NB: .add_load werkt niet omdat de Form nog niet volledig is geladen.
    $ShowExcelForm.add_Shown({
    
        # importeren Excel
        if ($ExcelPath -ne 'GEEN') {
        Import-ExcelFile -ExcelPath $ExcelPath
        }
 
        # de tekst aanpassen na het inlezen
        $ShowExcelForm.text = "Controleer de data die is ingelezen."

        # aantal ingelezen regels overnemen. 
        $aantal = $Exceldata.aantal
        $brinnrs = $Exceldata.brinnummers
        $grpidimp = $Exceldata.groepsidnrs
        $alle_dagen = $Exceldata.examendagen
        $vak_codes = $Exceldata.vakcodes
        # Weergeven in de TextBox
        Add-Output " "
        Add-Output "Aantal ingelezen rijen : $aantal"
        Add-Output "`nBrinnummers : $brinnrs "
        Add-Output "Examendagen : $alle_dagen"
        Add-Output "Vakcodes : $vak_codes"
        Add-Output "GroepsIdImport : $grpidimp"
        # Add-Output " "

        # Nu de knoppen laten zien
        $btnBack = New-Object System.Windows.Forms.Button
        $btnBack.Text = "Stoppen"
        $btnBack.Location = '10,210'
        $btnBack.Size = '150,30'
        $btnBack.BackColor = "White"
        $btnBack.DialogResult = [System.Windows.Forms.DialogResult]::cancel
        $ShowExcelForm.Controls.Add($btnBack)

        $btnReopen = New-Object System.Windows.Forms.Button
        $btnReopen.Text = "Opieuw kiezen"
        $btnReopen.Location = '170,210'
        $btnReopen.Size = '150,30'
        $btnReopen.BackColor = "White"
        $btnReopen.DialogResult = [System.Windows.Forms.DialogResult]::No
        $ShowExcelForm.Controls.Add($btnReopen)

        $btnNext = New-Object System.Windows.Forms.Button
        $btnNext.Text = "Bevestigen"
        $btnNext.Location = '330,210'
        $btnNext.Size = '150,30'
        $btnNext.BackColor = "White"
        $btnNext.DialogResult = [System.Windows.Forms.DialogResult]::ok
        $ShowExcelForm.Controls.Add($btnNext)
    }) # Einde Form Load event

    $result = $ShowExcelForm.ShowDialog()

    $ShowExcelForm.Close()

     if ($result -eq [system.windows.forms.dialogResult]::ok) { 
         Show-ConvertForm
     } elseif ($result -eq [system.windows.forms.dialogResult]::No) {
                Write-Host "Excelbestand opnieuw openen."
                $Geselecteerd = SelectExcelForm
                # terug naar het vorige scherm. Als $geselteerd = GEEN dan wordt bij Show-ExcelForm hier rekening mee gehouden 
                Show-ExcelForm -ExcelPath $Geselecteerd
                
            } 
} # einde Show-ExcelForm

function Show-ConvertForm {
    # param([array]$Exceldata)

    $convertForm = declareren_standaardvenster -titel "Overzicht en opties wijzigen" -size_x 600 -size_y 380

    # geselecteerde Exceldata uit array halen
    $SelectedFile = Split-Path -Path $Exceldata.geselecteert -Leaf
    $Selectedlocation = Split-Path -Path $Exceldata.geselecteert -Parent
    $outputname = $Exceldata.uitvoernaam
    $outputFolder = $Exceldata.uitvoermap

    $lblSelected = New-Object System.Windows.Forms.Label
    $lblSelected.Text = "Gekozen bestand: $SelectedFile"
    $lblSelected.Location = '10,10'
    $lblSelected.Size = '590,20'
    $convertForm.Controls.Add($lblSelected)

    $lbllocation = New-Object System.Windows.Forms.Label
    $lbllocation.Text = "Locatie: $Selectedlocation"
    $lbllocation.Location = '60,30'
    $lbllocation.Size = '540,40'
    $convertForm.Controls.Add($lbllocation)

    $standaardzoekmap = New-Object System.Windows.Forms.Checkbox 
    $standaardzoekmap.Location = New-Object System.Drawing.Point(60, 70)
    $standaardzoekmap.Size = New-Object System.Drawing.Size(540,30)
    $standaardzoekmap.Text = "Deze locatie gebruiken als standaard zoekmap voor Excelbestanden."
    # $wissennabackup.Font = 'Microsoft Sans Serif,11'
    # $wissennabackup.ForeColor = [System.Drawing.Color]::green
    $standaardzoekmap.checked = $true
    $convertForm.Controls.Add($standaardzoekmap)

    $lblPath = New-Object System.Windows.Forms.Label
    $lblPath.Text = "Uitvoermap:"
    $lblPath.Location = '10,110'
    $lblPath.Size = '90,30'
    $convertForm.Controls.Add($lblPath)

    $OutputPath = New-Object System.Windows.Forms.Label
    $OutputPath.Text = "$outputFolder"
    $OutputPath.Location = '100,110'
    $OutputPath.Size = '590,40'
    $convertForm.Controls.Add($OutputPath)

    $standaarduitvoermap = New-Object System.Windows.Forms.Checkbox 
    $standaarduitvoermap.Location = New-Object System.Drawing.Point(60, 150)
    $standaarduitvoermap.Size = New-Object System.Drawing.Size(540,30)
    $standaarduitvoermap.Text = "Deze uitvoermap gebruiken als standaard uitvoermap."
    # $wissennabackup.Font = 'Microsoft Sans Serif,11'
    # $wissennabackup.ForeColor = [System.Drawing.Color]::green
    $standaarduitvoermap.checked = $true
    $convertForm.Controls.Add($standaarduitvoermap)

    $lblOutput = New-Object System.Windows.Forms.Label
    $lblOutput.Text = "Uitvoernaam:"
    $lblOutput.Location = '10,190'
    $lblOutput.Size = '100,20'
    $convertForm.Controls.Add($lblOutput)

    $txtOutput = New-Object System.Windows.Forms.TextBox
    $txtOutput.Location = '120,190'
    $txtOutput.Size = '300,20'
    $txtOutput.Text = $outputname
    $convertForm.Controls.Add($txtOutput)

    $standaarduitvoernaam = New-Object System.Windows.Forms.Checkbox 
    $standaarduitvoernaam.Location = New-Object System.Drawing.Point(60, 210)
    $standaarduitvoernaam.Size = New-Object System.Drawing.Size(540,50)
    $standaarduitvoernaam.Text = "Deze format gebruiken als standaard naam voor het omgezette bestand."
    # $wissennabackup.Font = 'Microsoft Sans Serif,11'
    # $wissennabackup.ForeColor = [System.Drawing.Color]::green
    $standaarduitvoernaam.checked = $true
    $convertForm.Controls.Add($standaarduitvoernaam)

    $btnSelect = New-Object System.Windows.Forms.Button
    $btnSelect.Text = "Wijzig uitvoermap"
    $btnSelect.Location = '380,300'
    $btnSelect.Size = '200,30'
    $btnSelect.BackColor = "White"
    $btnSelect.Add_Click({
        # Outputmap selecteren
        $FolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
        $FolderBrowser.Description = "Selecteer de uitvoermap"
        # $FolderBrowser.RootFolder = [System.Environment+SpecialFolder]::MyDocuments

        if ($FolderBrowser.ShowDialog() -eq "OK") {
                $outputFolder = $FolderBrowser.SelectedPath
                $OutputPath.Text = "$outputFolder"
                Write-Host "Andere uitvoermap gekozen: `n$outputFolder"
        }    
    })
    $convertForm.Controls.Add($btnSelect)

    $btnBack = New-Object System.Windows.Forms.Button
    $btnBack.Text = "Stoppen"
    $btnBack.Location = '10,300'
    $btnBack.Size = '150,30'
    $btnBack.BackColor = "White"
    $btnBack.DialogResult = [System.Windows.Forms.DialogResult]::cancel
    $convertForm.Controls.Add($btnBack)

    $btnNext = New-Object System.Windows.Forms.Button
    $btnNext.Text = "Omzetten"
    $btnNext.Location = '180,300'
    $btnNext.Size = '150,30'
    $btnNext.BackColor = "White"
    # $btnNext.DialogResult = [System.Windows.Forms.DialogResult]::ok
    $btnNext.Add_Click({
        # Controleer of uitvoernaam is ingevuld
        if ([string]::IsNullOrWhiteSpace($txtOutput.Text)) {
            [System.Windows.Forms.MessageBox]::Show("Voer een uitvoernaam in.", "Foutmelding", "OK", "Error")
            return
        }
        # Controleer of uitvoermap is ingevuld
        if ([string]::IsNullOrWhiteSpace($OutputPath.Text)) {
            [System.Windows.Forms.MessageBox]::Show("Selecteer een uitvoermap.", "Foutmelding", "OK", "Error")
            return
        }
        # In de titel aangeven dat de data wordt omgezet
        $convertForm.Text = "Data wordt omgezet naar Facet formaat..."
        # Data weer in array zetten en naar procedure voor het omzetten
        $Exceldata.uitvoernaam = $txtOutput.Text
        $Exceldata.uitvoermap = $OutputPath.Text
        # en alleen outputbox zichtbaar
        $outputBox.Visible = $true
        
        # de knop om te sluiten is even niet klikbaar
        
        # alle andere knoppen en text onzichtbaar maken
        $btnBack.Visible = $false
        $btnSelect.Visible = $false
        $btnNext.Visible = $false
        $lblSelected.Visible = $false
        $lbllocation.Visible = $false
        $lblPath.Visible = $false
        $lblOutput.Visible = $false
        $txtOutput.Visible = $false
        $OutputPath.Visible = $false
        $standaardzoekmap.Visible = $false
        $standaarduitvoermap.Visible = $false
        $standaarduitvoernaam.Visible = $false
        # tijd geven om dit uit te voeren
        Start-Sleep -Milliseconds 100

        # Starten omzetten
        Convert-ExcelToFacet
        # in de titel aangeven dat de data is omgezet
        $convertForm.Text = "KLAAR! Data is omgezet naar Facet formaat."
        # de knop om te sluiten is weer klikbaar en tekst is veranderd
        $btnBack.Text = "Sluiten"
        # $btnBack.Location = '10,220'
        $btnBack.Visible = $true

    })
    $convertForm.Controls.Add($btnNext)

    # TextBox voor statusoutput
    $outputBox = New-Object System.Windows.Forms.TextBox
    $outputBox.Multiline = $true
    $outputBox.ScrollBars = "Vertical"
    $outputBox.Size = New-Object System.Drawing.Size(560, 280)
    $outputBox.Location = New-Object System.Drawing.Point(10, 5)
    $outputBox.ReadOnly = $true
    $outputBox.Visible = $false # niet zichtbaar in het begin
    $convertForm.Controls.Add($outputBox)

    $null = $convertForm.ShowDialog()

    $convertForm.Close()
} # einde Show-ConvertForm

function Convert-ExcelToFacet {
    
    # Dit is de logica om de Excel data om te zetten naar Facet formaat
   
    $BaseOutputFolder = Join-Path $PSScriptRoot "Output"
    $FinalZip = Join-Path $Exceldata.uitvoermap $Exceldata.uitvoernaam
    $FinalZip = -join ($FinalZip,'.zip')
    $zips = @()

    Write-Host "Omzetten van Excel data naar Facet formaat..."
    Add-Output "Omzetten van Excel data naar Facet formaat..."
    # Tekst naar console en logbestand. Met `n wordt een nieuwe regel toegevoegd. Met `t wordt een tab toegevoegd.
    # De 1e regel is een lege regel, ook zonder datum en tijd.
    Write-Log -Message " " -Notimestamp
    Write-Log -Message "Omzetten van Excel data naar Facet formaat is gestart.
    `tGekozen bestand: $($Exceldata.geselecteert)
    `tAantal ingelezen rijen: $($Exceldata.aantal)
    `tUitvoermap: $($Exceldata.uitvoermap)
    `tUitvoernaam: $($Exceldata.uitvoernaam)"
    
    # Controleer of de uitvoermap bestaat en verwijderen als dit zo is. Vervolgens maak deze aan.
    if (Test-Path $BaseOutputFolder) { Remove-Item $BaseOutputFolder -Recurse -Force }
    New-Item -ItemType Directory -Path $BaseOutputFolder | Out-Null

    # Alle groepen sorteren uit de Exceldata op mapnaam
    $groepen = $Exceldata.inhoud | Group-Object -Property MapNaam
    foreach ($groep in $groepen) {
        $groepFolder = Join-Path $BaseOutputFolder $groep.Name
        New-Item -ItemType Directory -Path $groepFolder | Out-Null

        # Alle subgroepen sorteren uit de Exceldata op GroepsIdImport
        $subgroepen = $groep.Group | Group-Object -Property GroepsIdImport

        foreach ($subgroep in $subgroepen) {

            # voor aangepaste examens wordt een aparte xml gemaakt.
            # kolom 14 bepaald of het een aangepaste examen betreft.
            $aangepasten = $subgroep.Group | Group-Object -Property kolom14

            foreach ($subaangepast in $aangepasten) {
            $xmlPath = Join-Path $groepFolder ("$($subgroep.Name).xml")
            # aangepast examen - xml bestand
            $xmlaangepast = Join-Path $groepFolder ("$($subgroep.Name)_dos.xml")
            # in kolom 14 wordt aangegeven of kandidaat een aangepaste examen aflegt - dan wordt aparte xml aangemaakt.
            $aangepastexamen = $subaangepast.group.kolom14
            
            if ([string]::IsNullOrWhiteSpace($aangepastexamen)) {
                New-XmlFile -DataRows $subaangepast.Group -OutputPath $xmlPath | Out-Null
                Write-Log -Message "`tXML gemaakt: $($xmlPath)" -Notimestamp
                Write-Host "XML gemaakt: $($xmlPath)"
                Add-Output "XML gemaakt: $($xmlPath)"
            } else {
                New-XmlFile -DataRows $subaangepast.Group -OutputPath $xmlaangepast | Out-Null
                Write-Log -Message "`tXML gemaakt: $($xmlaangepast)" -Notimestamp
                Write-Host "XML gemaakt: $($xmlaangepast)"
                Add-Output "XML gemaakt: $($xmlaangepast)"
            }
            }
        }
        # Nu de xml bestanden zijn aangemaakt, kan de zip bestand worden gemaakt.
        $zipPath = "$groepFolder.zip"
        Compress-Folder -FolderPath $groepFolder -ZipFilePath $zipPath
        Write-Log -Message "`tGroepsmap aangemaakt : $groepFolder" -Notimestamp
        Write-Host "Groepsmap aangemaakt : $groepFolder"
        Add-Output "Groepsmap aangemaakt : $groepFolder"
        # Toevoegen zip aan array met groepsmappen/ brinnummers
        $zips += $zipPath

        Remove-Item $groepFolder -Recurse -Force
        # Write-Log -Message "Map verwijderd: $groepFolder"
        
    }

    # tijdelijke map voor gezipte bestanden
    $mergeFolder = Join-Path $PSScriptRoot "MergedTemp"
    # deze verwijderen en opnieuw aanmaken. Dan is het zeker leeg.
    if (Test-Path $mergeFolder) { Remove-Item $mergeFolder -Recurse -Force }
    New-Item -ItemType Directory -Path $mergeFolder | Out-Null

    # Nu de zip bestanden in de tijdelijke map zetten
    foreach ($z in $zips) {
        Copy-Item $z -Destination $mergeFolder
    }

    # Nu de zip bestanden samenvoegen in 1 zip bestand. Dit is het uiteindelijke bestand.
    Compress-Folder -FolderPath $mergeFolder -ZipFilePath $FinalZip
    Write-Log -Message "`tBestand aangemaakt: $FinalZip" -Notimestamp
    Write-Host "Bestand aangemaakt: $FinalZip" -ForegroundColor Green
    Add-Output -Message "Bestand aangemaakt: $FinalZip"

    # tijdelijke bestanden en mappen verwijderen
    Remove-Item $mergeFolder -Recurse -Force
    Remove-Item $BaseOutputFolder -Recurse -Force
    
    Write-Log -Message "Klaar!" 
    Write-Host "Klaar!" -ForegroundColor Green
    
    Add-Output "Bestand aangemaakt: $FinalZip"
    Add-Output " "
    Add-Output "===== Klaar! ====="

} # einde Convert-ExcelToFacet

function New-XmlFile {
    # Hier wordt de xml aangemaakt adhv de gegeven $datarows en bewaard in $outputpath

    param([array]$DataRows, [string]$OutputPath)

    # Eenmalig Root element van de XML aanmaken
    [xml]$xmlDoc = New-Object System.Xml.XmlDocument
    $root = $xmlDoc.CreateElement("vastleggenAfnamegroep_V2")
    $root.SetAttribute("xmlns", "http://duo.nl/schema/DUO_Examengegevens_V1.0") | Out-Null   # Out-Null zorgt ervoor dat er geen tekst op de console komt
    $root.SetAttribute("xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance") | Out-Null
    $xmlDoc.AppendChild($root)

    # toevoegen van een extra letter aan de afnameleider en planner. A voor Albeda en een Z voor Zadkine.
    if ($DataRows[0].mapnaam -eq "00GT") {
        $extraletter = 'A'
        } elseif ($DataRows[0].mapnaam -eq "25LP") {
        $extraletter = 'Z'
        } else {
        # hier aangekomen betekent dat de brinnr niet correct is. De letter blijft leeg.
        $extraletter = ''
        }

    # Eenmalig <afnamegroep> element aanmaken
    $afnamegroep = $xmlDoc.CreateElement("afnamegroep")
    $afnamegroep.AppendChild($xmlDoc.CreateElement("groeps_id_import")).InnerText = $DataRows[0].GroepsIdImport
    $afnamegroep.AppendChild($xmlDoc.CreateElement("bRIN")).InnerText = $DataRows[0].mapnaam
    $afnamegroep.AppendChild($xmlDoc.CreateElement("bRINvolgnummer")).InnerText = $DataRows[0].kolom3
    $afnamegroep.AppendChild($xmlDoc.CreateElement("naamexamengroep")).InnerText = -join ($DataRows[0].kolom4,' ',$DataRows[0].kolom5)
    $afnamegroep.AppendChild($xmlDoc.CreateElement("naamexamenlokaal")).InnerText = $DataRows[0].kolom5
    $afnamegroep.AppendChild($xmlDoc.CreateElement("examenafnamedatum")).InnerText = $DataRows[0].kolom6
    $afnamegroep.AppendChild($xmlDoc.CreateElement("examenbegintijd")).InnerText = $DataRows[0].kolom7
    $afnamegroep.AppendChild($xmlDoc.CreateElement("online_indicatie")).InnerText = $DataRows[0].kolom8
    $afnamegroep.AppendChild($xmlDoc.CreateElement("inlognaamafnameplanner")).InnerText = -join ($DataRows[0].kolom9,$extraletter)
    $afnamegroep.AppendChild($xmlDoc.CreateElement("inlognaamafnameleider")).InnerText = -join ($DataRows[0].kolom10,$extraletter)
    
    # Eenmalig <examen>
    $examen = $xmlDoc.CreateElement("examen")
    $examen.AppendChild($xmlDoc.CreateElement("toetsgebruik")).InnerText = $DataRows[0].kolom11
    $examen.AppendChild($xmlDoc.CreateElement("vakcode")).InnerText = $DataRows[0].kolom12
    $examen.AppendChild($xmlDoc.CreateElement("opleidingniveau")).InnerText = $DataRows[0].kolom13
    # Als kolom kandidaatbijzonderheid niet leeg is...
    $bijzonderheid = $DataRows[0].kolom14 
    if (![string]::IsNullOrWhiteSpace($bijzonderheid)) {
        $examen.AppendChild($xmlDoc.CreateElement("kandidaatbijzonderheid")).InnerText = $bijzonderheid
    }
    $afnamegroep.AppendChild($examen)

    # Per kandidaat een <afname> element aanmaken
    foreach ($row in $DataRows) {
        $afname = $xmlDoc.CreateElement("afname")

        $afname.AppendChild($xmlDoc.CreateElement("burgerservicenummer")).InnerText = $Row.kolom15
        $afname.AppendChild($xmlDoc.CreateElement("inlognaam")).InnerText = $Row.kolom16
        $afname.AppendChild($xmlDoc.CreateElement("verklankingIndicatie")).InnerText = $Row.kolom17
        $afname.AppendChild($xmlDoc.CreateElement("geboortedatum")).InnerText = $Row.kolom18
        $afname.AppendChild($xmlDoc.CreateElement("voornamen")).InnerText = $Row.kolom19

        $achternaam = $Row.kolom20
        if ($achternaam -like "*,*") {
            $split = $achternaam -split ","
            $afname.AppendChild($xmlDoc.CreateElement("voorvoegsels")).InnerText = $split[1].Trim()
            $afname.AppendChild($xmlDoc.CreateElement("geslachtsnaam")).InnerText = $split[0].Trim()
            
        } else {
            $afname.AppendChild($xmlDoc.CreateElement("geslachtsnaam")).InnerText = $achternaam
        }

        $afname.AppendChild($xmlDoc.CreateElement("geslacht")).InnerText = $Row.kolom21
        $afname.AppendChild($xmlDoc.CreateElement("leerlingnummer")).InnerText = $Row.kolom22
        $afname.AppendChild($xmlDoc.CreateElement("kenmerk1")).InnerText = $Row.kolom23
        $afname.AppendChild($xmlDoc.CreateElement("kenmerk2")).InnerText = $Row.kolom24
        $afname.AppendChild($xmlDoc.CreateElement("kenmerk3")).InnerText = $Row.kolom25
        $afname.AppendChild($xmlDoc.CreateElement("wolfgroep")).InnerText = $Row.kolom26

        $afnamegroep.AppendChild($afname) | Out-Null
    }

    # afnamegroep afsluiten en toevoegen --- Nog toevoegen1 en de andere hiervoor verwijderen!
    $afnamegroep.AppendChild($xmlDoc.CreateElement("aantalvarianten")).InnerText = $DataRows[0].kolom27
    $root.AppendChild($afnamegroep) | Out-Null

    $xmlDoc.Save($OutputPath)
}

function Compress-Folder {
    # Hier wordt een zipbestand gemaakt adhv de 2 gegeven bestanden

    param([string]$FolderPath, [string]$ZipFilePath)

    if (Test-Path $ZipFilePath) {
        Remove-Item $ZipFilePath -Force
    }
    [System.IO.Compression.ZipFile]::CreateFromDirectory($FolderPath, $ZipFilePath)
}

Function Search-Update {
    # Controleer of er een update is voor dit script
    Write-Host "Controleren op een update."

    # Declareren venster met standaard waarden
    $SearchUpdateForm = declareren_standaardvenster -titel "Controleren op een update." -size_x 600 -size_y 300

    # TextBox voor statusoutput
    $outputBox = New-Object System.Windows.Forms.TextBox
    $outputBox.Multiline = $true
    $outputBox.ScrollBars = "Vertical"
    $outputBox.Size = New-Object System.Drawing.Size(560, 200)
    $outputBox.Location = New-Object System.Drawing.Point(10, 5)
    $outputBox.ReadOnly = $true
    $SearchUpdateForm.Controls.Add($outputBox)

    $SearchUpdateForm.Show()
    # nodig om proces de tijd te geven om de tekst te laten zien.
    Start-Sleep -Milliseconds 500  

    # Vanaf hier start de controle op een update
    Add-Output "Controleer op een update van dit script."

    $tijdelijkepad = "$PSScriptRoot\temp\" # dit is de tijdelijke map waar de update-bestanden worden opgeslagen
    # Als de tijdelijke map bestaat wordt deze geleegd en het update niet uitgevoerd.
    if (Test-Path $tijdelijkepad) {
        write-host "Er is al een update uitgevoerd."
        Add-Output "Er is al een update uitgevoerd."
        # Even wachten tot de gebruiker de tekst heeft gelezen.
        Start-Sleep -Seconds 5
        Remove-Item $tijdelijkepad -Recurse -Force
        # sluiten van venster.
        $SearchUpdateForm.dispose()
        return
    }

    $updateto = "0.0.0" # standaard waarde. Als er geen update is, dan blijft deze waarde 0.0.0
    $huidigeversie = $programma.versie # huidige versie van het programma

    $url = $programma.github # dit is de url van de github repository 
    # inhoud van een map in github ophalen
    $response = Invoke-RestMethod -Uri $url -Headers @{ "User-Agent" = "PowerShell" }
    
    # Loop door de items in de response en controleer of er een nieuw bestand is
    foreach ($item in $response) {
        
        if ($item.type -eq "file") {
            # Download het bestand

            $localFile = -join ($tijdelijkepad,$item.name)
            if (!(Test-Path $localFile)) {
                # Maak de map aan als deze nog niet bestaat
                New-Item -ItemType Directory -Path (Split-Path $localFile) -Force | Out-Null
            }

            # "Bestand $($item.name) downloaden naar de map latest."
            Invoke-WebRequest -Uri $item.download_url -OutFile $localFile
            
            # bepaal laatste versie van het script
            # bestnaam is de naam van het gedownloade bestand
            $bestnaam = $($item.name) 
            # versiemetzip is de versie met .zip in de naam.
            $versiemetzip = $bestnaam.split('_')[1]
            # de positie van de laatste punt in de versie bepalen
            $positiepunt = $versiemetzip.LastIndexOf(".")
            # updateto is alleen de versie zonder .zip. Dit is de versie die we willen vergelijken met de huidige versie.
            $updateto = $versiemetzip.Substring(0, $positiepunt) 
            # gedownloadebestand is de naam van het bestand dat is gedownload
            $gedownloadebestand = $localFile
        }
    }

    if ($updateto -eq "0.0.0") {
        Add-Output "Er is geen update gevonden in de GitHub repository: $url"
        } elseif ($huidigeversie -ge $updateto) {
        Add-Output "Het programma heeft de laatste update."
        } else {
        Add-Output "Er is een update beschikbaar voor dit script."
    
        Add-Output "De huidige versie is $huidigeversie en de laatste versie is $updateto."
        Add-Output "Het script wordt nu bijgewerkt."    
        
        # Nu het script updaten via de function Update-Script
        Update-Script $gedownloadebestand
        }
 
     # tijdelijkemap legen
    if (Test-Path $tijdelijkepad) {
        Remove-Item $tijdelijkepad -Recurse -Force
    }
    
    # sluiten van venster. Dit moet na het downloaden van de bestanden.
    $SearchUpdateForm.dispose()
} # einde Search-Update

Function Update-Script {
    # Deze functie moet de bestanden van Github downloaden en de inhoud van het script vervangen.
    # Deze functie wordt aangeroepen vanuit Search-Update.
    param (
        [string]$updateto
    )
    
    $startmap = "$PSScriptRoot" # dit is de map waar de bestanden worden uitgepakt
    # Expand-Archive -Path "$updateto" -DestinationPath "$startmap" -Force

    # Nu het script opnieuw starten
    Write-Host "Het script wordt opnieuw gestart in $startmap"
    Add-Output "Het script wordt opnieuw gestart in $startmap"
 
    # Let op dat de tijdelijke map niet meer wordt geleegd. Zie functie Search-Update.

    # Even wachten tot de bestanden zijn uitgepakt en de gebruker de tijd heeft om de tekst te lezen.    
    Start-Sleep -Seconds 5

    # sluiten van venster. Dit moet na het downloaden van de bestanden.
    $SearchUpdateForm.dispose()

    # Nu het nieuwe script starten
    powershell -file "$PSScriptRoot\cetool.ps1"

    # beëindigen van programma als updaten is uitgevoerd. Anders kan je na een update niet afsluiten.
    exit;
} # einde Update-Script


# Deze functie toont het hoofdmenu van de applicatie
function Show-MainForm {
    $mainForm = declareren_standaardvenster -titel "CE-tool omzetten Excel" -size_x 400 -size_y 300

    # Icoontjes laden
    $startIcon = Get-PNGImage "$PSScriptRoot\start.png"
    $logIcon   = Get-PNGImage "$PSScriptRoot\logboek.png"
    $exitIcon  = Get-PNGImage "$PSScriptRoot\stoppen.png"

    # Start Omzetten knop
    $btnStart = New-Object System.Windows.Forms.Button
    $btnStart.Text = "Start Omzetten"
    $btnStart.Size = '350,50'
    $btnStart.Location = '20,20'
    $btnStart.Image = $startIcon
    $btnStart.ImageAlign = "MiddleLeft"
    $btnStart.BackColor = "White"
    $btnStart.Add_Click({
        $mainForm.Hide() # Zorg dat hoofdmenu sluit
        $Geselecteerd = SelectExcelForm
        if ($Geselecteerd -ne 'GEEN') {
            # Import-ExcelFile -ExcelPath $Geselecteerd
            # Exceldata laten zien
            Show-ExcelForm -ExcelPath $Geselecteerd
        }
        $mainForm.Show()
    })
    $mainForm.Controls.Add($btnStart)

    # Logboek bekijken knop
    $btnLog = New-Object System.Windows.Forms.Button
    $btnLog.Text = "Logboek Bekijken"
    $btnLog.Size = '350,50'
    $btnLog.Location = '20,90'
    $btnLog.Image = $logIcon
    $btnLog.ImageAlign = "MiddleLeft"
    $btnLog.BackColor = "White"
    $btnLog.Add_Click({
        $mainForm.Hide() # Zorg dat hoofdmenu sluit
        Show-Logboek
        $mainForm.Show()
    })
    $mainForm.Controls.Add($btnLog)

    # Afsluiten knop
    $btnExit = New-Object System.Windows.Forms.Button
    $btnExit.Text = "Afsluiten"
    $btnExit.Size = '350,50'
    $btnExit.Location = '20,160'
    $btnExit.Image = $exitIcon
    $btnExit.ImageAlign = "MiddleLeft"
    $btnExit.BackColor = "White"
    $btnExit.Add_Click({
        $mainForm.Close()
    })
    $mainForm.Controls.Add($btnExit)

    # zorgen dat deze venster altijd bovenop komt
    $mainForm.TopMost = $true

    $null = $mainForm.ShowDialog()
} # einde Show-MainForm


############ start script ###############

# tijdelijk uitgeschakeld
Search-Update;

# Lees de instellingen van de gebruiker in
# Dit is de functie die de instellingen van de gebruiker leest en teruggeeft als een object
$gebruiker = ReadSettings

Show-MainForm;

<# 
#dit is een test

$gebr1 = gebruikersinstellingen

# SaveSettings $gebr

$gebr1

$gebr2 = ReadSettings

"========= new================"

$gebr2

SaveSettings $gebr1
pause

 #>