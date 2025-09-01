# Parameters
$logPad = "script_log_test.txt"
$dagenBehouden = 9
$grensDatum = (Get-Date).AddDays(-$dagenBehouden)

# controleer of het logbestand bestaat
if (-Not (Test-Path $logPad)) {
    Write-Host "Logbestand bestaat niet: $logPad"
    exit 1
}  

# Buffer voor de regels die we willen behouden
$regelsOmTeBehouden = @()

# Lees het logbestand regel voor regel
foreach ($regel in Get-Content $logPad) {
    if ($regel -match '^(\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2})') {
        $regelDatum = [datetime]::ParseExact($matches[1], 'yyyy-MM-dd HH:mm:ss', $null)
        
        if ($regelDatum -lt $grensDatum) {
            break  # Stop als deze regel ouder is dan de grens
        }
    }

    # Voeg toe aan lijst zolang grens niet overschreden is
    $regelsOmTeBehouden += $regel
}

# Overschrijf het originele bestand met de behouden regels
$regelsOmTeBehouden | Set-Content $logPad
