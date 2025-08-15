# test om een standaard uitvoermap in te stellen

$invoer = "[computer] [gebruiker] [datum]-[tijd] - [uren]:[minuten]:[seconden]"
$uitvoer = $invoer
do {
    
    $gevonden = $uitvoer.IndexOf('[')
    if ($gevonden -ge 0) {
        # eerste gevonden
        $gevonden2 = $uitvoer.IndexOf(']') 
        if ($gevonden2 -ge 0) {
            # het gevonden woord die automatisch wordt ingevuld
            $woord = $uitvoer.Substring($gevonden,$gevonden2-$gevonden+1)
            # afhankelijk van het woord, de nieuwe waarde ophalen
            switch ($woord) {
                "[jaar]" {
                $nieuw = get-date -Format "yyyy"
                         }
                "[maand]" {
                $nieuw = get-date -Format "MM"
                         }
                "[dag]" {
                $nieuw = get-date -Format "dd"
                         }
                "[datum]" {
                $nieuw = get-date -Format "yyMMdd"
                         }
                "[tijd]" {
                $nieuw = get-date -Format "HHmmss"
                         }
                "[minuten]" {
                $nieuw = get-date -Format "mm"
                         }
                "[uren]" {
                $nieuw = get-date -Format "HH"
                         }
                "[seconden]" {
                $nieuw = get-date -Format "ss"
                         }
                "[gebruiker]" {
                $nieuw = $env:USERNAME
                         }
                "[computer]" {
                $nieuw = $env:ComputerNAME
                         }
                default {
                # het woord tussen de haakjes is niet herkend en wordt terug gegeven zonder de haakjes.
                $nieuw = $woord.Substring(1,$woord.length-2)
                }
            } # einde switch
            # Het gevonden woord wordt vervangen door wat het representeert.
            $uitvoer = $uitvoer.Replace($woord,$nieuw)

        } else {
        # als de eerste teken [ is gevonden maar de 2e teken ] niet, dan stoppen. dus moet gevonden naar -1
        $gevonden = -1
        }
    }
} while ($gevonden -ge 0)

write-host "uit " $uitvoer
write-host "in  " $invoer