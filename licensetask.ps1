<#
        Most of this program is copy/pasted from the other one. The syncronized hashtable is one example of something that I had set up on the other one that doesn't really need to be here,
    but it was more work than I felt necessary to change it since it still works and doesn't hurt anything.

        For more in-depth comments, see the other script.
#>

$sync = [Hashtable]::Synchronized(@{})


$influxIP = ""                    # IP Address of the influxDB server
$sync.config = [PSCustomObject]@{
    PATH = "c:\lmutil"            # Path to LMUTIL program
    SERVER = "27000@license"      # Port@hostname or port@IP of the license server
}



$save = ""
$sync.years = ((c:\windows\system32\cmd.exe /c ($sync.config.PATH + "\lmutil.exe") lmstat -c $sync.config.server -a | select-string "_[0-9][0-9][0-9][0-9]_").matches.Value -replace "_","" | sort-object -unique) | select-object -last 2
foreach ($line in $sync.years) { 
    $save = $save + "_" +  $line 
}
if ((Test-Path ($sync.config.PATH + "\autodesk_names" + $save + ".csv")) -eq $false) {
    $resulta = foreach ($line in $sync.years) {
        $webrequest = invoke-webrequest https://knowledge.autodesk.com/customer-service/network-license-administration/managing-network-licenses/interpreting-your-license-file/feature-codes/$line-flexnet-feature-codes 
        $tables = @($WebRequest.ParsedHtml.getElementsByTagName("TABLE"))
        $table = $tables[0]
        $titles = @()
        $rows = @($table.Rows)
        foreach($row in $rows){
            $cells = @($row.Cells)
            $titles = @("Product","Feature","Package","Term")
            $resultObject = [Ordered] @{}
            for($counter = 0; $counter -lt $cells.Count; $counter++) {
                $title = $titles[$counter]
                if(-not $title) { continue }
                if (-not ([string]::IsNullOrWhiteSpace($cells[$counter].InnerText))) {
                    $resultObject[$title] = ("" + $cells[$counter].InnerText).Trim()
                } else {
                    $resultObject[$title] = ("plots").Trim()
                }
            }
            [PSCustomObject] $resultObject | where-object { $_.product -ne "Product Name" }
        }
    } 
    $resultb = import-csv ($sync.config.PATH + "\extras.csv")
    $sync.auto = $resulta + $resultb
    $sync.auto | export-csv ($sync.config.PATH + "\autodesk_names" + $save + ".csv") -NoTypeInformation
}
$sync.auto = import-csv ($sync.config.PATH + "\autodesk_names" + $save + ".csv")
$desk = c:\windows\system32\cmd.exe /c ($sync.config.PATH + "\lmutil.exe") lmstat -c $sync.config.server -a
$desk | select-object -skip 8 | where-object { $_ -notmatch '^  "' } | where-object { $_ -notmatch "vendor_string" } | 
    where-object { $_ -notmatch "floating license" } | foreach-object { $_ -replace '(?<=\()(.*)(?=start )', ') ' -replace '\(', '' -replace '\)', '' -replace 'Users of ',''} | set-variable desk
$desk | where-object { -not [String]::IsNullOrWhiteSpace($_) } | set-variable desk
$h = 0
$sync.result = @()
$sync.result.clear()
do { 
    foreach ($line2 in $desk) {
        if ( ($line2 -match $sync.auto[$h].feature) -or (($line2 -match $sync.auto[$h].term) -and ($line2 -match "pdcoll")) ) {
            $inuse = [int]($line2 | where-object { $_ -match "[0-9]?[0-9]?[0-9]?[0-9]? licenses? in use" } | ForEach-Object { ($matches[0] | select-string -pattern "[0-9]?[0-9]?[0-9]?[0-9]?").matches.value })
            $issued = [int]($line2 | where-object { $_ -match "[0-9]?[0-9]?[0-9]?[0-9]? licenses? issued" } | ForEach-Object { ($matches[0] | select-string -pattern "[0-9]?[0-9]?[0-9]?[0-9]?").matches.value })
            $Users = if ( ($inuse -gt 0) ) { (([string]($desk | select-string $line2 -list -quiet -context 0,($inuse)).context.PostContext) -split "   ") -replace "\s+"," " };
            if ($inuse -gt 0) {
                $temp2 = foreach ( $user in ($users | select-object -skip 1) ) {
                    [PSCustomObject]@{
                        User = $user.trim().split(' ')[0]
                        Host = $user.trim().split(' ')[1]
                        Time = (get-date -format yyyy) + "/" + $user.trim().split(' ')[5] + " " + $user.trim().split(' ')[6]
                    }
                }
            } else { $temp2 = "" }
            $temp = [pscustomobject]@{
                 Product = $sync.auto[$h].product
                 Feature = if ($sync.auto[$h].feature -ne "plots") {$sync.auto[$h].feature} else { "" }
                 Used = $inuse
                 Issued = $issued
                 Users = $temp2.user
                 Time = $temp2.time
            }
            $sync.result += $temp
        }
    }
    $h++
} until ( $h -ge $sync.auto.count )  

foreach ($line in $sync.result) {
    $user = ($line.users).count
    $users = $line.users
    $i = 0
    foreach ($use in $line.users) {
        write-influx -measure lmstat -tags @{Product=$line.product;User=$use} -metrics @{Product=$line.product;Date=$line.time[$i];Issued=$line.issued;Used=$line.used;Users=$use} -database lmstat -server $influxIP
        $i++
    }
}