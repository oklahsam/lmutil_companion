<#
    I made this program to make it easier to see the status of any license server running LMTOOLS (primarily Autodesk license servers). 
    The LMTOOLS GUI gives you the tiniest little text window to read the status. You can also pop it out into a text document, but it's not much better since it's a giant wall of text with
        mostly useless information.
    This can be run from any computer that has network access to the license server, so long as you grab a copy of the LMUTIL.EXE from the server.

    There are some stability issues if you compile it into an EXE. I haven't been able to track them all down, but they seem to be related to the datagridviews running in separate threads.... I think.
    The crashes don't seem to happen when running it as a regular Powershell script, though.
#>


Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

# for talking across runspaces.
$sync = [Hashtable]::Synchronized(@{})

# This checks for a config file in your USERPROFILE folder. This file is used to store the "lmutil.exe" location and License Server "port@name".
if ( (test-path ($ENV:userprofile + "\LMconfig.xml")) -eq $false ) {
    $sync.config = [PSCustomObject]@{
        PATH = read-host "Input path to lmutil.exe (i.e.: c:\folder)"
        SERVER = read-host "Input server port@name (i.e.: 27000@license)"
    }
    $sync.config | export-clixml ($ENV:userprofile + "\LMconfig.xml")
} else {
    $sync.config = import-clixml ($ENV:userprofile + "\LMconfig.xml")
}

# Hide PowerShell Console
Add-Type -Name Window -Namespace Console -MemberDefinition '
[DllImport("Kernel32.dll")]
public static extern IntPtr GetConsoleWindow();
[DllImport("user32.dll")]
public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
'
$consolePtr = [Console.Window]::GetConsoleWindow()
[Console.Window]::ShowWindow($consolePtr, 0)

# This is set because the REFRESH function doesn't need to do a few things on its first run.
$sync.firstrun = $true

# Begin Primary form setup
$Form                            = New-Object system.Windows.Forms.Form
$Form.ClientSize                 = '1005,720'
$Form.text                       = "Autodesk License Status"
$Form.TopMost                    = $false
$form.maximizebox                = $false
$Form.formborderstyle            = 'Fixed3D'

$ListBox1                        = New-Object system.Windows.Forms.datagridview
$listbox1.ColumnCount = 3
$Listbox1.ColumnHeadersVisible = $True
$listbox1.columns[0].name = "Product"
$listbox1.columns[1].name = "Used"
$listbox1.columns[2].name = "Issued"
$listbox1.AutoSizeColumnsMode = "AllCells"
$listbox1.selectionmode = "FullRowSelect"
$Listbox1.MultiSelect = $false
$listbox1.ReadOnly = $true
$listbox1.AllowUserToOrderColumns = $false
$listbox1.AllowUserToResizeColumns = $false
$Listbox1.AllowUserToResizeRows = $false
$ListBox1.text                   = ""
$ListBox1.width                  = 510
$ListBox1.height                 = 640
$ListBox1.location               = New-Object System.Drawing.Point(10,10)
$listbox1.Font                   = 'Droid Sans Mono,10'
$listbox1.BorderStyle = "Fixed3d"

$ListView1                       = New-Object system.Windows.Forms.datagridview
$listview1.ColumnCount = 5
$ListView1.ColumnHeadersVisible = $True
$listview1.columns[0].name = "User"
$listview1.columns[1].name = "Host"
$ListView1.columns[2].name = "Display"
$listview1.columns[2].Visible = $false
$listview1.columns[3].name = "Borrowed"
$listview1.columns[4].name = "Days Left"
$listview1.AutoSizeColumnsMode = "AllCells"
$listview1.selectionmode = "FullRowSelect"
$ListView1.MultiSelect = $false
$listview1.readonly = $true
$listview1.AllowUserToOrderColumns = $false
$listview1.AllowUserToResizeColumns = $false
$ListView1.AllowUserToResizeRows = $false
$ListView1.text                  = ""
$ListView1.width                 = 460
$ListView1.height                = 640
$ListView1.location              = New-Object System.Drawing.Point(535,10)
$listview1.Font                  = 'Droid Sans Mono,10'
$listview1.BorderStyle = "Fixed3d"

$TextBox2                        = New-Object system.Windows.Forms.TextBox
$TextBox2.multiline              = $false
$TextBox2.width                  = 985
$TextBox2.height                 = 20
$TextBox2.location               = New-Object System.Drawing.Point(10,655)
$TextBox2.Font                   = 'Microsoft Sans Serif,10'
$Textbox2.ReadOnly               = $True

$Button2                         = New-Object system.Windows.Forms.Button
$Button2.text                    = "Refresh"
$Button2.width                   = 875
$Button2.height                  = 30
$Button2.location                = New-Object System.Drawing.Point(10,685)
$Button2.Font                    = 'Droid Sans Mono,10'

$Button1                         = New-Object system.Windows.Forms.Button
$Button1.text                    = "Kick"
$Button1.width                   = 50
$Button1.height                  = 30
$Button1.location                = New-Object System.Drawing.Point(945,685)
$Button1.Font                    = 'Droid Sans Mono,10'

$Button4                         = New-Object system.Windows.Forms.Button
$Button4.text                    = "Filter"
$Button4.width                   = 50
$Button4.height                  = 30
$Button4.location                = New-Object System.Drawing.Point(890,685)
$Button4.Font                    = 'Droid Sans Mono,10'


# This sets certain form controls to the "$sync" synchronized hash table. This lets me update them from separate runspaces, helping to keep the form responsive.
$sync.features = $ListBox1
$sync.users = $ListView1
$sync.start = $Button2
$sync.kick = $button1
$sync.progresstext = $TextBox2
$sync.filterbutton = $button4

$Form.controls.AddRange(@($sync.features,$sync.users,$sync.start,$sync.kick,$sync.progresstext,$sync.filterbutton))

# Begin Filter form setup
$Form2                            = New-Object system.Windows.Forms.Form
$Form2.ClientSize                 = '120,150'
$Form2.text                       = "Filter"
$Form2.TopMost                    = $false
$form2.maximizebox                = $false
$Form2.formborderstyle            = 'Fixed3D'
$form2.MinimizeBox                = $false

$ListBox2                        = New-Object system.Windows.Forms.listbox
$listbox2.SelectionMode          = "MultiExtended"
$ListBox2.text                   = ""
$ListBox2.width                  = 110
$ListBox2.height                 = 115
$ListBox2.location               = New-Object System.Drawing.Point(5,5)
$listbox2.Font                   = 'Droid Sans Mono,10'

$Button3                         = New-Object system.Windows.Forms.Button
$Button3.text                    = "Submit"
$Button3.width                   = 110
$Button3.height                  = 30
$Button3.location                = New-Object System.Drawing.Point(5,115)
$Button3.Font                    = 'Droid Sans Mono,10'

$sync.listbox2 = $listbox2


$form2.controls.addrange(@($listbox2,$button3))




# Refresh Button action
$sync.start.Add_Click({
    $script:refreshrunspace = [PowerShell]::Create().AddScript({
        invoke-expression $sync.refresh
        $script:refreshrunspace.runspace.dispose()
        $script:refreshrunspace.dispose()
    })
    $runspace = [RunspaceFactory]::CreateRunspace()
    $runspace.ApartmentState = "STA"
    $runspace.ThreadOptions = "ReuseThread"
    $runspace.Open()
    $runspace.SessionStateProxy.SetVariable("sync", $sync)
    $script:refreshrunspace.Runspace = $runspace
    $script:refreshrunspace.BeginInvoke()
})

# Kick Button action
$sync.kick.add_click({
    $script:refreshrunspace = [PowerShell]::Create().AddScript({
        $kickresult = c:\windows\system32\cmd.exe /c ($sync.config.PATH + "\lmutil.exe") lmremove -c $sync.config.server $sync.select.Feature $sync.select.user $sync.select.host $sync.select.Display
        if ( $kickresult -match "(-64,200)" ) { 
            $sync.progresstext.text = "Too many kicks. Please wait about 5 minutes and try again..." 
        } else {
            $sync.progresstext.text = "Kick request sent..." 
            invoke-expression $sync.refresh
        }   
        $script:refreshrunspace.runspace.dispose()
        $script:refreshrunspace.dispose()
    })
    $runspace = [RunspaceFactory]::CreateRunspace()
    $runspace.ApartmentState = "STA"
    $runspace.ThreadOptions = "ReuseThread"
    $runspace.Open()
    $runspace.SessionStateProxy.SetVariable("sync", $sync)
    $script:refreshrunspace.Runspace = $runspace
    $script:refreshrunspace.BeginInvoke()

})

# Loads user list when clicking on Product List
$sync.features.add_click({
    invoke-expression $sync.userlist
})

# Sets variables and enables kick button when a User is selected in the User List
$sync.users.add_click({
    $sync.index = $sync.result | where-object { $_.product -eq $sync.features.SelectedCells.value[0]}
    $sync.select = [pscustomobject]@{
        User = $sync.users.selectedcells.value[0]
        Host = $sync.users.selectedcells.value[1]
        Display = $sync.users.selectedcells.value[2]
        Feature = $sync.index.feature
    }
    $sync.kick.enabled = $true 
})

# Button to open Filter form
$sync.filterbutton.add_click({
    [void]$Form2.ShowDialog()
})

# Submit button on Filter form. Runs a refresh and closes Filter form.
$button3.add_click({
    $script:refreshrunspace = [PowerShell]::Create().AddScript({
        invoke-expression $sync.refresh
        $script:refreshrunspace.runspace.dispose()
        $script:refreshrunspace.dispose()
    })
    $runspace = [RunspaceFactory]::CreateRunspace()
    $runspace.ApartmentState = "STA"
    $runspace.ThreadOptions = "ReuseThread"
    $runspace.Open()
    $runspace.SessionStateProxy.SetVariable("sync", $sync)
    $script:refreshrunspace.Runspace = $runspace
    $script:refreshrunspace.BeginInvoke()
    $form2.close()
})

# This is the "Function" that loads the users into the User List
function USERLIST {
    $sync.users.rows.clear()
    $sync.users.refresh()
    $sync.index = $sync.result | where-object { $_.product -eq $sync.features.SelectedCells.value[0]}
    $h = 0
    if ( -not [String]::IsNullOrWhiteSpace($sync.index.users)) {
        if ($sync.index.users.count -gt 1 ) {
            do {
                if (-not [String]::IsNullOrWhiteSpace($sync.index.users[$h])) {
                    $sync.users.rows.add($sync.index.users[$h],$sync.index.host[$h],$sync.index.display[$h],$sync.index.borrowed[$h],$sync.index.remaining[$h])
                }
                $h++
            } until ( $h -eq $sync.index.users.count )
        } else {
            $sync.users.rows.add($sync.index.users,$sync.index.host,$sync.index.display,$sync.index.borrowed,$sync.index.remaining)
        }
    } else {
        $sync.users.rows.add("")
    }
}
$sync.userlist = get-content FUNCTION:\USERLIST

<#
        This "Function" only runs at launch. It queries the license server to get a list of years. Then it uses those years to download the respective Feature Code lists from Autodesk. 
    It saves the codes in a CSV file that it checks for on subsequent runs so it doesn't have to redownload them every time. The CSV file is saved to the LMUTIL directory set at the beginning.
    If a new year is detected on the License server, it will redownload its list.
    By default, it only grabs the latest 2 years. This can be changed by adjusting the "select-object -last 2" to a larger or smaller number on the "$sync.years = " line.
    I wouldn't go back any farther than 2017 though, since the feature code lists from Autodesk don't have all the necessary information before then.

    You can add extra product/feature codes in an "extras.csv" file in the same location as the "lmutil.exe" if you want to add your own. 
    The CSV should be laid out like this:

    Product,Feature,Package,Term
    Autodesk Vault 2019 Enterprise Addin,87238VLTEAD_2019_0F,plots,86820VLTEAD_T_F
    Autodesk Product Design Collection 2019,plots,plots,PDCOLL_T_F

    Any blank spots should be filled with the word "plots". I'm not sure why I picked that word.
#>
function FEATURECODES {
    $save = ""
    $sync.years = ((c:\windows\system32\cmd.exe /c ($sync.config.PATH + "\lmutil.exe") lmstat -c $sync.config.server -a | select-string "_[0-9][0-9][0-9][0-9]_").matches.Value -replace "_","" | sort-object -unique) | select-object -last 2
    foreach ($line in $sync.years) { 
        $save = $save + "_" +  $line 
        $sync.listbox2.items.add($line)
    }
    $sync.listbox2.items.add("Other")
    if ((Test-Path ($sync.config.PATH + "\autodesk_names" + $save + ".csv")) -eq $false) {
        $sync.progresstext.text = "Feature Code CSV not found. Downloading latest version..."
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
}
$sync.featurecodes = get-content FUNCTION:\FEATURECODES

<#
    This is the "meat and potatoes" of the script. It grabs a fresh status from the license server, then parses through that to get feature codes, users, license counts, and borrowed license statuses. 
    Then it compares that against the feature code list from Autodesk to give them readable Product names.
    Finally, it takes the new list it generated and compares it against the Filters selected in the filter menu before populating the Product table on the form.
#>
function REFRESH {
    $sync.progresstext.text = "Refreshing list..."
    $desk = cmd.exe /c ($sync.config.PATH + "\lmutil.exe") lmstat -c $sync.config.server -a
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
                        if ($user -match "linger:") {
                            $lingera = [int]($user -replace ".*linger:","").trim().split("")[0]
                            $lingerb = [int]($user -replace ".*linger:","").trim().split("")[2]
                        }
                        [PSCustomObject]@{
                            User = $user.trim().split(' ')[0]
                            Host = $user.trim().split(' ')[1]
                            Display = if (($user.trim()).split(' ')[2] -ne "start") {($user.trim()).split(' ')[2]} else { "" }
                            Borrowed = if ($user -match "linger:") { "Yes" } else { "No" };
                            Remaining = if ($user -match "linger:") { [math]::round((($lingerb - $lingera) / 86400),2) } else { 0 }
                        }
                    }
                } else { $temp2 = "" }
                $temp = [pscustomobject]@{
                     Product = $sync.auto[$h].product
                     Feature = if ($sync.auto[$h].feature -ne "plots") {$sync.auto[$h].feature} else { "" }
                     Used = $inuse
                     Issued = $issued
                     Users = $temp2.user
                     Host = $temp2.host
                     Display = $temp2.display
                     Borrowed = $temp2.borrowed
                     Remaining = $temp2.remaining
                }
                $sync.progresstext.text = $temp.Product
                $sync.result += $temp
            }
        }
        $h++
    } until ( $h -ge $sync.auto.count )
    foreach ($year in $sync.years) {
        foreach ($line in $sync.result) {
            if (($line.product -notmatch $year) -and ($line.feature -match $year) ) {
                $line.product = $line.product + " " + $year
            }
        }
    }
    $sync.result = $sync.result | Sort-Object used -Descending
    if ($false -eq $sync.firstrun) {
        $sync.selectionindex = [int]$sync.features.SelectedRows.index
        $sync.features.ClearSelection()
        $sync.users.ClearSelection()
        $sync.features.rows.clear()
        $sync.users.rows.clear()
        $sync.features.refresh()
        $sync.users.refresh()
    }
    if ($sync.users.selectedcells) { $sync.kick.enabled = $false }
    $null = foreach ($line3 in $sync.result) {
        foreach ($filter in $sync.listbox2.selecteditems) {
            if ($line3.product -match $filter) {
                $sync.features.rows.add($line3.product,$line3.used,$line3.issued)
            } elseif (($filter -eq "Other") -and (($line3.product | select-string -pattern "[0-9][0-9][0-9][0-9]").matches.value -le ( $sync.years[0] - 1 ))) {
                $sync.features.rows.add($line3.product,$line3.used,$line3.issued)
            }
        }
    }
    if ($false -eq $sync.firstrun) {
        $sync.features.rows[$sync.selectionindex].selected = $true
        invoke-expression $sync.userlist
    } elseif ($true -eq $sync.firstrun) {
        $sync.features.rows[0].selected = $true
        invoke-expression $sync.userlist
    }
    $sync.firstrun = $false
    $sync.progresstext.text = "Done."
}
$sync.refresh = get-content FUNCTION:\REFRESH
$script:startrunspace = [PowerShell]::Create().AddScript({
    invoke-expression $sync.featurecodes
    $script:startrunspace.runspace.dispose()
    $script:startrunspace.dispose()
})
$runspace = [RunspaceFactory]::CreateRunspace()
$runspace.ApartmentState = "STA"
$runspace.ThreadOptions = "ReuseThread"
$runspace.Open()
$runspace.SessionStateProxy.SetVariable("sync", $sync)
$script:startrunspace.Runspace = $runspace
$script:startrunspace.BeginInvoke()
[void]$Form2.ShowDialog()
[void]$Form.ShowDialog()


<# 
    This is the best way I've found to make sure the script closes when it's finished. 
    The runspaces didn't always like to clean up, and tended to leave the script open in the background.
#>
stop-process $pid