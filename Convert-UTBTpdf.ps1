# 
# Convert-UTBTpdf.ps1 - This imports a LT report pdf, parses it for DINs, and exports the DIN list as a csv
# Author              - Larkin O'quinn
# Date                - 121521
#
####--Changelog--####
# Rev 1 - Initial version
# Rev 2 - Consolidated variable declarations, changed file prompt text to "Lifetrak" rather than "UTBT", and changed final file name to reflect original file. LO 121621
# Rev 3 - Added new message at end and Invoke-Item for exported document. LO 121721
# Rev 4 - Added specific handling for different LT reports. Added regex [-] to filter out some weird parsing behavior in some reports. LO 010522
# Rev 5 - Added TSRUNSC2 as acceptable report type. LO 051622

#Variable Declarations
[string]$contents = $null
[string[]]$substrings = $null
[System.Collections.Generic.List[pscustomobject]]$finallist = @{}
#$shortdate = (get-date).ToString('MMddyy')
[string]$reporttype = $null

#Library import
Write-Host "Importing libraries from file server, please wait..." -ForegroundColor Yellow
Add-Type -Path \\ctsfile\public\larkin1\library\itextsharp.dll -ErrorAction SilentlyContinue
Add-Type -path \\ctsfile\public\larkin1\library\BouncyCastle.Crypto.dll

#OpenFileDialog form for retrieval of pdf
Write-host "Please select LifeTrak pdf to convert" -BackgroundColor Black
Add-Type -AssemblyName System.Windows.Forms
$file = New-Object System.Windows.Forms.OpenFileDialog
$file.filter = "pdf (*.pdf)| *.pdf"
$file.Title = "Please select LifeTrak pdf to convert"
[void]$file.ShowDialog()
$file.FileName

#pdf conversion to string array
$pdf = New-Object itextsharp.text.pdf.pdfreader -ArgumentList "$($file.filename)"
for ($i = 1;$i -le $pdf.NumberOfPages; $i++) {
    $contents += [iTextSharp.text.pdf.parser.PdfTextExtractor]::GetTextFromPage($pdf,$i)
    }
$substrings = $contents.Split([Environment]::NewLine)

#Checks type of report
for ($j =0; $j -lt 6; $j++) {
    switch ($substrings[$j]) {
        {($_ -like '*RGREXLOG*')} {
            $reporttype = "RGREXLOG"
            }
        {($_ -like '*MCRUINT2*')} {
            $reporttype = 'MCRUINT2'
            }
        {($_ -like '*TSRUNTST*')} {
            $reporttype = 'TSRUNTST'
            }
        {$_ -like '*TSRUNTS2*'} {
            $reporttype = 'TSRUNTS2'
            }
        {($_ -like '*RSMREPRT*')} {
            $reporttype = 'RSMREPRT'
            }
        {($_ -like '*TERVOID*')} {
            $reporttype = 'TERVOID'
            }
        {($_ -like '*TSRRSTA*')} {
            $reporttype = 'TSRRSTA'
            }
        {($_ -like '*TSRUNSC2*')} {
            $reporttype = 'TSRUNSC2'
            }
        }
    }
if ($reporttype -eq "") {
    Write-Host "You have selected an incompatible report type. Current supported reports: RGREXLOG, MCRUINT2, TSRUNTST, TSRUNTS2, RSMREPRT, TERVOID, TSRRSTA, and TSRUNSC2.`n`nProgram will be closed`n" -ForegroundColor Red
    pause
    exit}

#Filter and trim DINs
#--Basic cleaning (splits on - preceding the check digit, keeps [0]), optional duplicate cleaning
if ($reporttype -eq 'MCRUINT2' -or $reporttype -eq 'TSRUNTST' -or $reporttype -eq 'TERVOID' -or $reporttype -eq 'TSRRSTA' -or $reporttype -eq 'TSRUNTS2' -or $reporttype -eq 'TSRUNSC2') {
    foreach ($string in $substrings) {
        if ($string -match 'W\d\d\d\d\d\d\d\d\d\d\d\d[-]') {
            $temp = $string.Split("-")
            $result = '' | select DIN
            $result.DIN = $temp[0]
            $finallist.Add($result) = $null
            }
        if ($reporttype -eq 'TERVOID') {
            #Duplicate cleaner module. Yuck
            $duplicate = ($finallist | Group-Object -Property "DIN" | Where-Object -FilterScript {$_.Count -gt 1})
            foreach ($thing in $duplicate) {
                $finallist.Remove($thing.group[0]) | Out-Null
                }
            }
        }
    }
#--RGREX cleaning (splits on spaces, keeps [1])
elseif ($reporttype -eq 'RGREXLOG') {
    foreach ($string in $substrings) {
        if ($string -match 'W\d\d\d\d\d\d\d\d\d\d\d\d') {
            $temp = $string.Split(" ")
            $result = '' | select DIN
            $result.DIN = $temp[1]
            $finallist.Add($result) = $null
            }
        }
    }
#--RSM cleaning (splits on spaces, keeps [0]), cleans duplicates
elseif ($reporttype -eq 'RSMREPRT') {
    foreach ($string in $substrings) {
        if ($string -match 'W\d\d\d\d\d\d\d\d\d\d\d\d') {
            $temp = $string.Split(" ")
            $result = '' | select DIN
            $result.DIN = $temp[0]
            $finallist.Add($result) = $null
            }
        }
    #Duplicate cleaner module. Yuck
    $duplicate = ($finallist | Group-Object -Property "DIN" | Where-Object -FilterScript {$_.Count -gt 1})
    foreach ($thing in $duplicate) {
        $finallist.Remove($thing.group[0]) | Out-Null
        }

    }


#Export to desktop
$path = "C:\Users\$env:USERNAME\Desktop\$($($file.SafeFileName).Replace(".pdf",''))DINList.csv"
$finallist | Export-Csv -Path "C:\Users\$env:USERNAME\Desktop\$($($file.SafeFileName).Replace(".pdf",''))DINList.csv" -NoTypeInformation



Write-host ""
Write-Host "$($($file.SafeFileName).Replace(".pdf",''))DINList.csv has been placed on the desktop" -ForegroundColor DarkGreen
Write-host ""
Write-Host "Please verify number of exported DINs against Total Units on the LT report!" -ForegroundColor Green
Write-Host ""

Invoke-Item -Path $path

pause