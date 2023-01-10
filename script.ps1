#region DATA
$basepath="D:\data"
$computerNameFilePath=Join-Path -Path $basePath -ChildPath "computers.txt"
$minutes=3
$pathXLSX=Join-Path -Path $basePath -ChildPath "$(Get-Date -Format "dd.MM.yyyy").xlsx"
$pathAttendance=Join-Path -Path $basePath -ChildPath "$(Get-Date -Format "dd.MM.yyyy")_attendance.txt"
$startDate=$(Get-Date)
$endDate=$($startDate.AddMinutes(15))
$computerList=[System.Collections.ArrayList]::new()
#endregion DATA

#region Functions
function Get-ComputerObject
{
    [CmdletBinding()]
    param (
        [Parameter(HelpMessage = "FilePath to get computernames",Position=0,ValueFromPipeline)]
        [Alias("FilePath")]
        [string]$path
    )
    foreach($line in [System.IO.File]::ReadLines($path))
    {
        if ($line -like "*Temat*")
        {
            $topic=$line
        }
        else 
        {
            [PSCustomObject]@{
                COMPUTERNAME= $line
                TOPIC = $($topic.Split(":")[1])
                COLOR = $($topic.Split(":")[2])
            }
        }
    }
}
function Test-Connection 
{
    [CmdletBinding()]
    param (
        [Parameter(HelpMessage = "ComputerName to check connection", Position = 0, ValueFromPipeline)]
        [Alias("Computer")]
        [string]$computerName
    )
    begin
    {}
    process
    {
        [PSCustomObject]@{
            COMPUTERNAME  = $computerName
            CONNECTION = Test-Path "\\$computerName\C$"
        }
    }
    end
    {}
}

function Check-LogonSessions
{
[CmdletBinding()]
    param (
        [Parameter(HelpMessage = "ComputerName to check logon sessions",Position=0,ValueFromPipeline)]
        [Alias("Computers")]
        [string]$computerName
    )
    begin
    {}
    process
    {
        $userFromStationCSV=(quser /server:$computerName 2>&1) -split "\n" -replace '\s{2,}', ','
        if ($userFromStationCSV -match "ID")
        {
            $userObjects=$userFromStationCSV | convertfrom-csv -Delimiter ','
            foreach($userObject in $userObjects)
            {
                if ($($userObject.STATE) -like "Active")
                {
                    [PScustomObject]@{
                        COMPUTERNAME=$computerName
                        USERNAME=$($userObject.USERNAME)
                        TIME=Get-Date -Format "dd.MM.yyyy_HH.mm"
                        ID=$($userObject.ID)
                        STATE=$($userObject.STATE)
                    }
                }
                else 
                {
                    [PScustomObject]@{
                        COMPUTERNAME=$computerName
                        USERNAME=$($userObject.USERNAME)
                        TIME=Get-Date -Format "dd.MM.yyyy_HH.mm"
                        ID=$($userObject.SESSIONNAME)
                        STATE=$($userObject.ID)
                    }
                }
            }    
        }
        else 
        {
            [PScustomObject]@{
                COMPUTERNAME=$computerName
                USERNAME="BRAK"
                TIME=Get-Date -Format "dd.MM.yyyy_HH.mm"
            }
        }

    }
    end
    {}
}

#WRAPPER FROM PSWriteOffice
function Get-ExcelTranslateFromR1C1 
{
    [CmdletBinding()]
    param (
        [Parameter(HelpMessage = "Row for Range",Position=0,ValueFromPipeline)]
        [Alias("Row")]
        [int]$rowNumber,
        [Parameter(HelpMessage = "Column for Range",Position=0,ValueFromPipeline)]
        [Alias("Column")]
        [int]$columnNumber = 1
    )
    $range = [OfficeOpenXml.ExcelAddress]::TranslateFromR1C1("R[$rowNumber]C[$columnNumber]", 0, 0)
    return $range
}


$computerList.AddRange($(Get-ComputerObject -FilePath $computerNameFilePath))

while($startDate -lt $endDate)
{
    $data=$(Get-Date -Format "dd.MM.yyyy_HH.mm")
    Write-Host $data -ForegroundColor Green
    
foreach($computerObject in $computerList)
{
    if ($($(Test-Connection -ComputerName $($computerObject.COMPUTERNAME)).CONNECTION))
    {
        
        $computerObjectUserName=$($($($computerObject.COMPUTERNAME) | Check-LogonSessions | Where-Object {$_.STATE -like "*Active*"}).USERNAME)
        if ([String]::IsNullOrEmpty($computerObjectUserName))
        {
            Add-Member -InputObject $computerObject -MemberType NoteProperty -Name $($data.Split("_")[1]) -Value "NO USER"
        }
        else
        {
            Add-Member -InputObject $computerObject -MemberType NoteProperty -Name $($data.Split("_")[1]) -Value $computerObjectUserName
        }
    }
    else 
    {
        Add-Member -InputObject $computerObject -MemberType NoteProperty -Name $($data.Split("_")[1]) -Value "NO CONNECTION"
    }
}

#TODO: Zmienić ścieżkę
$output=$($computerList | Format-Table -GroupBy TOPIC -Property COMPUTERNAME,TOPIC,"*.*" -Wrap)
$output>> $pathAttendance #FILE
$output | fl * #TERMINAL


Start-Sleep -Seconds $($minutes*60)
$startDate=$(Get-Date)
}

#Wypisanie danych na terminal
$excelHandler=$computerList | Export-Excel -Path $pathXLSX -PassThru -AutoSize

$ws = $excelHandler.Workbook.Worksheets["Sheet1"]
$lastRow=$ws.Dimension.End.Row
$lastDataColumn=$ws.Dimension.End.Column

#wykonanie naglowkow górnej tabeli
Set-ExcelRange -Address $ws.Cells["$(Get-ExcelTranslateFromR1C1 -Row 1 -Column ($lastDataColumn+1))"] -Value "LOGGED"
Set-ExcelRange -Address $ws.Cells["$(Get-ExcelTranslateFromR1C1 -Row 1 -Column ($lastDataColumn+2))"] -Value "NOLOGGED"
Set-ExcelRange -Address $ws.Cells["$(Get-ExcelTranslateFromR1C1 -Row 1 -Column ($lastDataColumn+3))"] -Value "COMPUTER USED"

#wykonanie naglowkow dolnej tabeli
Set-ExcelRange -Address $ws.Cells["A$($lastRow+6):B$($lastRow+6)"] -Value $ws.Cells["A1:B1"].Value
Set-ExcelRange -Address $ws.Cells["D$($lastRow+6):F$($lastRow+6)"] -Value $ws.Cells["$(Get-ExcelTranslateFromR1C1 -Row 1 -Column ($lastDataColumn+1)):$(Get-ExcelTranslateFromR1C1 -Row 1 -Column ($lastDataColumn+3))"].Value



2..$lastRow | Foreach-Object {
    #kolorowanie pierwszych 2 kolumn
    Set-ExcelRange -Address $ws.Cells["A$($_):B$($_)"] -BackgroundColor $($ws.Cells["C$($_)"].Value)
    
    #wypełnianie ostatnich kolumn z zajetoscia stacji
    $dataRowRangeSecondCoordinate=$(Get-ExcelTranslateFromR1C1 -Row $($_) -Column $lastDataColumn)
    Set-ExcelRange -Address $ws.Cells["$(Get-ExcelTranslateFromR1C1 -Row $($_) -Column ($lastDataColumn+1))"] -Formula "=COUNTIF(D$($_):$dataRowRangeSecondCoordinate,`"*.*`")" -NumberFormat 'General'
    Set-ExcelRange -Address $ws.Cells["$(Get-ExcelTranslateFromR1C1 -Row $($_) -Column ($lastDataColumn+2))"] -Formula "=COUNTIF(D$($_):$dataRowRangeSecondCoordinate,`"<>*.*`")" -NumberFormat 'General'
    Set-ExcelRange -Address $ws.Cells["$(Get-ExcelTranslateFromR1C1 -Row $($_) -Column ($lastDataColumn+3))"] -Formula "=($(Get-ExcelTranslateFromR1C1 -Row $($_) -Column ($lastDataColumn+1)))/(($(Get-ExcelTranslateFromR1C1 -Row $($_) -Column ($lastDataColumn+1)))+($(Get-ExcelTranslateFromR1C1 -Row $($_) -Column ($lastDataColumn+2))))" -NumberFormat 'Percentage'

    #kolorowanie pierwszych 2 kolumn wraz ze wstawieniem wartości-poniżej
    Set-ExcelRange -Address $ws.Cells["A$($lastRow+$($_)+5):B$($lastRow+$($_)+5)"] -BackgroundColor $($ws.Cells["C$($_)"].Value) -Value $ws.Cells["A$($_):B$($_)"].Value
    Set-ExcelRange -Address $ws.Cells["D$($lastRow+$($_)+5)"] -Formula "=COUNTIF(D$($_):$dataRowRangeSecondCoordinate,`"*.*`")" -NumberFormat 'General'
    Set-ExcelRange -Address $ws.Cells["E$($lastRow+$($_)+5)"] -Formula "=COUNTIF(D$($_):$dataRowRangeSecondCoordinate,`"<>*.*`")" -NumberFormat 'General'
    Set-ExcelRange -Address $ws.Cells["F$($lastRow+$($_)+5)"] -Formula "=($(Get-ExcelTranslateFromR1C1 -Row ($lastRow+$($_)+5) -Column 4))/(($(Get-ExcelTranslateFromR1C1 -Row ($lastRow+$($_)+5) -Column 4))+($(Get-ExcelTranslateFromR1C1 -Row ($lastRow+$($_)+5) -Column 5)))" -NumberFormat 'Percentage'
}


#Tworzenie wykresów
#TODO:Dostylizować wykresy poprzez nadanie im etykiet danych
#zakresy
$computersColumnRange="$(Get-ExcelTranslateFromR1C1 -Row 2 -Column 1):$(Get-ExcelTranslateFromR1C1 -Row $lastRow -Column 1)"
$loggedColumnRange="$(Get-ExcelTranslateFromR1C1 -Row 2 -Column ($lastDataColumn)):$(Get-ExcelTranslateFromR1C1 -Row $lastRow -Column ($lastDataColumn))"
$unloggedColumnRange="$(Get-ExcelTranslateFromR1C1 -Row 2 -Column ($lastDataColumn+1)):$(Get-ExcelTranslateFromR1C1 -Row $lastRow -Column ($lastDataColumn+1))"
$usageColumnPercentRange="$(Get-ExcelTranslateFromR1C1 -Row 2 -Column ($lastDataColumn+2)):$(Get-ExcelTranslateFromR1C1 -Row $lastRow -Column ($lastDataColumn+2))"

Add-ExcelChart -Worksheet $ws -ChartType ColumnClustered  -XRange $computersColumnRange -YRange $usageColumnPercentRange -Title "COMPUTER USAGE PERCENT" -XAxisTitleText "Procenty" -YMinValue 0 -YMaxValue 1.0 -YAxisTitleText "Komputery" -Width 1000 -Height 500 -LegendPosition Bottom -SeriesHeader "COMPUTER USAGE PERCENT" -Row $($lastRow+3) -Column 5
Add-ExcelChart -Worksheet $ws -ChartType ColumnStacked -XRange $computersColumnRange -YRange $loggedColumnRange,$unloggedColumnRange -Title "LOGGED IN PERCENT AND NO LOGGED IN PERCENT" -XAxisTitleText "Procenty" -YAxisTitleText "Komputery" -Width 1000 -Height 500 -LegendPosition Bottom -SeriesHeader "LOGGED","NO LOGGED" -Row $($lastRow+3) -Column 12

#kolorowanie logowań
Add-ConditionalFormatting -Address "C2:$(Get-ExcelTranslateFromR1C1 -Row $lastRow -Column $lastDataColumn)" -Worksheet $ws -RuleType ContainsText -ConditionValue "NO USER" -BackgroundColor "RED"
Add-ConditionalFormatting -Address "C2:$(Get-ExcelTranslateFromR1C1 -Row $lastRow -Column $lastDataColumn)" -Worksheet $ws -RuleType ContainsText -ConditionValue "NO CONNECTION" -BackgroundColor "BLUE"
Add-ConditionalFormatting -Address "C2:$(Get-ExcelTranslateFromR1C1 -Row $lastRow -Column $lastDataColumn)" -Worksheet $ws -RuleType ContainsText -ConditionValue "*.*" -BackgroundColor "GREEN"

$ws.DeleteColumn(3)

Close-ExcelPackage -Show $excelHandler

