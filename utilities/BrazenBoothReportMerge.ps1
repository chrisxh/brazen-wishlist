#Requires -Modules ImportExcel
$BoothRatingsInputDirectory = "C:\Users\ChrisHoehn\Google Drive - chris@educatemn.org\Events\2021-12-14 Statewide Fair\Event Data\ercmn-20211214-rep-ratings-6xYYyCr"
$ReportOutputDirectory = "C:\Users\ChrisHoehn\Google Drive - chris@educatemn.org\Events\2021-12-14 Statewide Fair\Event Data\ercmn-20211214-rep-ratings-6xYYyCr\combined\"

$AllBoothEngagementCsv = "C:\Users\ChrisHoehn\Google Drive - chris@educatemn.org\Events\2021-12-14 Statewide Fair\Event Data\ercmn-Booth-Engagement-Report-Educate-Minnesota-Virtual-Career-Fair-for-Educators-6xYYyCr-12142021.csv"
$AllScheduledChatRatingCsv = "C:\Users\ChrisHoehn\Google Drive - chris@educatemn.org\Events\2021-12-14 Statewide Fair\Event Data\ercmn-20211214-rep-ratings-6xYYyCr\ercmn-20211214--scheduled-chats-rep-ratings-6xYYyCr.csv"

$BufferCsvPath =  "$env:TEMP\Booth-Buffer-"


#Create output directory if doesn't exist
If(-Not (test-path $ReportOutputDirectory))
{
      New-Item -ItemType Directory -Force -Path $ReportOutputDirectory
}


# Load the Excel application
$excel = New-Object -ComObject excel.application
# Suppress alerts when overwriting
$excel.DisplayAlerts = $False


Get-ChildItem –Path $BoothRatingsInputDirectory -Filter *.csv | Foreach-Object {

    #Do something with $_.FullName
    $Path = $_.DirectoryName
    $filename = $_.BaseName
    $csv = $_.FullName #Location of the source file
    if($csv -EQ $AllScheduledChatRatingCsv){
        #Skip the scheduled chat file
        continue
    }

    $xlsx = $ReportOutputDirectory + $filename + ".xlsx" # Names Excel file same name as CSV

    [regex]$rx = "\w+-(?<EventDate>\d{8,8})--(?<BoothCode>[\d\w\-]+)-rep-ratings-(?<EventCode>[\w\d]+)$"
    $m = $rx.Matches($filename) 
    
    if(-Not $m){
        #skip files for which we cannot match the event date, code, and booth code
        Write-Host "INVALID FILE: " + $csv
        continue
    }
    $eventdate = $m[0].Groups["EventDate"]
    $eventcode = $m[0].Groups["EventCode"]
    $boothcode = $m[0].Groups["BoothCode"]

    $BoothRatingBufferCsv = $BufferCsvPath + $boothcode + "-Rating.csv"
    $BoothEngagementBufferCsv = $BufferCsvPath + $boothcode + "-Engagement.csv"
    
    ########################
    #Load the ratings into the buffer
    Import-CSV $csv | Export-CSV $BoothRatingBufferCsv  -Force -NoTypeInformation

    ########################
    #Split out the ScheduledChats for this booth based on the email domain
    $email = Import-Csv $csv | Select-Object -First 1 -ExpandProperty "raterEmailAddress"
    if($email) {
        [regex]$rx = ".+(?<EmailDomain>@[^\s]+)$"
        $m = $rx.Matches($email) 

        if(-Not $m){
            #skip files for which the first row rater email is invalid 
            Write-Host "INVALID RATER EMAIL [$email]"
            Write-Host "SCAN for Scheduled Chats from booth $boothcode in $AllScheduledChatRatingCsv"
        }
        else{
            $emaildomain = $m[0].Groups["EmailDomain"]
            Write-Host "Searching for Scheduled Chats associated with $emaildomain"

            # Filter the relevant scheduled chat rows and append to the rating buffer
            (Import-CSV $AllScheduledChatRatingCsv | 
                ? "raterEmailAddress" -like "*$emaildomain" | 
                Select-Object | 
                ConvertTo-Csv -NoTypeInformation) | 
            Select-Object -Skip 1 | 
            Add-Content -Path $BoothRatingBufferCsv
            
        }
    }
    else{
        #skip files for which we cannot match the event date, code, and booth code
        Write-Host "NO RATER EMAIL FOUND IN " + $csv
        Write-Host "SCAN for Scheduled Chats from booth $boothcode in $AllScheduledChatRatingCsv"
    }

    ########################
    # Now, convert the ratings buffer csv to an Excel file
    
    # Create a new Excel workbook with one empty sheet
    $workbook = $excel.Workbooks.Add(1)
    #Import the ratings into the default sheet
    $worksheet = $workbook.worksheets.Item(1)
    $worksheet.Name = "Rep Ratings"

    $RatingsMeasure = (Import-Csv $BoothRatingBufferCsv | Measure-Object)
    if($RatingsMeasure.Count -GT 0){

        # Build the QueryTables.Add command and reformat the data
        $TxtConnector = ("TEXT;" + $BoothRatingBufferCsv)
        $Connector = $worksheet.QueryTables.add($TxtConnector,$worksheet.Range("A1"))
        $query = $worksheet.QueryTables.item($Connector.name)
        $query.TextFileCommaDelimiter = $true
        $query.TextFileParseType = 1
        $query.TextFileColumnDataTypes = ,1 * $worksheet.Cells.Columns.Count
        $query.AdjustColumnWidth = 1

        # Execute & delete the import ratings query
        $query.Refresh()
        $query.Delete()
    }
    else{
        Write-Host "NO RATINGS FOUND FOR BOOTH $boothcode"
    }
       


    ########################
    #Add the Booth Visits sheet and import matching rows from the master engagement sheet
    $worksheet = $workbook.worksheets.Add()
    $worksheet.Name = "Booth Visits"
    # Filter the relevant rows into a temp buffer
    Import-CSV $AllBoothEngagementCsv | ? "Booth Code" -EQ $boothcode | Export-CSV -Path $BoothEngagementBufferCsv -Force -NoTypeInformation

    # Build the QueryTables.Add command and reformat the data
    $TxtConnector = ("TEXT;" + $BoothEngagementBufferCsv)
    $Connector = $worksheet.QueryTables.add($TxtConnector,$worksheet.Range("A1"))
    $query3 = $worksheet.QueryTables.item($Connector.name)
    $query3.TextFileCommaDelimiter = $true
    $query3.TextFileParseType = 1
    $query3.TextFileColumnDataTypes = ,1 * $worksheet.Cells.Columns.Count
    $query3.AdjustColumnWidth = 1

    # Execute & delete the import ratings query
    $query3.Refresh()
    $query3.Delete()

    # Save & close the Workbook
    #$workbook.Save()
    $workbook.SaveAs($xlsx,51)
    $Workbook.Close()
    
    
    # release resources
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Connector) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($query) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($query3) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null

    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

$excel.Quit()

[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel) | Out-Null
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()