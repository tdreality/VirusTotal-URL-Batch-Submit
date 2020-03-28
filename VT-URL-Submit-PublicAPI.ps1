#This code is for submitting a batch of links to VirusTotal, created instead of submitting the links one by one on their website
#This code is designed for the public API use, you will see below that this pauses/sleeps for 70 seconds since VirusTotal API can only do 4 queries a minute.
#This code is still not in its finest form but feel free to use it.
#Author: tdreality
#Date: 20171118
#Version: 1.1



#Open Excel App
$EXCEL = New-Object -ComObject Excel.Application
$EXCEL.visible = $true
$Workbook = $EXCEL.Workbooks.Add()
$Worksheet = $Workbook.WorkSheets.Item(1)

#Write Headers
$Worksheet.Cells.Item(1,1) = 'Link'
$Worksheet.Cells.Item(1,2) = 'Positives'

#Assign initial counter
$rowCounterLinks = 2
$rowCounterPositives = 2
$MaxPerMinute = 0



##########################CHANGE THIS SECTION##########################
#Set Variables
$CurlLocation = "<LOCATION OF CURL> ...\bin\curl.exe"
$ReportLinks = New-Object System.Collections.ArrayList
$apikey="<YOUR API KEY, REGISTER @ VIRUSTOTAL TO GET YOUR API KEY>"
$tempFolder = "<TEMP FOLDER FOR TRAVERSING OUTPUTS> ....\temp"
######################################################################

#Get Names of computer
$LinksToUpload=Get-Content "<LOCATION OF Links.txt"

#Check installed Application
foreach ($Link in $LinksToUpload)
{
    #Add the "Link" to the spreadsheet
    $Worksheet.Cells.Item($rowCounterLinks,1) = $Link
    $rowCounterLinks++

    $arglist = "--request POST --url `"https://www.virustotal.com/vtapi/v2/url/scan`" --data `"apikey=$apikey`" --data `"url=$Link`""
    Start-Process -FilePath $CurlLocation -ArgumentList $arglist -redirectstandardoutput "$tempFolder\Output$MaxPerMinute.txt"
    
    #Sleep for 2 seconds, file would still be null at this point until reply from VirusTotal arrives.
    Start-Sleep -s 4

    #Retreive data from Output textfile and delete the file after
    $rawScanData = Get-Content "$tempFolder\Output$MaxPerMinute.txt"
    Remove-Item "$tempFolder\Output$MaxPerMinute.txt"

    #Parse using parenthesis(") delimiter
    $splitScanData = $rawScanData.Split("`"")

    #Get the value of the scan ID
    $ReportLinks.Add($splitScanData[21].Split("/")) > $null
    
    #Clear variables
    Clear-Variable -Name "rawScanData"
    Clear-Variable -Name "splitScanData"
    Clear-Variable -Name "arglist"

    #Increament Counters
    $MaxPerMinute++

    #VirusTotal Public API can only do 4 queries a minute. This puts the program to sleep and waits for a minute before uploading again. Made it to 70seconds instead of 60seconds since sometimes it still gives an error @ 60 seconds
    If($MaxPerMinute -eq 4)
    {
    #Sleep for 70 seconds. See last comment above for reason
    Start-Sleep -s 70

    #Retrieving URL Reports/Results from VirusTotal
    #After sleeping for 50 seconds, submitted URLs should finish scanning and this would be a good time to retrieve the results
    #Loop for 4x since you have 4 submitted queries/urls
    for ($i=0; $i -le 3; $i++)
    {
    #Build up the argument list with the ReportLinks gathered earlier
    $arglist = "--request GET --url `"https://www.virustotal.com/vtapi/v2/url/report?apikey=$apikey&resource=`""
    $arglist += $ReportLinks[$i]
    $arglist += "`""
    #Execute CURL
    Start-Process -FilePath $CurlLocation -ArgumentList $arglist -redirectstandardoutput "$tempFolder\OutputReport$i.txt"

    #Sleep for 2 seconds to wait for the response from VirusTotal
    Start-Sleep -s 4

    #Retreive data from Output textfile and delete the file after
    $rawScanDataRetrieve = Get-Content "$tempFolder\OutputReport$i.txt"
    Remove-Item "$tempFolder\OutputReport$i.txt"
    $splitRawScanDataRetrieve = $rawScanDataRetrieve.Split(",")
    $splitScanDataRetrieve    = $splitRawScanDataRetrieve[9].Split(":")
    #Add the "Link" to the spreadsheet
    $Worksheet.Cells.Item($rowCounterPositives,2) = $splitScanDataRetrieve[1]
    $rowCounterPositives++

    #Delete Array to clear variables
    Clear-Variable -Name "rawScanDataRetrieve"
    Clear-Variable -Name "splitRawScanDataRetrieve"
    Clear-Variable -Name "splitScanDataRetrieve"
    Clear-Variable -Name "arglist"
    }

    #Delete Array to clear variables
    Clear-Variable -Name "ReportLinks"
    Clear-Variable -Name "MaxPerMinute"

    #Recreate Array after deleting
    $ReportLinks = New-Object System.Collections.ArrayList
 
    #Reset Counter to zero(0). A minute has passed and the program can do 4 queries again
    $MaxPerMinute=0

    #VirusTotal Public API can only do 4 queries a minute. This puts the program to sleep and waits for a minute before uploading again. Made it to 70seconds instead of 60seconds since sometimes it still gives an error @ 60 seconds
    Start-Sleep -s 70
    }
}
