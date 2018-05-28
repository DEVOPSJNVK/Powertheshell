function Task-Violation-excel {

#Declaring Parameters

Param(

    [parameter(Mandatory=$false)]
    [String]$reportPath,

    [parameter(Mandatory=$false)]
    [ValidateSet("Implementation","Verification")]
    [String]$taskType,

    [parameter(Mandatory=$false)]
    [String]$templateName = "Task_Violation_Template.xlsx",

    [parameter(Mandatory=$false)]
    [String]$masterReport = "Change_Implementation_and_Verification_Task_Violation*",

    [parameter(Mandatory=$false)]
    [String]$masterSheet = "Report 1",
    
    [parameter(Mandatory=$false)]
    [String]$emailRecipients = "Samynathan, Ragupathi <ragupathi.samynathan@hpe.com>;Murthy, Manoj Suresh <manoj-suresh.murthy@hpe.com>;Vardhan N, Harsha <harsha.vardhan-n@hpe.com>"
)

#$taskType

# Need to remove later

$reportDir = "C:\BIONIX\DB Reports"


# Creating File Name

$date = Get-Date -Format "dd MMM yyy"


$reportDir1 = "C:\BIONIX\Generated Reports"

if($taskType -eq "Implementation"){

     $reportPath = ("C:\BIONIX\Generated Reports"+"\"+"Implementation Task Violation - "+$date+".xlsx")

    }

else{
   
    $reportPath = ("C:\BIONIX\Generated Reports"+"\"+"Verification Task Violation - "+$date+".xlsx")

}



#Copy Template


    
    Write-Host "Copying template"

    Copy-Item -Path ("C:\BIONIX\Templates"+"\"+$templateName) -Destination $reportPath -Force

#Import Excel Module


    Write-Host "Importing Powershell Module"

    Import-Module -Name 'C:\BIONIX\Modules\ImportExcel\4.0.13\ImportExcel.psm1'



# Reading Master File


$masterReport = (Get-ChildItem -Path $reportDir | Where-Object -FilterScript {$_.Name -like $masterReport}).Name

$exFilePath = ($reportDir+"\"+$masterReport)

$xl = New-Object -comobject excel.application


$xl.Visible = $False
$xl.DisplayAlerts = $False
$wb = $xl.Workbooks.Open($exFilePath)
$ws = $wb.Worksheets.Item($masterSheet)



#Remove first 3 blank rows columns and close the file



for ($i = 1; $i -le 3; $i++) {
     If ($ws.Cells.Item($i, 2).Formula -eq "") {
        $Range = $ws.Cells.Item($i, 1).EntireRow
        $Range.Delete() > $null
        $i = $i - 1
        If ($ws.Cells.Item($i+1, 1).Formula -ne "") {break}
     }
}

for ($i = 1; $i -le 3; $i++) {
     If ($ws.Cells.Item(1, $i).Formula -eq "") {
        $Range = $ws.Cells.Item(1, $i).EntireColumn
        $Range.Delete() > $null
        $i = $i - 1
        If ($ws.Cells.Item(1, $i+1).Formula -ne "") {break}
     }
}

$wb.Save()
$wb.Close()
$xl.Quit()

[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb)
[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl)
[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($ws)


#Import Excel and Task violations excel


Write-Host "Importing and creating the sheets"

#If the type is Implememntaion Task Vioaltions

if($taskType -eq "Implementation"){
    
    Import-Excel -Path $exFilePath | `
    Where-Object -FilterScript {$_.'Task Reference Number' -like "IT*"} | `
    Export-Excel $reportPath -WorkSheetname "Implementation Task Violation"

    Import-Excel -Path $exFilePath | `
    Where-Object -FilterScript {$_.'Task Reference Number' -like "IT*"} | `
    Where-Object -FilterScript {($_.'Task Implementer' -eq "") -or ($_.'Task Implementer' -eq $null)} | `
    Export-Excel $reportPath -WorkSheetname "Blank Implementers"

    }

else{
    
    Import-Excel -Path $exFilePath | `
    Where-Object -FilterScript {$_.'Task Reference Number' -like "VT*"} | `
    Export-Excel $reportPath -WorkSheetname "Implementation Task Violation"

}


#Reading created Excel file for changing the time format


$colStart = "A1"
$colEnd = "AA1"
$colRange = ($colStart+":"+$colEnd)


$implxl = New-Object -comobject excel.application
$implxl.Visible = $true
$implxl.DisplayAlerts = $False
$wb1 = $implxl.Workbooks.Open($reportPath)

$sheetsName = "Implementation Task Violation","Blank Implementers"

foreach($sheetName in $sheetsName){

        $ws = $wb1.Worksheets.Item($sheetName)


        # Change the date and time format


        $Test = Import-Excel -Path $exFilePath
        $dateCells = ($Test | Get-Member).Name -like "*date*"


        foreach($dateCell in $dateCells) {
    
            $getNme = $ws.Range($colRange).find($dateCell)
            $cellAddress = $getNme.Address($False,$False)
            $ws.Columns.EntireColumn.item($cellAddress.Remove(1)).NumberFormat ="dd-mm-yyyy hh:mm:ss"
        }

}

    if($taskType -eq "Verification"){
        
        $veriSheet = $wb1.Worksheets.Item("Implementation Task Violation")
        [void]$veriSheet.Activate()
        $veriSheet.Name = "Verification task violation"
    }
$wb1.Save()
$wb1.Close()
$implxl.Quit()

[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb)
[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl)
[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($implxl)
[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb1)
[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($ws)
[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($getNme)
[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($cellAddress)
[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($veriSheet)

}

Task-Violation-excel -taskType Implementation
Task-Violation-excel -taskType Verification



#################################################### SCRIPT 2  #####################################################

$startCell = "A1"

$dormTemplate = "C:\BIONIX\Templates\DORM Hypercare-Template.pptx"

$date = Get-Date -Format "dd MMM yyyy"

# Copy template to Generate report Folder

$destPath = ("C:\BIONIX\Generated Reports\DORM Hypercare - "+$date+".pptx")

Copy-Item -Path $dormTemplate -Destination $destPath


$dormTemplate = $destPath
$path = "C:\BIONIX\Generated Reports\"
$impVioPath = $path+"Implementation Task Violation - "+$date+".xlsx"
$verVioPath = $path+"Verification Task Violation - "+$date+".xlsx"



#Create an instance of Excel.
$xl=New-Object -ComObject "Excel.Application"
$xl.Visible = $True


#Open the Excel file containing the table. 
$wb = $xl.workbooks.open($impVioPath)



#Create an instance of Powerpoint.
$objPPT = New-Object -ComObject "Powerpoint.Application"
#$objPPT.Visible ='Msotrue'
$objPresentatio = $objPPT.Presentations.Open($dormTemplate)

##########
##########
#Working on VIolation Pivot
###########
###########

$ws = $wb.Worksheets.Item("Dashboard")

#getting the last cell address
$dateCell= (Import-Excel $impVioPath -WorksheetName 'Dashboard' | `
                Where-Object -FilterScript {$_.'Row Labels' -like "*total*"}).'Count of Task Reference Number'

$colRange = $ws.UsedRange.AddressLocal($false,$false)

$getNme = $ws.Range($colRange).find($dateCell)
$endCell = $getNme.Address($False,$False)

#Creating Cell Range

$usedRange = ($startCell+":"+$endCell)

$pivotRange = $ws.Range("$usedRange")
$pivotRange.CopyPicture() 



$slide = $objPresentatio.Slides.Item(2)

$pic = $slide.Shapes.PasteSpecial()
$pic.Height = 300
$pic.Left = 100
$pic.Top = 100

$pivotWidth = $pic.Width


#########
#########
# Blank Implementers Pivot
#########

#getting the last cell address
$dateCell1= (Import-Excel $impVioPath -WorksheetName 'Dash Board - Blank Implementers' | `
                Where-Object -FilterScript {$_.'Row Labels' -like "*total*"}).'Count of Task Reference Number'

$ws2 = $wb.Worksheets.Item("Dash Board - Blank Implementers")

$colRange1 = $ws2.UsedRange.AddressLocal($false,$false)

            #$getNme = $ws.Range($colRange1).find($dateCell1)
            #$endCell = $getNme.Address($False,$False)

#Creating Cell Range

            #$usedRange = ($startCell+":"+$endCell)

$pivotRange1 = $ws2.Range("$colRange1")
$pivotRange1.CopyPicture() 

$slide = $objPresentatio.Slides.Item(2)

$pic1 = $slide.Shapes.PasteSpecial()
$pic1.Width = $pivotWidth
$pic1.Left = 600
$pic1.Top = 200


$objPresentatio.Save()
$objPresentatio.Close()

$objPPT.Quit()

$wb.Save()
$wb.Close()
$xl.Quit()

[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl)
[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb)
[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($ws)
[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($colRange)
[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($usedRange)
[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($pivotRange)

###########################################  Verification Task Violations ###################################



$xl1=New-Object -ComObject "Excel.Application"
$xl1.Visible = $True


#Open the Excel file containing the table. 
$wb1 = $xl1.workbooks.open($verVioPath)


#Create an instance of Powerpoint.
$objPPT = New-Object -ComObject "Powerpoint.Application"
$objPPT.Visible ='Msotrue'
$objPresentatio = $objPPT.Presentations.Open($dormTemplate)

##########
##########
#        Working on Verification Task Violation Pivot
###########
###########

$ws = $wb1.Worksheets.Item("Dashboard")

#getting the last cell address
$dateCell= (Import-Excel $verVioPath -WorksheetName 'Dashboard' | `
                Where-Object -FilterScript {$_.'Row Labels' -like "*total*"}).'Count of Task Reference Number'

$colRange = $ws.UsedRange.AddressLocal($false,$false)

$getNme = $ws.Range($colRange).find($dateCell)
$endCell = $getNme.Address($False,$False)

#Creating Cell Range

$usedRange = ($startCell+":"+$endCell)

$pivotRange = $ws.Range("$usedRange")
$pivotRange.CopyPicture() 



$slide = $objPresentatio.Slides.Item(3)

$pic = $slide.Shapes.PasteSpecial()
$pic.Width = 400
$pic.Top = 100
$pic.Left = 350


$objPresentatio.Save()
$objPresentatio.Close()

$objPPT.Quit()


$wb1.Save()
$wb1.Close()
$xl1.Quit()


##################################################  HEAD LINES TO SLIDES  #######################


$objPPT = New-Object -ComObject "Powerpoint.Application"
$objPPT.Visible ='Msotrue'
$objPresentatio = $objPPT.Presentations.Open($dormTemplate)

# Updating the Slide Headers

$text1 = $objPresentatio.Slides.Item(1)
$text1.Shapes.Title.TextFrame.TextRange.Text = ("Change Management Daily Violations – "+$date+" ServiceNow")

$text2 = $objPresentatio.Slides.Item(2)
$text2.Shapes.Title.TextFrame.TextRange.Text = ("Implementation Task Daily Violations "+$date+" Service Now")

$text3 = $objPresentatio.Slides.Item(3)
$text3.Shapes.Title.TextFrame.TextRange.Text = ("Verification Task Daily Violations – "+$date+" Service Now")

$text4 = $objPresentatio.Slides.Item(4)
$text4.Shapes.Title.TextFrame.TextRange.Text = ("Pending Approvals Violation List (Holds HP and DB)– "+$date)

$objPresentatio.Save()
$objPresentatio.Close()
$objPPT.Quit()

[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($objPPT)
[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($objPresentatio)
[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($text1)
[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($text2)
[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($text3)
[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($text4)



############################################   SCRIPT 3 ###############################################################



# pagFilePath -- Pending Approval Group File Path

$pagFilePath = (Get-ChildItem -Path "C:\BIONIX\DB Reports" | Where-Object -FilterScript {$_.Name -like "Pending Approval Groups*"}).FullName
$wbtoCopy = $pagFilePath

# wbApproval -- Template file for getting chart
$wbApproval = "C:\BIONIX\Templates\Pending Approvals (Violations).xlsx"

$date = Get-Date -Format "dd MMM yyyy"

#Copy file
Copy-Item -Path $wbApproval -Destination ("C:\BIONIX\Generated Reports\Pending Approvals (Violations) "+$date+".xlsx")
$wbApproval = ("C:\BIONIX\Generated Reports\Pending Approvals (Violations) "+$date+".xlsx")

# $dormTemplate -- Template file for Dorm PPT
#$dormTemplate = "C:\BIONIX\Templates\DORM Hypercare-Template.pptx"

# Creating excel Object

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $true
$workBook = $xl.Workbooks.Open($wbtoCopy)


$xl1 = New-Object -ComObject Excel.Application
$xl1.Visible = $true
$workBook1 = $xl1.Workbooks.Open($wbApproval)

$wstoCopy = $workBook.Worksheets.Item('Pending Approval Group')
$rangetoCopy = $wstoCopy.UsedRange
$rangetoCopy.Copy() | Out-Null


$worksheet1 = $workBook1.Worksheets.Item('RAWDATA')
$range1 = $worksheet1.Range("A1") 
$worksheet1.Paste()

#Closing First File
$workBook.Save()
$workBook.Close()
$xl.Quit()
[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl)
[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($workBook)
[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($wstoCopy)
[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($rangetoCopy)

$databaseSheet = $workBook1.Worksheets.Item('Database')
$reportDate = $databaseSheet.Cells.Range('J1')
$reportDate.Value() = Get-Date

$worksheet2 = $workBook1.Worksheets.Item('Dashboard')
$hptktCopy = $worksheet2.Range("B2:M23","B26:M27")
$hptktCopy.CopyPicture() | Out-Null

#Pasting it in Powerpoint

#Creating the Powerpoint Object
$objPPT = New-Object -ComObject "Powerpoint.Application"
$objPPT.Visible ='Msotrue'
$objPresentatio = $objPPT.Presentations.Open($dormTemplate)
$slide = $objPresentatio.Slides.Item(4)
$pic = $slide.Shapes.PasteSpecial()
$pic.Height = 350
$pic.Left = 225
$pic.Top = 75
$pic.Width = 684.1888


$worksheet2 = $workBook1.Worksheets.Item('Dashboard')
$hptktCopy = $worksheet2.Range("B45:M48")
$hptktCopy.CopyPicture() | Out-Null

$pic = $slide.Shapes.PasteSpecial()
$pic.Left = 225
$pic.Top = 430
$pic.Width = 684.1888

$objPresentatio.Save()
$objPresentatio.Close()
$objPPT.Quit()

$workBook1.Save()
$workBook1.Close()
$xl1.Quit()
[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl)
[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($workBook1)
[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet1)
[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($range1)
[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($databaseSheet)
[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($reportDate)
[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet2)
[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($hptktCopy)
[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($objPPT)
[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($objPresentatio)
[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($slide)
[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($pic)


##################################### SCRIPT 4 ##########################################

$date = Get-Date -Format "dd MMM yyyy"
$reportPath = ("C:\BIONIX\Generated Reports"+"\"+"Implementation Task Violation - "+$date+".xlsx")

$dormPath = (Get-ChildItem -Path "C:\BIONIX\DB Reports" | Where-Object -FilterScript {$_.Name -like "DORM Master Sheet*"}).FullName
$destDormPath = ("C:\BIONIX\Generated Reports\DORM Master Sheet - "+$date+".xlsx")

Copy-Item -Path $dormPath -Destination $destDormPath -Force

$xlShiftToRight = -4161

$dormxl=New-Object -ComObject Excel.Application
$dormxl.Visible = $true
$dormwb=$dormxl.workbooks.open($destDormPath)
$dormws = $dormwb.worksheets.Item('Status-2018')

$dormCellRange = $dormws.Range('F1').EntireColumn

#$dormCellRange = $dormxl.Range('F1').EntireColumn
[void] $dormCellRange.Insert($xlShiftToRight)

$dormDate = Get-Date -Format "dd MMM"
$dormDataCell = $dormws.Cells.Item(2,6)

$dormDataCell.Value() = $dormDate
$dormDataCell.NumberFormat = "dd MMM"

$implData = Import-Excel $reportPath

$dormUsedRange = $dormws.UsedRange.Address($False,$False)

$rowLabels = $implData.('Row Labels')

foreach($rowLable in $rowLabels){

            $rowLable

            if($rowLable -notlike "*Total*"){
            $getNme = $dormws.Range($dormUsedRange).find($rowLable)
            $cellAddress = $getNme.Address($False,$False)
            $cellAddress = $cellAddress.Replace('A',$null)
            $dormws.Cells.Item($cellAddress,6).value() = ( $implData | Where-Object -FilterScript {$_.'Row Labels' -eq $rowLable} ).'Count of Task Reference Number'
            }
       
            if($rowLable -like "*Total*"){
            $getNme = $dormws.Range($dormUsedRange).find($rowLable)
            $cellAddress = $getNme.Address($False,$False)
            $cellAddress = $cellAddress.Replace('A',$null)
            $dormws.Cells.Item(70,6).value() = ( $implData | Where-Object -FilterScript {$_.'Row Labels' -eq $rowLable} ).'Count of Task Reference Number'
            }

}
            
$dormwb.Save()

## Drawing line graph

$dormChart = $dormws.Shapes.AddChart().Chart
$dormChart.HasTitle = $true
$dormChart.ChartTitle.Text = "DORM Table"
$dormDataRange = $dormws.Range("F2:L2,F70:L70")
$table = $dormChart.SetSourceData($dormDataRange)


$xlChart=[Microsoft.Office.Interop.Excel.XLChartType]
$dormChart.ChartType = $xlChart::xlLine

$dormChart.CopyPicture()


$objPPT = New-Object -ComObject "Powerpoint.Application"
#$objPPT.Visible ='Msotrue'
$objPresentatio = $objPPT.Presentations.Open($destPath)
$slide = $objPresentatio.Slides.Item(1)

$pic = $slide.Shapes.PasteSpecial()
$pic.Height = 400
$pic.Left = 100
$pic.Top = 150



$dormwb.Close($false)
$dormxl.Quit()

$objPresentatio.Save()
$objPresentatio.Close()
$objPPT.Quit()

[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($objPPT)
[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($objPresentatio)
[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($pic)
[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($slide)



[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($dormxl)
[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($dormwb)
[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($dormws)
[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($dormCellRange)

###################  SCRIPT 5 (SEND EMAIL)  ########################

$Outlook = New-Object -ComObject Outlook.Application;
$Mail = $Outlook.CreateItem(0);
$mail.To = "Samynathan, Ragupathi <ragupathi.samynathan@hpe.com>;Murthy, Manoj Suresh <manoj-suresh.murthy@hpe.com>;Vardhan N, Harsha <harsha.vardhan-n@hpe.com>";

$date = Get-Date -Format "dd MMM yyyy"

$Mail.Subject = "Daily Report - $date"

$filesToAttach = (Get-ChildItem -Path "C:\BIONIX\Generated Reports" | `
            Where-Object -FilterScript {$_.Name -like "*$date*"} | `
            Where-Object -FilterScript {($_.Name -like "*DORM Hypercare -*") -or ($_.Name -like "*Implementation Task Violation -*") `
            -or ($_.Name -like "Verification Task Violation -*")}).FullName

foreach($file in $filesToAttach){

$Mail.Attachments.Add($file)

}

$Mail.HTMLBody = @"
<!DOCTYPE html>
<html>
<p>Hi All,</p>
<p>Please find the attached daily Change Violation report along with pending DXC Group Approvals for <b>$date</b></p>
<p>Below is the link for raw data file for Pending Group Approvals.</p>

<h3><p><a href="https://hpe.sharepoint.com/teams/DBK/LiveServices/07ServiceMgmt/Forms/AllItems.aspx?RootFolder=%2Fteams
%2FDBK%2FLiveServices%2F07ServiceMgmt%2F04%20Change%20Management%2FPending%20Approvals&FolderCTID=0x012000C782A4
88B11A234EA5709BDE95699FC6&View=%7B04247B17%2D1A57%2D49A1%2D9931%2D31B35F2C3BB0%7D#InplviewHash04247b17-1a57-49a
1-9931-31b35f2c3bb0=Paged%3DTRUE-p_SortBehavior%3D0-p_FileLeafRef%3DPending%2520Approvals%2520%2528Violations%252
9%252d18th%2520May%252exlsx-p_ID%3D11277-FolderCTID%3D0x012000C782A488B11A234EA5709BDE95699FC6-PageFirstRow%3D91-">
Pending Approval's Sharepoint Link</a></p></h3>

</html>
"@


$Mail.Send()