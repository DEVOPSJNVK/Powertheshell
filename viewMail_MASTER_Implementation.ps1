param(
    
    [Parameter(Mandatory=$true)]
    [ValidateSet(
    "UNIX",
    "Provisioning-CTB",
    "Activation-DCF-TS-Decommission",
    "Database-WINTEL-VHS",
    "Storage-Backup",
    "Factory")]$capability,

    [Parameter(Mandatory=$true)]
    [ValidateSet("Implementation","Verification")][string]$MailType,

   
    $LogPath = "..\Logs"


)


if($MailType -eq "Implementation"){

    $TaskType = "ITSK*"
    
}
else{

    $TaskType = "VTSK*"

}



# If Capability is Unix

if($capability -eq "UNIX"){

        $AssignmentGroups = "HP_GLOBAL_UNIX", `
                            "HP_GLOBAL_UNIX_BESTSHORE", `
                            "HP_GLOBAL_UNIX_PATCHING", `
                            "HP_GLOBAL_CTB_UNIX", `
                            "HP_GLOBAL_CTB_UNIX_BESTSHORE"



        $HP_GLOBAL_UNIX = "Chemmeri, Riju <riju.chemmeri@hpe.com>; `
                            'hp_unix_india_bss@list.db.com'; `
                            'hp_global_unix_operations@list.db.com'; `
                            ITO India GDC - DB UNIX Change Leads <itogdcindia.unixchglead@hpe.com>; `
                            hp_unix_change_coordinators@list.db.com; `
                            Dutta, Subhash Chandra <subhas2@hpe.com>; `
                            Ganesan, Sangeetha (Deutsche Bank - India UNIX) <sangeethag@hpe.com>; `
                            A, Ramana Kumar <ramana.kum.a@hpe.com>;"

        $HP_GLOBAL_UNIX_BESTSHORE = "Chemmeri, Riju <riju.chemmeri@hpe.com>; `
                                    'hp_unix_india_bss@list.db.com'; `
                                    'hp_global_unix_operations@list.db.com'; `
                                    ITO India GDC - DB UNIX Change Leads <itogdcindia.unixchglead@hpe.com>; `
                                    hp_unix_change_coordinators@list.db.com; Dutta, Subhash Chandra <subhas2@hpe.com>; `
                                    Ganesan, Sangeetha (Deutsche Bank - India UNIX) <sangeethag@hpe.com>; `
                                    A, Ramana Kumar <ramana.kum.a@hpe.com>;"

        $HP_GLOBAL_UNIX_PATCHING = "Mani, Manoj <mani@hpe.com>; `
                                    Pask, Skip Anthony <skip.pask@hpe.com>; `
                                    Mogallapu, Geetha <geetha.mogallapu@hpe.com>; `
                                    George, Prasanth Kuruvilla (ITO GCI) <prashanth.george@hpe.com>; `
                                    D, Saritha <saritha.d.vijay@hpe.com>; `
                                    Dasanna, Rangaiah (ITO) <rangaiah.d@hpe.com>; `
                                    F J, Arun <arun.fat.joseph@hpe.com>; `
                                    G M, Mohan (ITO GDC India DB Shift Management) <mohan.g-m@hpe.com>; `
                                    ITO GDC India DB Shift Management <itogdcindiadbshiftmgt@hpe.com>; `
                                    hpglsm@list.db.com;"

        $HP_GLOBAL_CTB_UNIX = "Patil, Yash <yash.patil@hpe.com>; `
                                Taylor, Neil <neilt@hpe.com>; `
                                Mcbennett, Neil <neil.mcbennett@hpe.com>; `
                                Bose, Anirban <anirban.bose@hpe.com>; `
                                Pradhan, Sabina (Delivery Manager) <sabina.pradhan@hpe.com>;"

        $HP_GLOBAL_CTB_UNIX_BESTSHORE = "Pradhan, Sabina (Delivery Manager) <sabina.pradhan@hpe.com>; `
                                        Lakshmanan, Vediappan <vediappan.lakshmanan@hpe.com>;"

        $COMMON_CHG = "Praveen Kumar B (praveen-kumarb@hpe.com); `
                        C, Praveen <praveen.c@hpe.com>; `
                        Kulkarni, Jai Ashok <jai.ash.kulkarni@hpe.com>; `
                        Deutsche Bank HP Change Mgmt Team <deutsche.bank.hp.change.mgmt.team@hpe.com>;"

}

# If Capability is Provisioning and CTB

if($capability -eq "Provisioning-CTB"){

        #$capability = 'Provisioning and CTB'
        $AssignmentGroups = "HP_GLOBAL_CTB_DM_BESTSHORE", `
                            "HP_GLOBAL_CTB_DM", `
                            "HP_GLOBAL_PROVISIONING_OPS", `
                            "HP_GLOBAL_CTB_STORAGE_UPLIFT"


        $HP_GLOBAL_CTB_DM_BESTSHORE = "HPCTB-DM-BESTSHORE@list.db.com; ajit.lakshman@hpe.com;"
        $HP_GLOBAL_CTB_DM = "HPCTB-DM@LIST.DB.DCOM; kanak.rana@hpe.com;"
        $HP_GLOBAL_PROVISIONING_OPS = "HPGP-OPS@list.db.com; anirban-hp.bose@db.com;"
        $HP_GLOBAL_CTB_STORAGE_UPLIFT = "india_storage_uplift_stm@list.db.com; mayank-hp.bajpai@db.com; mbajpai@hpe.com;"


        $COMMON_CHG = "Deutsche Bank HP Change Mgmt Team <deutsche.bank.hp.change.mgmt.team@hpe.com>; `
                        Talath, Ruhi <ruhi.talath@hpe.com>; `
                        Manish <manish2@hpe.com>;"


}


# If capability is Activation-DCF-TS-Decommission

if($capability -eq "Activation-DCF-TS-Decommission"){

        $AssignmentGroups = "HP_GLOBAL_ACTIVATION", `
                            "HP_GLOBAL_DECOMMISSION", `
                            "HP_APJ_DCENTRE_FACILITIES", `
                            "HP_UK_DCENTRE_FACILITIES", `
                            "HP_AMS_DCENTRE_FACILITIES", `
                            "HP_GERMANY_DCENTRE_FACILITIES", `
                            "HP_AMS_TECHNOLOGY_SERVICES", `
                            "HP_UK_TECHNOLOGY_SERVICES", `
                            "HP_APJ_TECHNOLOGY_SERVICES", `
                            "HP_GERMANY_TECHNOLOGY_SERVICES"

        $HP_GLOBAL_ACTIVATION = "globalactivations@list.db.com; 
                                 satish.narayanan-kutty@hpe.com;"
        $HP_GLOBAL_DECOMMISSION = "global_decommissions@list.db.com; 
                                   piyush.prasad@dxc.com; 
                                   Sai.dee.baliga@hpe.com;"
        $HP_APJ_DCENTRE_FACILITIES = "tom.kotaidis@hpe.com; jagan.krishnamoorthy@hpe.com; 
                                      ravi-xh.kiran@db.com; 
                                      twinky.leung@hpe.com; narayanar@hpe.com;"
        $HP_UK_DCENTRE_FACILITIES = "damien.alcock@hpe.com; ukitf.croydon@db.com;"
        $HP_AMS_DCENTRE_FACILITIES = "brian.carney@hpe.com; hbr.itfacilities@db.com ; gh.itfacilities@db.com;"
        $HP_GERMANY_DCENTRE_FACILITIES = "frank.teklenburg@hpe.com;felipe.foy@hpe.com;"
        $HP_AMS_TECHNOLOGY_SERVICES = "Hp-eng@list.db.com; shobna.persaud@hpe.com;"
        $HP_UK_TECHNOLOGY_SERVICES = "steve.littell@hpe.com; hpcdswatford@list.db.com; hpcdscroydon@list.db.com; hpcds@list.db.com;"
        $HP_APJ_TECHNOLOGY_SERVICES = "hpre.sg@list.db.com ;sheridan-fl.chen@hpe.com;"
        $HP_GERMANY_TECHNOLOGY_SERVICES = "stefanie.fischer@hpe.com;  mvs-hp-support@list.db.com;"


        $COMMON_CHG = "Deutsche Bank HP Change Mgmt Team <deutsche.bank.hp.change.mgmt.team@hpe.com>; 
                        T M, Zaheeruddin <zaheeruddin.tm@hpe.com>; 
                        S, Anusree <anusree.s@hpe.com>;"







}


# If capability is "Database-WINTEL-VHS"

if($capability -eq "Database-WINTEL-VHS"){

        $AssignmentGroups = "HP_GLOBAL_WINTEL", `
                            "HP_GLOBAL_WINTEL_BESTSHORE",
                            "HP_GLOBAL_ORACLE", ` 
                            "HP_GLOBAL_ORACLE_BESTSHORE", `
                            "HP_GLOBAL_SYBASE", `
                            "HP_GLOBAL_SYBASE_BESTSHORE", `
                            "HP_GLOBAL_MSSQL", `
                            "HP_GLOBAL_MSSQL_BESTSHORE"


        $HP_GLOBAL_WINTEL = "hp_global_wintel@list.db.com;"
        $HP_GLOBAL_WINTEL_BESTSHORE = "bineesh.madathil@hpe.com; tanmaya.kumar-nanda@hpe.com;"
        $HP_GLOBAL_ORACLE = "hpglobal.oracle@db.com; prakash.sr@hpe.com; nitin.bhargava2@hpe.com;"
        $HP_GLOBAL_ORACLE_BESTSHORE = "hpglobaloraclebestshore@list.db.com; nitin.bhargava2@hpe.com;"
        $HP_GLOBAL_SYBASE = "hpglobalsybase@list.db.com;" 
        $HP_GLOBAL_SYBASE_BESTSHORE = "hpglobalsybase_bestshore@list.db.com;"
        $HP_GLOBAL_MSSQL = "hpglobalmssql@list.db.com;"
        $HP_GLOBAL_MSSQL_BESTSHORE = "hpglobalmssql_bestshore@list.db.com;"

        $COMMON_CHG = "Deutsche Bank HP Change Mgmt Team <deutsche.bank.hp.change.mgmt.team@hpe.com>; `
                       N, Vinodkumar <vinodkumar.n@hpe.com>; `
                       Guldas, Vishwanath <vishwanath.guldas@hpe.com>; `
                       Manish <manish2@hpe.com>;"

}


# If capabilty is "Storage-Backup"


if($capability -eq "Storage-Backup"){

        $AssignmentGroups = "HP_GLOBAL_STORAGE", `
                            "HP_GLOBAL_STORAGE_NAS_CAS", `
                            "HP_GLOBAL_BACKUP"


        $HP_GLOBAL_STORAGE = "hp-nucleus-storage@hpe.com;"
        $HP_GLOBAL_STORAGE_NAS_CAS = "itogdcindia.db.nascas@hpe.com;"
        $HP_GLOBAL_BACKUP = "itogdcindia.db.backup.tsm@hpe.com;"

        $COMMON_CHG = "Deutsche Bank HP Change Mgmt Team <deutsche.bank.hp.change.mgmt.team@hpe.com>; `
                        Govind, Nalini <nalini.govind@hpe.com>; `
                        Talath, Ruhi <ruhi.talath@hpe.com>;"

}


#If capability is "Factory"

if($capability -eq "Factory"){

        $AssignmentGroups = "HP_GLOBAL_FACTORY", `
                            "HP_GLOBAL_FACTORY_UNIXSA", `
                            "HP_GLOBAL_FACTORY_ORACLEDBA", `
                            "HP_GLOBAL_FACTORY_MSSQL_SYBASE", `
                            "HP_GLOBAL_FACTORY_WINTELBA"



        $HP_GLOBAL_FACTORY = "HP_GLOBAL_FACTORY@list.db.com; Meddings, `
                                Andrew <andrew.meddings@hpe.com>;"

        $HP_GLOBAL_FACTORY_UNIXSA = "HP_GLOBAL_FACTORY_UNIXSA@list.db.com; `
                                    hp-unix-factory-india@list.db.com; `
                                    Beeching, David <david.beeching@hpe.com>;"

        $HP_GLOBAL_FACTORY_ORACLEDBA = "HP_GLOBAL_FACTORY_ORACLEDBA@list.db.com; `
                                        ITO India GDC - DB Transformation Oracle <itoindiagdc.dbtransoracle@hpe.com>; `
                                        Griffiths, Glynn <glynn.griffiths@hpe.com>;"

        $HP_GLOBAL_FACTORY_MSSQL_SYBASE = "HP_GLOBAL_FACTORY_MSSQL_SYBASE@list.db.com; `
                                            ITO India GDC - DB Transformation MS SQL Sybase <itoindiagdc.dbtransmssqlsybase@hpe.com>;"

        $HP_GLOBAL_FACTORY_WINTELBA = "HP_GLOBAL_FACTORY_WINTELSA@list.db.com; David.spendley@db.com;"


        $COMMON_CHG = "Deutsche Bank HP Change Mgmt Team <deutsche.bank.hp.change.mgmt.team@hpe.com>; `
                        T M, Zaheeruddin <zaheeruddin.tm@hpe.com>; `
                        S, Anusree <anusree.s@hpe.com>;"

}


# -----------------------------------------------------------
# -----------------------------------------------------------

#Function to Write Log

function Write-Log {

    <#
    .SYNOPSIS
       Write Custom log to a file

    .DESCRIPTION
       Writes log with date, Message Type and Keeps adding to the file.
       Call Write-Log into use the function.
       

    .EXAMPLE
   
       #Writes Date, Time, Error type to the path Provided
            For Message Type INFO
           Write-Log -Path C:\Temp\ -Message "Connected to vCenter" -MessageType INFO
    
    #>
    [cmdletbinding()]
    Param(

        [parameter(Mandatory=$true)]
           [string]$LogPath,
        
        [parameter(Mandatory=$true)]
        [ValidateSet('INFO','WARNING','ERROR')]
           [string]$MessageType,

        [parameter(Mandatory=$true)]
           [string]$Message

    )

    #######  Writing LOG  ######

    $FormattedDate = Get-Date -Format "yyyy-MM-dd"
    $LogFormattedDate = Get-Date -Format "hh-mm-ss yyyy-MM-dd"
    "[$MessageType] [$LogFormattedDate] $Message" | Out-File -FilePath $LogPath\$FormattedDate.log -Append
    
}


# ------------------------------------------------------------
# ------------------------------------------------------------


# File Name like "Open Implementation tasks to be closed by Unix teams - 18th Apr'18"

$ErrorActionPreference = "Stop"

Write-Log -LogPath $LogPath -MessageType INFO -Message "====================================="
Write-Log -LogPath $LogPath -MessageType INFO -Message ($MailType+" Task Violation for "+$capability)
Write-Log -LogPath $LogPath -MessageType INFO -Message "====================================="



$Date = Get-Date -Format "dd MMM yyyy"


$dirCheck = Test-Path -Path ("..\Generated Reports\"+$Date)

if($dirCheck -eq $false){

    $ReportDir = New-Item -Path ("..\Generated Reports\"+$Date) -ItemType directory
    $ReportDir =  $ReportDir.FullName
    Write-Log -LogPath $LogPath -MessageType INFO -Message ("created directory "+$ReportDir)

}
else{

    $ReportDir = (Get-ChildItem -Path ("..\Generated Reports") | Where-Object -FilterScript {$_.Name -eq $Date}).FullName
    Write-Log -LogPath $LogPath -MessageType INFO -Message ("Directory exists "+$ReportDir)

}


# Copy Template to Generated Reports

$FileName = "Open Implementation tasks to be closed by "+$capability+" teams - "+$Date+".xlsx"
$destFile = ($ReportDir+"\"+$FileName)
try{

    $CopyTemplate = Copy-Item -Path '..\Templates\Open Implementation tasks to be closed.xlsx' -Destination ($ReportDir+"\"+$FileName)
    Write-Log -LogPath $LogPath -MessageType INFO -Message ("Copied Task Vioation template to Genrated Reports as "+$FileName)

}

catch{

    Write-Log -LogPath $LogPath -MessageType ERROR -Message ("Unable to copy template to Generated Reports Location")
    Write-Log -LogPath $LogPath -MessageType ERROR -Message $_.Exception.Message
    Exit
}


# Import Excel Module

    Write-Log -LogPath $LogPath -MessageType INFO -Message "Importing Excel Module"

try{
    
    Import-Module -Name '..\Modules\ImportExcel\4.0.13\ImportExcel.psm1'
    Write-Log -LogPath $LogPath -MessageType INFO -Message "Excel Module Imported Successfully"
}
catch{
    Write-Log -LogPath $LogPath -MessageType ERROR -Message "Failed to Import Module"
    Write-Log -LogPath $LogPath -MessageType ERROR -Message $_.Exception.Message
        
}


# Import and export excel with required data
try{
    Write-Log -LogPath $LogPath -MessageType INFO "Getting the file name of the data sheet"
    $MasterFile = (Get-ChildItem -Path '..\DB Reports' | `
                       Where-Object -FilterScript {$_.Name -like "Change_Implementation_and_Verification_Task_Violation*"}).FullName
    Write-Log -LogPath $LogPath -MessageType INFO "File name is found"
    Write-Log -LogPath $LogPath -MessageType INFO $MasterFile
    Write-Log -LogPath $LogPath -MessageType INFO "Importing data"
    $RawData = Import-Excel -Path $MasterFile
    Write-Log -LogPath $LogPath -MessageType INFO "Imported data sccessfully"
}
catch{
    Write-Log -LogPath $LogPath -MessageType ERROR -Message "Unable to import data"
    Write-Log -LogPath $LogPath -MessageType ERROR -Message $_.Exception.Message
        
}

$data = @()
$MailsCC = @()

#Getting data for required Task Assignment Group

foreach($AssignmentGroup in $AssignmentGroups){

    try{

        Write-Log -LogPath $LogPath -MessageType INFO -Message "Exporting data for $AssignmentGroup"

        
       $da=  $RawData | `
            Where-Object -FilterScript {($_.'Task Assignment Group' -eq $AssignmentGroup) `
                                        -and ($_.'Task Reference Number' -like $TaskType)}

        Write-Log -LogPath $LogPath -MessageType INFO -Message "Exported data for $AssignmentGroup"


        # Assignmeng Group should not be null, if it is, then dont send Mail
        $Mailto = $null
        if ($da -ne $null){
            
            $Mailto = (Get-Variable -Name $AssignmentGroup).Value
            #$Mailto
    
        }




        # Adding data
        $MailsCC += $Mailto
        $data += $da
        
    }

    catch{

        Write-Log -LogPath $LogPath -MessageType ERROR -Message "Failed to Export data for $AssignmentGroup"
        Write-Log -LogPath $LogPath -MessageType ERROR -Message $_.Exception.Message
    
    }

    
}
try{

    Write-Log -LogPath $LogPath -MessageType INFO -Message "Exporting data to Excel sheet"
    
    $excel = $data | `
                Export-Excel -Path $destFile `
                                -WorkSheetname "Implementation Task Violation" `
                                -PassThru
            $sheet = $excel.Workbook.Worksheets["Implementation Task Violation"]
        
            #$sheet.Column(11) | Set-Format -NumberFormat "dd-mm-yyyy hh:mm:ss"
            Set-Format -Address $sheet.Cells["K:Z"] -NumberFormat "dd-mm-yyyy hh:mm:ss" 
            $excel.Save()
            $excel.Dispose()

            

    Write-Log -LogPath $LogPath -MessageType INFO "Exported data to $destFile"
}
catch{
    Write-Log -LogPath $LogPath -MessageType ERROR -Message "Failed to Export data"
    Write-Log -LogPath $LogPath -MessageType ERROR -Message $_.Exception.Message
        
}

$MailsCC += $COMMON_CHG
$MailsCC = [string]$MailsCC

# Sending Mail

# Creating Mail Object

Write-Log -LogPath $LogPath -MessageType INFO -Message "Creating Mail Object"

$Outlook = New-Object -ComObject Outlook.Application;
$Mail = $Outlook.CreateItem(0);
$mail.CC = $MailsCC
$MailsTo = [string]$data.'Task Implementer Email'
$Mail.To =  $MailsTo.Replace(" ",";")
$Mail.Subject = ("Immediate Action Required to close the "+$MailType+" tasks - "+$Date)

$Mail.Attachments.Add($destFile)

$Namespace = $Outlook.GetNameSpace("MAPI")
$User = $Namespace.CurrentUser.Name

$Mail.HTMLBody = @"
<html>
<p class=MsoNormal><span lang=EN-US style='color:black;mso-ansi-language:EN-US'>Hi
All,<o:p></o:p></span></p>
<p class=MsoNormal><span lang=EN-US style='color:black;mso-ansi-language:EN-US'>Please
find the attached list of tasks that has to be closed by you immediately.<o:p></o:p></span></p>
<p class=MsoNormal><b><span lang=EN-US style='background:yellow;mso-highlight:
yellow;mso-ansi-language:EN-US'>Note:</span></b><span lang=EN-US
style='background:yellow;mso-highlight:yellow;mso-ansi-language:EN-US'> <span
style='color:black'>All DXC $MailType teams are required to close tasks on
completion of their activity.</span></span><span lang=EN-US style='color:black;
mso-ansi-language:EN-US'><o:p></o:p></span></p>

<p class=MsoNormal><b><span lang=EN-US style='color:black;background:yellow;
mso-highlight:yellow;mso-ansi-language:EN-US'>Enforce Closure Mode: - New</span></b><b><span
lang=EN-US style='color:black;mso-ansi-language:EN-US'> <o:p></o:p></span></b></p>
<p class=MsoNormal><span style='background:yellow;mso-highlight:yellow'>Please
be informed that Enforced Closure Mode’ will be applied on <b>Wednesday 21<sup>st</sup>
March 2018 </b>and <b>reduces the overdue threshold from 5 days to 3 days.</b></span><b>
- New</b></p>

<p class=MsoNormal><span lang=EN-US style='color:black;mso-ansi-language:EN-US'>When
an Implementer Group is in ‘Enforced Closure Mode’, this will apply to all
members of that Implementer Group. When in ‘Enforced Closure Mode’ all members
of the Implementer Group will not be able to approve, create or plan any RfC
tasks – they will only be able to update and close their existing overdue
tasks. <o:p></o:p></span></p>

<p class=MsoNormal><span lang=EN-US style='color:black;mso-ansi-language:EN-US'>We
provide daily updates through the DORM to all capability leaders report of any
task &gt;24 hours so should never reach 20 days.<o:p></o:p></span></p>


<p class=MsoNormal><b><span lang=EN-US style='font-family:"Arial",sans-serif;
color:black;mso-ansi-language:EN-US'>Regards,<o:p></o:p></span></b></p>

<p class=MsoNormal><b><span lang=EN-US style='font-family:"Arial",sans-serif;
color:black;mso-ansi-language:EN-US'>$User</span></b><span
lang=EN-US style='font-family:"Arial",sans-serif;color:black;mso-ansi-language:
EN-US'><br>
Change Manager (Deutsche Bank Account)</span><span lang=EN-US style='font-size:
10.0pt;font-family:"Arial",sans-serif;color:black;mso-ansi-language:EN-US'><br>
<br>
T +91 8033858272<br>
<br>
<b>DXC Technology</b><br>
Pratik Tech Park, <br>
Electronic City Phase 1, </span><span lang=EN-US style='font-size:10.0pt;
mso-ansi-language:EN-US'><o:p></o:p></span></p>

<p class=MsoNormal><span lang=EN-US style='font-size:10.0pt;font-family:"Arial",sans-serif;
color:black;mso-ansi-language:EN-US'>Bangalore, Karnataka, India - 560100<br>
<br>
<a href="http://www.dxc.technology/"><span style='color:black'>dxc.technology</span></a>
</span><span lang=EN-US style='mso-ansi-language:EN-US'><o:p></o:p></span></p>

<p class=MsoNormal><span lang=EN-US style='mso-ansi-language:EN-US'><o:p>&nbsp;</o:p></span></p>
</html>



"@


$Mail.Display()
Write-Log -LogPath $LogPath -MessageType INFO -Message "Mail is displayed"