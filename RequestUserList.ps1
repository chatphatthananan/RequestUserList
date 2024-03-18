# Configuration
Write-Host "Importing dependencies"
$Cred = Import-Clixml -Path D:\GenericAccountCredentials.xml

# region Install-Module ImportExcel -AllowClobber -Force
Get-Module ImportExcel -ListAvailable | Import-Module -Force -Verbose
$modPath = "D:\SGTAM_DP\Project\Powershell Tool\functions\modUtil.psm1"
Import-Module -Name $modPath -Force
# endregion


# Clients PIC settings, check on this 1 week or 2 before the scheduled run.
$ClientPICs = @{
    <#
    ABSCBN = @{
        clientid = 55
        name = "xxx"
        email = "xxx"
    }
    #>
    AsiaTodayZeeTv = @{
        clientid = 52
        name = "xxx"
        email = "xxx"
    }
    BBC = @{
        clientid = 54
        name = "xxx"
        email = "xxx"
    }
    Dentsu = @{
        clientid = 7
        name = "xxx"
        email = "xxx"
    }
    <#
    DisneyFoxSG = @{
        clientid = 31
        name = "xxx"
        email = "xxx"
    }
    #>
    GroupM = @{
        clientid = 2
        name = "xxx"
        email = "xxx"
    }
    Hakuhodo = @{
        clientid = 18
        name = "xxx"
        email = "xxx"
    }
    Havas = @{
        clientid = 9
        name = "xxx and xxx,"
        email = @("xxx","xxx")
    }
    IMDA = @{
        clientid = 45
        name = "xxx"
        email = "xxx"
    }
    IPG =@{
        clientid = 10
        name = "xxx"
        email = "xxx"
    }
    MCI = @{
        clientid = 46
        name = "xxx"
        email = "xxx"
    }
    EssenceMediaCom = @{
        clientid = 13
        name = "xxx"
        email = "xxx"
    }
    MediacorpCIA = @{
        clientid = 3
        name = "xxx and xxx,"
        email = @("xxx", "xxx")
    }
    MediacorpOther = @{
        clientid = 34
        name = "xxx and xxx,"
        email = @("xxx", "xxx")
    }
    MindShare = @{
        clientid = 14
        name = "xxx"
        email = "xxx"
    }
    OMD = @{
        clientid = 15
        name = "xxx"
        email = "xxx"
    }
    Singtel = @{
        clientid = 4
        name = "xxx"
        email = "xxx"
    }
    Spark = @{
        clientid = 53
        name = "xxx"
        email = "xxx"
    }
    StarCom = @{
        clientid = 5
        name = "xxx"
        email = "xxx"
    }
    Wavemaker = @{
        clientid = 48
        name = "xxx"
        email = "xxx"
    }
    Zenith = @{
        clientid = 17
        name = "xxx"
        email = "xxx"
    }
}


# setup sql environment
$sqlserver = "xxx"
$sqlUser = $Cred.UserProg.UserName
$sqlPw = [System.Net.NetworkCredential]::New("", $Cred.UserProg.Password).Password
$SGTAMProdConnStr = "Data Source=$sqlserver;User ID=$sqlUser;Password=$sqlPw;Initial Catalog=xxx"

# email settings
# Recipients
$testingRecipients = @("xxx")


$emailSubjectErr = "[ERROR] Request User List"
$emailBodyErr = "Request User List error, please check the log. *This is auto generated email, do not reply to it."

# DataTemplate 
$dataTemplate = "D:\SGTAM_DP\Working Project\RequestUserList\DoNotDelete\DataTemplate.xlsm"
$excelPw = "xxx"

# variables
$todayDate = (Get-Date).toString("yyyy-MM-dd HHmmss")

# Exported files directory
$OutputPath = "D:\SGTAM_DP\Working Project\RequestUserList\Output\"

$SqlQueryOne = "EXEC SP_GetEvogeniusDistinctClientPIC"

# Transcript/Log path
$transcriptLog = "D:\SGTAM_DP\Working Project\RequestUserList\Log\RequestUserList_log_$($todayDate).log"

# Setup SGTAM log
$SGTAMLogConfig = @{
    connectionString = $SGTAMProdConnStr
    logTaskID = 103
    statusFlag = 2
    logMsg = 'User List was requested from the clients'
    logID = ''
}

try{

    Start-Transcript -Path $transcriptLog -Append

    $SGTAMLogConfig.statusFlag, $SGTAMLogConfig.logID = InsertSGTAMProgLog $SGTAMLogConfig

    Write-Host "INFO: Retrieving user lists and exporting to excels."

    # Establish connection to SQL database and retrieve user list data, store inside $DistinctEvogeniusPIC
    $DistinctEvogeniusPIC = Invoke-ExecuteQuery -connectionString $SGTAMProdConnStr -sqlCommand $SqlQueryOne
    #$DistinctEvogeniusPIC


    $clientNameArrayList = New-Object System.Collections.ArrayList
    $clientIDArrayList = New-Object System.Collections.ArrayList


    try{
        
        foreach ($row in $DistinctEvogeniusPIC){

            $ClientID = $row.clientID

            # if its DisneyFoxSG, rename client name to "Disney FOX SG"
            if($ClientID -eq 31){
                $ClientName = "Disney FOX SG"
            }
            else{
                $ClientName = $row.clientName
            }

          
            # Excluding Discovery, ABSCBN and Disney infos from the arrayLists, these 2 arrays will be used in the sending email parts below.
            if($ClientID -eq 29 -or $ClientID -eq 31 -or $ClientID -eq 55){
                Write-Host "Skipping "$ClientName
            }
            else{
                $clientIDArrayList.Add($ClientID)
                $clientNameArrayList.Add($ClientName)  
            }
            
          

            $SqlQueryTwo = "EXEC SP_GetEvogeniusUserListByClientID @ClientID = $ClientID"
            $UsersFromEachClient = Invoke-ExecuteQuery -connectionString $SGTAMProdConnStr -sqlCommand $SqlQueryTwo
    
            # Create excel object
            $objExcel = new-object -comobject excel.application
            $objExcel.Visible = $False
            $objExcel.DisplayAlerts = $False

            # Skip discovery, disney and ABNCBS, export each client userlist out to excel
            if($ClientID -ne 29 -or $ClientID -ne 55 -or $ClientID -31){
				
				# Special handling for Singtel2, Singtel2 users will be add to the Singtel list.
				if($ClientID -eq 4){
					
					if(Test-Path $OutputPath$ClientName".xlsx"){
						Remove-Item $OutputPath$ClientName".xlsx"
					}

					
					$row = $UsersFromEachClient.NewRow()
					$row."Client Name" = ("xxx")
					$row."User's Name" = ("xxx")
					$row.Email = ("xxx")
					$UsersFromEachClient.Rows.Add($row)

                    $UsersFromEachClient | Select * -ExcludeProperty "Client Name", RowError, RowState, HasErrors, Name, Table, ItemArray | Export-Excel -Path $OutputPath$ClientName".xlsx"
					$objWorkbook = $objExcel.Workbooks.Open($OutputPath+$ClientName+".xlsx")
					$worksheet = $objWorkbook.Worksheets['Sheet1']
        
					# Apply autofit to all columns in the final file
					$worksheet.Cells.EntireColumn.AutoFit()
					
				}	
				else{	
					if(Test-Path $OutputPath$ClientName".xlsx"){
						Remove-Item $OutputPath$ClientName".xlsx"
					}
					$UsersFromEachClient | Select * -ExcludeProperty "Client Name", RowError, RowState, HasErrors, Name, Table, ItemArray | Export-Excel -Path $OutputPath$ClientName".xlsx"
					$objWorkbook = $objExcel.Workbooks.Open($OutputPath+$ClientName+".xlsx")
					$worksheet = $objWorkbook.Worksheets['Sheet1']
        
					# Apply autofit to all columns in the final file
					$worksheet.Cells.EntireColumn.AutoFit()
				}
            }

            $workSheet.Cells[1, 1].EntireRow.Font.Bold = $true
            #$objWorkbook.Save()
			$objWorkbook.Close($true)

        }

        #$objExcel.Quit()
        $objExcel.Quit()
        #[System.Runtime.Interopservices.Marshal]::ReleaseComObject($objExcel)
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($objExcel)
        Stop-Process -name excel
       
       }catch{
            #throw "Error in retrieving users from SGTAMProd or error in exporting to excel, please check."
            Write-Host $_
       }
}catch{
    $SGTAMLogConfig.statusFlag = 2
    $SGTAMLogConfig.logMsg = $_.Exception.Message
    Write-Warning $SGTAMLogConfig.logMsg

}finally{
    if ($SGTAMLogConfig.statusFlag -eq 2){
        Write-Host "Sending Exception email"
		try{
			Send-MailMessage -From "xxx" -Subject $emailSubjectErr -To "xxx" -Body $emailBodyErr -SmtpServer "xxx" -Port 25 -BodyAsHtml -ErrorAction Stop
			Write-Host "Error email sent!"
			
		}catch{
			$SGTAMLogConfig.logMsg = $SGTAMLogConfig.logMsg + "`n" + $_.Exception.Message
			Write-Warning $SGTAMLogConfig.logMsg
		}

    }elseif ($SGTAMLogConfig.statusFlag -eq 1){
        $SGTAMLogConfig.logMsg = 'Request User List completed succesfully'
        
		try{        
            
            for ($counter=0; $counter -lt $clientIDArrayList.Count; $counter++){

                $clName = $clientNameArrayList[$counter] # clientName to look for the attachment for email
                switch($clientIDArrayList[$counter]){
                    2 {$clientName = "GroupM"
                       $picName = $ClientPICs.GroupM.name
                       $picEmail = $ClientPICs.GroupM.email
                                    
                    }
                    3 {$clientName = "Mediacorp CIA"
                       $picName = $ClientPICs.MediacorpCIA.name
                       $picEmail = $ClientPICs.MediacorpCIA.email
                                    
                    }
                    4 {$clientName = "Singtel"
                       $picName = $ClientPICs.Singtel.name
                       $picEmail = $ClientPICs.Singtel.email
                                    
                    }
                    5 {$clientName = "StarCom"
                       $picName = $ClientPICs.StarCom.name
                       $picEmail = $ClientPICs.StarCom.email
                                    
                    }
                    7 {$clientName = "Dentsu"
                       $picName = $ClientPICs.Dentsu.name
                       $picEmail = $ClientPICs.Dentsu.email
                                    
                    }
                    9 {$clientName = "Havas"
                       $picName = $ClientPICs.Havas.name
                       $picEmail = $ClientPICs.Havas.email
                                    
                    }
                    10 {$clientName = "IPG"
                       $picName = $ClientPICs.IPG.name
                       $picEmail = $ClientPICs.IPG.email
                                    
                    }
                    13 {$clientName = "EssenceMediacom"
                       $picName = $ClientPICs.EssenceMediaCom.name
                       $picEmail = $ClientPICs.EssenceMediaCom.email
                                    
                    }
                    14 {$clientName = "Mindshare"
                       $picName = $ClientPICs.MindShare.name
                       $picEmail = $ClientPICs.MindShare.email
                                    
                    }
                    15 {$clientName = "OMD"
                       $picName = $ClientPICs.OMD.name
                       $picEmail = $ClientPICs.OMD.email
                                    
                    }
                    17 {$clientName = "Zenith"
                       $picName = $ClientPICs.Zenith.name
                       $picEmail = $ClientPICs.Zenith.email
                                    
                    }
                    18 {$clientName = "Hakuhodo"
                       $picName = $ClientPICs.Hakuhodo.name
                       $picEmail = $ClientPICs.Hakuhodo.email
                                    
                    }
                    34 {$clientName = "Mediacorp Other"
                       $picName = $ClientPICs.MediacorpOther.name
                       $picEmail = $ClientPICs.MediacorpOther.email
                                    
                    }
                    45 {$clientName = "IMDA"
                       $picName = $ClientPICs.IMDA.name
                       $picEmail = $ClientPICs.IMDA.email
                                    
                    }
                    46 {$clientName = "MCI"
                       $picName = $ClientPICs.MCI.name
                       $picEmail = $ClientPICs.MCI.email
                                    
                    }
                    48 {$clientName = "Wavemaker"
                       $picName = $ClientPICs.Wavemaker.name
                       $picEmail = $ClientPICs.Wavemaker.email
                                    
                    }
                    52 {$clientName = "Asia Today"
                       $picName = $ClientPICs.AsiaTodayZeeTv.name
                       $picEmail = $ClientPICs.AsiaTodayZeeTv.email
                                    
                    }
                    53 {$clientName = "Spark Foundry"
                       $picName = $ClientPICs.Spark.name
                       $picEmail = $ClientPICs.Spark.email
                                    
                    }
                    54 {$clientName = "BBC"
                       $picName = $ClientPICs.BBC.name
                       $picEmail = $ClientPICs.BBC.email
                                    
                    }

                }
                
                
                    
                $emailSubject = "User List Verification - $clientName"

                
                $emailBody = "
				<p>*This is auto generated email, kindly reply and attach the list back to SGTAMDPTeam@gfk.com instead.</p>
				<p>Hi $picName</p>
			    <p>Good day to you.</p>
			    <p>Could you please help to check the attached user list if it is accurate as compared to your end?</p>
			    <p>Thank you.</p>"
                
				
                Send-MailMessage -From "xxx" -Subject $emailSubject -To $picEmail  -Cc "xxx" -Body $emailBody -Attachments $OutputPath$clName".xlsx" -SmtpServer "xxx" -Port 25 -BodyAsHtml -ErrorAction Stop
			    Write-Host "Successfully sent email to $clientName."

            }


		    Write-Host $SGTAMLogConfig.logMsg
        }catch{
			$SGTAMLogConfig.statusFlag = 3
			$SGTAMLogConfig.logMsg = $_.Exception.Message
			Write-Warning $SGTAMLogConfig.logMsg
		}
    }
	UpdateSGTAMProgLog $SGTAMLogConfig
    Stop-Transcript


}


