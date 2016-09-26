#######################################################################################################################################################################
##
## getContactsFromAD
##
## 2016-09-26
##
## Connect to AD, pull out all the information regarding users and retrieve an csv file
##
## João Perdigão
##
#######################################################################################################################################################################



PROCESS #This is where the script executes 
{ 

$starttime = Get-Date -Format HH:mm:ss

write-host "----------> [companyName] Systems" -ForegroundColor red -BackgroundColor blue
write-host "----------> Getting all users from [officeName] office ... @" $starttime -ForegroundColor red -BackgroundColor blue


    $path = Split-Path -parent "C:\PowerShell\*.*" 
	$pathexist = Test-Path -Path $path 
    If ($pathexist -eq $false) 
    {New-Item -type directory -Path $path} 
     
 
    $csvreportfile = $path + "\[fileName].csv" #overwrite the *.csv if it already exists

    #import the ActiveDirectory Module 
    Import-Module ActiveDirectory 
     
 #   Get-ADUser -Properties * -Filter  { Office -eq "[officeName]" } |
 Get-ADUser -Properties * -Filter *|
                  Select-Object @{Label = "First Name";Expression = {$_.GivenName}},  
                  @{Label = "Last Name";Expression = {$_.Surname}}, 
                  @{Label = "Display Name";Expression = {$_.DisplayName}}, 
                  @{Label = "Logon Name";Expression = {$_.sAMAccountName}}, 
                  @{Label = "City";Expression = {$_.City}}, 
                  @{Label = "Job Title";Expression = {$_.Title}}, 
                  @{Label = "Company";Expression = {$_.Company}}, 
                  @{Label = "Department";Expression = {$_.Department}}, 
                  @{Label = "Office";Expression = {$_.OfficeName}}, 
                  @{Label = "Phone";Expression = {$_.telephoneNumber}}, 
                  @{Label = "Email";Expression = {$_.Mail}}, 
                  @{Label = "Pager";Expression = {$_.pager}}, 
          		  @{Label = "Manager";Expression = {%{(Get-AdUser $_.Manager -Properties DisplayName).DisplayName}}},
                  @{Label = "Account Status";Expression = {if (($_.Enabled -eq 'TRUE')  ) {'Enabled'} Else {'Disabled'}}}|
               
                   
                  #Export CSV report 
                  Export-Csv -Path $csvreportfile -NoTypeInformation     

$endtime = Get-Date -Format HH:mm:ss
write-host "----------> Done! @" $endtime -ForegroundColor green -BackgroundColor blue

}
##########################################################################END##########################################################################################