$Employee_Data = Import-CSV -Path "UNC_PATH"
######Uncomment for script testing######## "C:\PowerShell_Test\Automated_Group_Test.csv"

$Employee_DC = "DOMAIN_CONTROLLER"

$Employee_Array = @()

$Non_Employee_Group1 = "Student_License"
$Non_Employee_Group1_Members = Get-ADGroupMember -Identity $Non_Employee_Group1 -Server $Employee_DC -Recursive | Select -ExpandProperty samaccountname
$Non_Employee_Group2 = "Alumni_License"
$Non_Employee_Group2_Members = Get-ADGroupMember -Identity $Non_Employee_Group2 -Server $Employee_DC -Recursive | Select -ExpandProperty samaccountname
$Match = "*EMAIL_DOMAIN*"

    Foreach ($Employee in $Employee_Data)
        {
            Import-Module ActiveDirectory

            $Employee_First = ($Employee.FIRST_NAME)
            $Employee_Last = ($Employee.LAST_NAME)
            $Employee_Type = ($Employee.EMPLOYEE_TYPE)
            $Employee_Group = ($Employee.GROUP_NAME)
            $Employee_Email = ($Employee.WORK_EMAIL)
            $Employee_Status = ($Employee.EMPLOYMENT_STATUS)
            $Employee_DisplayName = "$Employee_Last, $Employee_First"
            $Employee_Name = "$Employee_First $Employee_Last"

            If ($Employee_Email -ne "")
                {
                    $Find_Employee_AD = Get-ADUser -Filter {(mail -eq $Employee_Email) -and (Enabled -eq $true)} -Server $Employee_DC -Properties samaccountname, distinguishedname | Select samaccountname, distinguishedname
                    $Employee_DN = ($Find_Employee_AD.distinguishedname)
                    $Employee_SAM = ($Find_Employee_AD.samaccountname)
                        
                        If ($Non_Employee_Group1_Members -notcontains $Employee_SAM)

                        {
                            If ($Non_Employee_Group2_Members -notcontains $Employee_SAM)
                            
                            {
                                If ($Employee_Email -like $Match)
                                    
                                    {
                
            Switch ($Employee_Type)
                {
                    ("Administrative")
                        {
                            $Group1_Name = "Administration"
                            
                            $Group1_Members = Get-ADGroupmember -Identity $Group1_Name -Server $Employee_DC -Recursive | Select -ExpandProperty samaccountname
                                
                                if ($Group1_Members -notcontains $Employee_SAM)
                                    
                                    {
                                         Write-Host "Add $Employee_Email to $Group1_Name" -Background "Green" -Foreground "Black"
                                         
                                         #Add-ADGroupMember -Identity $Group1_Name -Members $Employee_DN -Verbose
                                         
                                         $Employee_Add = @"
                                         Username,GroupName
                                         "",""
                                         $Employee_Email,$Group1_Name
"@ | ConvertFrom-CSV
                                        $Employee_Array += $Employee_Add
                                    }
                                Else
                                    {
                                        Write-Host "$Employee_Email is a member of Group: $Group1_Name" -Background "White" -Foreground "Black"
                                    }
                        }
                    ("Support Staff") 
                        {

                            $Group2_Name = "Support Staff"

                            $Group2_Members = Get-ADGroupmember -Identity $Group2_Name -Recursive | Select -ExpandProperty samaccountname

                                if ($Group2_Members -notcontains $Employee_SAM)
                                    
                                    {
                                         Write-Host "Add $Employee_Email to $Group2_Name" -Background "Green" -Foreground "Black"
                                         
                                         #Add-ADGroupMember -Identity $Group2_Name -Members $Employee_DN -Verbose
                                         

                                         $Employee_Add = @"
                                         Username,GroupName
                                         "",""
                                         $Employee_Email,$Group2_Name
"@ | ConvertFrom-CSV
                                        $Employee_Array += $Employee_Add
 
                                    }
                                Else
                                    {
                                        Write-Host "$Employee_Email is a member of Group: $Group2_Name" -Background "White" -Foreground "Black"
                                    } 
                        }
                    ("Faculty") 
                        { 
                            $Group3_Name = "Faculty"

                            $Group3_Members = Get-ADGroupmember -Identity $Group3_Name -Recursive | Select -ExpandProperty samaccountname

                                if ($Group3_Members -notcontains $Employee_SAM)
                                    
                                    {
                                         Write-Host "Add $Employee_Email to $Group3_Name" -Background "Green" -Foreground "Black"
                                         
                                         #Add-ADGroupMember -Identity $Group3_Name -Members $Employee_DN -Verbose
                                         

                                         $Employee_Add = @"
                                         Username,GroupName
                                         "",""
                                         $Employee_Email,$Group3_Name
"@ | ConvertFrom-CSV
                                        $Employee_Array += $Employee_Add

                                    }
                                Else
                                    {
                                        Write-Host "$Employee_Email is a member of Group: $Group3_Name" -Background "White" -Foreground "Black"
                                    }
                        }
                }
            Switch ($Employee_Status)
                {
                    ("Full-Time")
                        {
                                                            
                            $Group5_Name = "FT-Employees"

                            $Group5_Members = Get-ADGroupmember -Identity $Group5_Name -Recursive | Select -ExpandProperty samaccountname

                                if ($Group5_Members -notcontains $Employee_SAM)
                                    
                                    {
                                         Write-Host "Add $Employee_Email to $Group5_Name" -Background "Green" -Foreground "Black"
                                         
                                         #Add-ADGroupMember -Identity $Group5_Name -Members $Employee_DN -Verbose
                                         
                                         $Employee_Add = @"
                                         Username,GroupName
                                         "",""
                                         $Employee_Email,$Group5_Name
"@ | ConvertFrom-CSV
                                        $Employee_Array += $Employee_Add

                                    }
                                Else
                                    {
                                        Write-Host "$Employee_Email is a member of Group: $Group5_Name" -Background "White" -Foreground "Black"
                                    }

                }
                    ("Part-Time")
                        {

                            $Group6_Name = "PT-Employees"
                            
                            $Group6_Members = Get-ADGroupmember -Identity "PT-Employees" -Recursive | Select -ExpandProperty samaccountname

                                if ($Group6_Members -notcontains $Employee_SAM)
                                    
                                    {
                                         Write-Host "Add $Employee_Email to $Group6_Name" -Background "Green" -Foreground "Black"
                                         
                                         #Add-ADGroupMember -Identity $Group6_Name -Members $Employee_DN -Verbose
                                         

                                         $Employee_Add = @"
                                         Username,GroupName
                                         "",""
                                         $Employee_Email,$Group6_Name
"@ | ConvertFrom-CSV
                                        $Employee_Array += $Employee_Add

                                    }
                                Else
                                    {
                                        Write-Host "$Employee_Email is a member of Group: $Group6_Name" -Background "White" -Foreground "Black"
                                    }

                }
                    ("Occasional-Staff")
                        {
                            $Group7_Name = "Occasional-Staff"

                            $Group7_Members = Get-ADGroupmember -Identity $Group7_Name -Recursive | Select -ExpandProperty samaccountname

                                if ($Group7_Members -notcontains $Employee_SAM)
                                    
                                    {
                                         Write-Host "Add $Employee_Email to $Group7_Name" -Background "Green" -Foreground "Black"
                                         
                                         #Add-ADGroupMember -Identity $Group7_Name -Members $Employee_DN -Verbose
                                         

                                         $Employee_Add = @"
                                         Username,GroupName
                                         "",""
                                         $Employee_Email,$Group7_Name
"@ | ConvertFrom-CSV
                                        $Employee_Array += $Employee_Add

                                    }
                                Else
                                    {
                                        Write-Host "$Employee_Email is a member of Group: $Group7_Name" -Background "White" -Foreground "Black"
                                    }

                        }
                    }
                }
            }
        }
    }
            If ($Employee_Email -eq "")
                {
                $Find_Employee_NE = Get-ADUser -Filter {(DisplayName -eq $Employee_DisplayName) -and (Enabled -eq $true)} -Server $Employee_DC -Properties samaccountname, distinguishedname, mail | Select samaccountname, distinguishedname, mail
                $Employee_NE_DN = ($Find_Employee_NE.distinguishedname)
                $Employee_NE_SAM = ($Find_Employee_NE.samaccountname)
                $Employee_NE_Mail = ($Find_Employee_NE.mail)

                 If ($Non_Employee_Group1_Members -notcontains $Employee_NE_SAM)

                        {
                            If ($Non_Employee_Group2_Members -notcontains $Employee_NE_SAM)
                            
                            {

                                Switch ($Employee_Type)
                                
                                {
                                    ("Retired")
                                        {

                                            $Group8_Name = "Retired"

                                                $Group8_Members = Get-ADGroupmember -Identity $Group8_Name -Recursive | Select -ExpandProperty samaccountname

                                                    if ($Group8_Members -notcontains $Employee_NE_SAM)
                                    
                                                        {
                                                            Write-Host "Add $Employee_DisplayName to $Group8_Name" -Background "Green" -Foreground "Black"

                                                            #Add-ADGroupMember -Identity $Group8_Name -Members $Employee_DN -Verbose

                                                            $Employee_Add = @"
                                                            Username,GroupName
                                                            "",""
                                                            $Employee_Name,$Group8_Name
"@ | ConvertFrom-CSV
                                                            $Employee_Array += $Employee_Add

                                                        }
                                                    Else
                                                        {
                                                            Write-Host "$Employee_Name is a member of Group: $Group8_Name" -Background "White" -Foreground "Black"
                                                        }
                                            }
                                }
            
                        }
                    }
                }
            }        
###########################################
#             HTML Varables               #
###########################################
$Employee_HTML = "<style>"
$Employee_HTML = $Employee_HTML + "BODY{background-color:white;}"
$Employee_HTML = $Employee_HTML + "TABLE{border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}"
$Employee_HTML = $Employee_HTML + "TH{border-width: 1px;padding: 0px;border-style: solid;border-color: black;background-color:thistle}"
$Employee_HTML = $Employee_HTML + "TD{border-width: 1px;padding: 0px;border-style: solid;border-color: black;background-color:palegoldenrod}"
$Employee_HTML = $Employee_HTML + "</style>"
$Employee_Email_Body = $Employee_Array | ConvertTo-Html -Head $Employee_HTML -Body "<H2>New Group Members</H2>" | Out-String
###########################################
#            SMTP Variables               #
###########################################           
$From = "FROM_ADDRESS"
$To = "TO_ADDRESS"
$Subject = "AD Group Membership"
$SMTPServer = "SMTP_SERVER"
##########################################
#              Send the Mail             #
##########################################
If ($Employee_Array.Count -gt 1) 
    {
     #Send-MailMessage -From $From -to $To -Subject $Subject -Body $Employee_Email_Body -BodyAsHtml -SmtpServer $SMTPServer 
    
    $Employee_uri = 'TEAMS_URI'

    # Time
    $Employee_Time = get-date -format "MM-dd-yyyy HH:mm"

    # Script Name
    $Employee_Script_Name = "Automated_Groups.ps1"

    # Script Name
    $Employee_RealName = "Automated_Groups"

    # these values would be retrieved from or set by an application

    $Employee_Teams_Body = ConvertTo-Json -Depth 4 @{
        title    = "$Employee_Script_Name Completed Successfully"
        text	 = " "
        sections = @(
            @{
            activityTitle    = 'Automated_Groups Completed'
            activitySubtitle = 'Automated Group Membership Updates'
            #activityText	 = ' '
            activityImage    = 'PNG IMAGE' # this value would be a path to a nice image you would like to display in notifications
        },
        @{
         title = '<h2 style=color:blue;>Script Details'
         facts = @(
            @{
              name  = 'ScriptName'
              value = $Employee_Script_Name
            }
            @{
              name  = 'Finished'
              value = $Employee_Time
            },
            @{
              name = 'Group Additionals'
              value = $Employee_Email_Body
            }

      )
    }
  )
}
Invoke-RestMethod -uri $Employee_uri -Method Post -body $Employee_Teams_Body -ContentType 'application/json'
    }
