# Using SQL query code from https://tsql.tech/how-to-read-data-from-oracle-database-via-powershell-without-using-odbc-or-installing-oracle-client-and-import-it-to-sql-server-too/

#Parameters
$OracleDLLPath = ".\Oracle.ManagedDataAccess.dll"
#The oracle DataSource as you would compile it in TNSNAMES.ORA
$datasource = " (DESCRIPTION = 
                (ADDRESS = (PROTOCOL = TCP)(HOST  = " + $Env:POWERSCHOOL_PROD_DB_IP + ")(PORT = "+ $Env:POWERSCHOOL_PROD_DB_PORT +"))
                (CONNECT_DATA = 
                (SERVER =  DEDICATED)
                (SERVICE_NAME = " + $Env:POWERSCHOOL_PROD_DB_NAME + ")
                (FAILOVER_MODE = (TYPE = SELECT)
                (METHOD =  BASIC)
                (RETRIES = 180)
                (DELAY = 5))))"
$username = $Env:POWERSCHOOL_READ_USER # get the username of read-only account from environment variables
$password = $Env:POWERSCHOOL_DB_PASSWORD # get the password from environment variable

#Load Required Types and modules
Add-Type -Path $OracleDLLPath
Import-Module SqlServer
Import-Module ActiveDirectory

# Clear out log file from previous run
Clear-Content -Path .\studentSyncLog.txt
# repadmin.exe /showrepl *
repadmin.exe /syncall D118-DIST-OFF /Aed # synchronize the controllers so they all have updated data
# break

#Create the connection string
$connectionstring = 'User Id=' + $username + ';Password=' + $password + ';Data Source=' + $datasource 
#Create the connection object
$con = New-Object Oracle.ManagedDataAccess.Client.OracleConnection($connectionstring)

# make a query to find a list of schools
$querySchools = "SELECT name, school_number, abbreviation FROM schools ORDER BY school_number" 

#Create a command and configure it
$cmd = $con.CreateCommand()
$cmd.CommandText = $querySchools
$cmd.CommandTimeout = 3600 #Seconds
$cmd.FetchSize = 10000000 #10MB

#Creates a data adapter for the command
$da = New-Object Oracle.ManagedDataAccess.Client.OracleDataAdapter($cmd);
#The Data adapter will fill this DataTable
$schools = New-Object System.Data.DataTable # each row is one entry, the columns can be accessed with []
#Only here the query is sent and executed in Oracle 
[void]$da.fill($schools)

$spacer = "--------------------------------------------------" # just a spacer for printing easily repeatedly
$constantOU = $Env:CONSTANT_OU_ENDING # define the constant parts of our AD OU structure. Essentially everything after our building level
$defaultPassword = ConvertTo-SecureString $Env:AD_NEW_STUDENT_PASSWORD -AsPlainText -Force  # define the default password used for new accounts

$badNames = 'Use', 'Training1','Trianing2','Trianing3','Trianing4','Planning','Admin','Nurse','User', 'Use ', 'Test', 'Testtt', 'Do Not', 'Do', 'Not', 'Tbd', 'Lunch', 'Formbuilder', 'Human', 'Teststudent' # define list of names to ignore

# define our district wide employee AD groups
$papercutGroup = "Papercut Students Group"
# find the members of these district wide groups so we only have to do it once and then can reference them later
# Get-ADGroupMember has a limit of 5000 results for users, since we have more we need to get the group properties, pipe the member properties to a find user and then select those samAccountNames. Slow but avoids the limit
$papercutStudentMembers = Get-ADGroup $papercutGroup -Properties Member | Select-Object -ExpandProperty Member | Get-ADUser | Select-Object sAMAccountName | ForEach-Object {$_.sAMAccountName}
# $papercutStudentMembers | Out-File -FilePath .\studentSyncLog.txt -Append # debug member list by printing

# define hashtable for converting the integer grade_level to the grade string
$gradeLevels = @{-2="PreKindergarten"; -1="PreKindergarten"; 0="Kindergarten"; 1="1st"; 2="2nd"; 3="3rd"; 4="4th"; 5="5th"; 6="6th"; 7="7th"; 8="8th"; 9="9th"; 10="10th"; 11="11th"; 12="12th"}


foreach ($school in $Schools)
{
    $schoolName = $school[0].ToString().ToUpper()
    $schoolNum = $school[1]
    $schoolAbbrev = $school[2]
    $OUPath = "OU=Students,OU=$schoolName,$constantOU"
    $schoolInfo = "STARTING BUILDING: $schoolName | $schoolNum | $OUPath  | $schoolAbbrev"

    # print out a space line and the school info header to console and log file
    Write-Output $spacer
    Write-Output $spacer | Out-File -FilePath .\studentSyncLog.txt -Append
    Write-Output $schoolInfo 
    Write-Output $schoolInfo | Out-File -FilePath .\studentSyncLog.txt -Append
    Write-Output $spacer
    Write-Output $spacer | Out-File -FilePath .\studentSyncLog.txt -Append

    # get the members of the student groups at the current building, for reference in each user without querying every time. Ignoring buildings where these groups do not exist
    if (($schoolAbbrev -ne "O-HR") -and ($schoolAbbrev -notlike "SU?") -and ($schoolAbbrev -notlike "DNU *") -and ($schoolAbbrev -ne "Graduated Students") -and ($schoolAbbrev -ne "AUX") -and ($schoolAbbrev -ne "TRAN") -and ($schoolAbbrev -ne "MNT") -and ($schoolAbbrev -ne "PRE") -and ($schoolAbbrev -notlike "* OFF") -and ($schoolAbbrev -ne "CO"))
    {
        $schoolStudentGroup = $schoolAbbrev + " Students"
        $schoolStudentMembers = Get-ADGroupMember -Identity $schoolStudentGroup -Recursive | Select-Object -ExpandProperty samAccountName
    }
    # create a new query to find the users in the current building
    $userQuery = "SELECT last_name, first_name, student_number, grade_level, enroll_status, classof FROM students WHERE schoolid = $schoolNum ORDER BY student_number"
    #Use the query in our command and get SQL results
    $cmd.CommandText = $userQuery
    $da = New-Object Oracle.ManagedDataAccess.Client.OracleDataAdapter($cmd);
    #The Data adapter will fill this DataTable
    $resultSet = New-Object System.Data.DataTable # each row is one entry, the columns can be accessed with []
    #Only here the query is sent and executed in Oracle 
    [void]$da.fill($resultSet)
    #Close the connection
    $con.Close()
    foreach ($result in $resultSet)
    {
        $lastName = $result[0].ToLower() # take the last name and convert to all lower case
        $lastName = (Get-Culture).TextInfo.ToTitleCase($lastName) # take the last name all lowercase string and now convert to title case
        $firstName = $result[1].ToLower()
        $firstName = (Get-Culture).TextInfo.ToTitleCase($firstName) # take the last name all lowercase string and now convert to title case
        $studentNumber = $result[2]
        $grade = [int]$result[3]
        $gradeString = $gradeLevels.$grade
        $status = $result[4] # -2=Inactive, -1=Pre-registered, 0=Currently Enrolled, 1=inactive, 2=transferred out, 3=graduated, 4=imported as historical, other=inactive
        $gradyear = $result[5]
        $OUPath = "OU=$gradeString,OU=Students,OU=$schoolName,$constantOU" # set their OU path including the grade level sub-OUs
        if (($badNames -notcontains $firstName) -and ($badNames -notcontains $lastName))
        {
            $email = [string]$studentNumber + "@d118.org"
            $displayName = $firstName + " " + $lastName
            $classOf = "Class of " + $gradyear
            $description = $lastName + ", " + $firstName
            $samAccountName = [string]$studentNumber
            $newHomedirectory = $env:STUDENT_SHARED_DRIVE_BASE_PATH + $samAccountName
            $userInfo = "INFO: Processing User: First: $firstName | Last: $lastName | ID: $studentNumber | Status: $status | Grade: $grade | Grade String: $gradeString"
            Write-Output $userInfo
            $userInfo | Out-File -FilePath .\studentSyncLog.txt -Append # output to the studentSyncLog.txt file
            if (($status -eq -1) -or ($status -eq 0)) # process the active students
            {
                $adUser = Get-ADUser -Filter {sAMAccountName -eq $samAccountName} -Properties description,homedirectory,office,mail,displayname
                # If we find a match for the student number samName, update all their info
                if ($adUser)
                {
                    $currentSamAccountName = $adUser.SamAccountName
                    $currentFullName = $adUser.name
                    $message = "  Student $studentNumber already exists under samname $currentSamAccountName, object full name $currentFullName. Updating any info"
                    Write-Output $message # write to console
                    $message | Out-File -FilePath .\studentSyncLog.txt -Append # write to log file

                    # Check to make sure their user account is enabled
                    if (!$adUser.Enabled)
                    {
                        $message = "      ACTION: ENABLE: Enabling user $currentSamAccountName - $uDCID - $email"
                        Write-Output $message # write to console
                        $message | Out-File -FilePath .\studentSyncLog.txt -Append # write to log file 
                        Enable-ADAccount $adUser # enables the selected account
                    }

                    # Check to see if their name has changed, update the name fields
                    if (($firstName -cne $adUser.GivenName) -or ($lastName -cne $adUser.Surname) )
                    {
                        $currentFirst = $adUser.GivenName
                        $currentLast = $adUser.Surname
                        $message = "      ACTION: NAME: User $studentNumber changed names, updating from $currentFirst $currentLast to $firstName $lastName"
                        Write-Output $message # write to console
                        $message | Out-File -FilePath .\studentSyncLog.txt -Append # write to log file
                        Set-ADUser $adUser -GivenName $firstName -Surname $lastName
                    }

                    # Check to make sure their display name is correct
                    if ($adUser.displayname -cne $displayName)
                    {
                        $currentDisplayName = $adUser.displayname
                        $message = "      ACTION: DISPLAY NAME: User $studentNumber changed diplay names, updating from $currentDisplayName to $displayName"
                        Write-Output $message # write to console
                        $message | Out-File -FilePath .\studentSyncLog.txt -Append # write to log file
                        Set-ADUser $adUser -DisplayName $displayName
                    }

                    # Check to see if their email is correct
                    if ($email -ne $adUser.mail)
                    {
                        $oldEmail = $adUser.mail
                        $message = "      ACTION: EMAIL: User $studentNumber has had their email change from $oldEmail to $email, changing"
                        Write-Output $message # write to console
                        $message | Out-File -FilePath .\studentSyncLog.txt -Append # write to log file 
                        Set-ADUser $adUser -EmailAddress $email -UserPrincipalName $email # update the user's email and principal name which is also their email
                    }

                    # Check to see if their description is correct
                    if ($description -cne $adUser.description)
                    {
                        $oldDescription = $adUser.description
                        $message = "      ACTION: DESCRIPTION: User $studentNumber - has had their description change from $oldDescription to $description, changing"
                        Write-Output $message # write to console
                        $message | Out-File -FilePath .\studentSyncLog.txt -Append # write to log file 
                         Set-ADUser $adUser -Description $description # update the user's description
                    }

                    # Check to see if their "office" is correct which holds their class of 20xx info
                    if ($classOf -cne $adUser.office)
                    {
                        $oldOffice = $adUser.office
                        $message = "      ACTION: OFFICE: User $studentNumber - has had their office/class of change from $oldOffice to $classOf, changing"
                        Write-Output $message # write to console
                        $message | Out-File -FilePath .\studentSyncLog.txt -Append # write to log file 
                         Set-ADUser $adUser -Office $classOf # update the user's description
                    }

                    # Check to see if they are in the right OU, move them if not
                    $properDistinguished = "CN=$currentFullName,$OUPath" # construct what the desired/proper distinguished name should be based on their samaccount name and the OU they should be in
                    if ($properDistinguished -ne $adUser.DistinguishedName)
                    {
                        $currentDistinguished =  $adUser.DistinguishedName
                        try
                        {
                            $message = "      ACTION: OU: User $studentNumber NOT in correct OU, moving from $currentDistinguished to $properDistinguished"
                            Write-Output $message # write to console
                            $message | Out-File -FilePath .\studentSyncLog.txt -Append # write to log file 
                            Move-ADObject $adUser -TargetPath $OUPath # moves the targeted AD user account to the correct OU
                        }
                        catch 
                        {
                            $message =  "          ERROR: Could not move $studentNumber to $OUPath"
                            Write-Output $message # write to console
                            $message | Out-File -FilePath .\studentSyncLog.txt -Append # write to log file 
                            Write-Output $_ # write out the actual error
                            $_ | Out-File -FilePath .\studentSyncLog.txt -Append
                        }
                    }

                    # Check to ensure the user is a member of the papercut student group
                    if ($papercutStudentMembers -notcontains $adUser.samAccountName)
                    {
                        $message =  "      ACTION: GROUP: User $currentSamAccountName is not a member of $papercutGroup, will add them"
                        Write-Output $message # write to console
                        $message | Out-File -FilePath .\studentSyncLog.txt -Append # write to log file
                        try 
                        {
                            Add-ADGroupMember -Identity $papercutGroup -Members $adUser.samAccountName # add the user to the group
                        }
                        catch
                        {
                            $message = "     ERROR: Could not add $currentSameAccountName to $papercutGroup"
                            Write-Output $message # write to console
                            $message | Out-File -FilePath .\studentSyncLog.txt -Append # write to log file
                        }
                    }

                    # Check to ensure the user is a member of the school student group, ignoring all the non-student buildings
                    if (($schoolAbbrev -ne "O-HR") -and ($schoolAbbrev -notlike "SU?") -and ($schoolAbbrev -notlike "DNU *") -and ($schoolAbbrev -ne "Graduated Students") -and ($schoolAbbrev -ne "AUX") -and ($schoolAbbrev -ne "TRAN") -and ($schoolAbbrev -ne "MNT") -and ($schoolAbbrev -ne "PRE") -and ($schoolAbbrev -notlike "* OFF") -and ($schoolAbbrev -ne "CO"))
                    {
                        if ($schoolStudentMembers -notcontains $adUser.samAccountName)
                        {
                            $message =  "      ACTION: GROUP: User $currentSamAccountName is not a member of $schoolStudentGroup, will add them"
                            Write-Output $message # write to console
                            $message | Out-File -FilePath .\studentSyncLog.txt -Append # write to log file
                            try 
                            {
                                Add-ADGroupMember -Identity $schoolStudentGroup -Members $adUser.samAccountName # add the user to the group
                            }
                            catch
                            {
                                $message = "     ERROR: Could not add $currentSameAccountName to $schoolStudentGroup"
                                Write-Output $message # write to console
                                $message | Out-File -FilePath .\studentSyncLog.txt -Append # write to log file
                                Write-Output $_ # write out the actual error
                                $_ | Out-File -FilePath .\studentSyncLog.txt -Append
                            }
                        }
                    }

                    # Check to see if they have a homedrive populated, if not we want to assign them one
                    if([string]::IsNullOrEmpty($adUser.homedirectory))
                    {
                        $newHomedirectory = $env:STUDENT_SHARED_DRIVE_BASE_PATH + $currentSamAccountName
                        $message = "      ACTION: HOMEDRIVE: User $currentSamAccountName does not have a home directory mapped, will be assigned one at $newHomedirectory"
                        Write-Output $message # write to console
                        $message | Out-File -FilePath .\studentSyncLog.txt -Append # write to log file
                        try 
                        {
                            Set-ADUser $adUser -HomeDirectory $newHomedirectory -HomeDrive "H:" # set their home drive to be H: and mapped to the directory constructed from their building and name
                        }
                        catch 
                        {
                            $message =  "          ERROR: Could not map homedrive for $currentSamAccountName to $newHomedirectory"
                            Write-Output $message # write to console
                            $message | Out-File -FilePath .\studentSyncLog.txt -Append # write to log file 
                            Write-Output $_ # write out the actual error
                            $_ | Out-File -FilePath .\studentSyncLog.txt -Append
                        }
                    }

                    # Check to see if the "Full Name" is the same as their samAccountName, if not, change it to match. Do this last otherwise it will break the other operations due to the object being renamed
                    if ($currentFullName -ne $currentSamAccountName)
                    {
                        $message = "      ACTION: FULL NAME: Updating user $uDCID's 'full name' from $currentFullName to $currentSamAccountName"
                        Write-Output $message # write to console
                        $message | Out-File -FilePath .\studentSyncLog.txt -Append # write to log file 
                        Rename-ADObject $adUser -NewName $currentSamAccountName
                    }
                }
                # If we do not find a match for our current student number, try to create the account
                else 
                {
                    $message =  "  ACTION: CREATION: Student $studentNumber does not exist, will try to create them as $samAccountName in $OUPath"
                    Write-Output $message # write to console
                    $message | Out-File -FilePath .\studentSyncLog.txt -Append # write to log file 
                    # Account creation
                    try
                    {
                        New-ADUser -SamAccountName $samAccountName -Name $samAccountName -DisplayName $displayName -GivenName $firstName -Surname $lastName -EmailAddress $email -UserPrincipalName $email -Path $OUPath -Office $classOf -Description $description -AccountPassword $defaultPassword -HomeDrive "H:" -HomeDirectory $newHomedirectory -ChangePasswordAtLogon $false -Enabled $true
                    }
                    catch
                    {
                        $message =  "       ERROR: Could not create user $samAccountName"
                        Write-Output $message # write to console
                        $message | Out-File -FilePath .\studentSyncLog.txt -Append # write to log file 
                        Write-Output $_ # write out the actual error
                        $_ | Out-File -FilePath .\studentSyncLog.txt -Append
                    }
                }
            }
            # Start of inactive student block, should be disabled and moved to suspended accounts OU. Otherwise do not update any info
            else 
            {
                # check to see if they are a graduated student as they have a different sub-OU
                if (($grade -eq 99) -or ($status -eq 3)) 
                {
                    $OUPath = "OU=Graduated Students,OU=SUSPENDED ACCOUNTS,$ConstantOU"
                }
                else 
                {
                    $OUPath = "OU=Students,OU=SUSPENDED ACCOUNTS,$ConstantOU"
                }
                $adUser = Get-ADUser -Filter {sAMAccountName -eq $samAccountName}
                # if we find a user with a matching student number, just update their info
                if ($adUser)
                { 
                    $currentFullName = $adUser.name
                    $properDistinguished = "CN=$currentFullName,$OUPath"
                    $currentSamAccountName = $adUser.SamAccountName
                    $message = "  User $studentNumber already exists under $currentSamAccountName, ensuring they are suspended"
                    Write-Output $message # write to console
                    $message | Out-File -FilePath .\studentSyncLog.txt -Append # write to log file 
                    # check to see if the account is enabled, if so we need to disable it
                    if ($adUser.Enabled)
                    {
                        try 
                        {
                            $message = "      ACTION: SUSPENDED DISABLE: Disabling user $currentSamAccountName"
                            Write-Output $message # write to console
                            $message | Out-File -FilePath .\studentSyncLog.txt -Append # write to log file
                            Disable-ADAccount $adUser # disables the selected account
                        }
                        catch 
                        {
                            $message =  "          ERROR: Could not suspend $currentSamAccountName"
                            Write-Output $message # write to console
                            $message | Out-File -FilePath .\studentSyncLog.txt -Append # write to log file 
                            Write-Output $_ # write out the actual error
                            $_ | Out-File -FilePath .\studentSyncLog.txt -Append
                        }
                        
                    }
                    # check and see if their account is in the correct suspended users OU, if not move them
                    if ($properDistinguished -ne $adUser.DistinguishedName)
                    {
                        try 
                        {
                            $message = "      ACTION: SUSPENDED OU: Moving user $currentSamAccountName to $OUPath"
                            Write-Output $message # write to console
                            $message | Out-File -FilePath .\studentSyncLog.txt -Append # write to log file 
                            Move-ADObject $adUser -TargetPath $OUPath # moves the targeted AD user account to the correct suspended accounts OU
                        }
                        catch 
                        {
                            $message =  "          ERROR: Could not move  $currentSamAccountName to $OUPath"
                            Write-Output $message # write to console
                            $message | Out-File -FilePath .\studentSyncLog.txt -Append # write to log file 
                            Write-Output $_ # write out the actual error
                            $_ | Out-File -FilePath .\studentSyncLog.txt -Append
                        }
                    } 
                }
                else
                {
                    $message = "  WARNING: Found inactive user $studentNumber without matching AD account."
                    Write-Output $message # write to console
                    $message | Out-File -FilePath .\studentSyncLog.txt -Append # write to log file 
                }
            }
        }
        else # otherwise if their name was found in the bad names list, just give a warning
        {
            $message = "INFO: found user matching name in bad names list: $firstName $LastName"
            Write-Output $message
            $message | Out-File -FilePath .\studentSyncLog.txt -Append
        }
    }
}