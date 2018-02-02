$LocalVersion = "0.5"
cls

Import-Module ActiveDirectory | Out-Null

$AssemblyDictionary = (
"Accessibility",
"Anonymously Hosted DynamicMethods Assembly",
"MetadataViewProxies_b32ae82b-c6f6-49e4-84e0-e60a2410395f",
"Microsoft.GeneratedCode",
"Microsoft.Management.Infrastructure",
"Microsoft.PowerShell.Commands.Utility",
"Microsoft.PowerShell.Editor",
"Microsoft.PowerShell.GPowerShell",
"Microsoft.PowerShell.GraphicalHost",
"Microsoft.PowerShell.ISECommon",
"Microsoft.PowerShell.Security",
"mscorlib",
"powershell_ise",
"PresentationCore",
"PresentationFramework",
"PresentationFramework.Aero2",
"PresentationFramework-SystemCore",
"PresentationFramework-SystemData",
"PresentationFramework-SystemXml",
"SMDiagnostics",
"System",
"System.ComponentModel.Composition",
"System.Configuration",
"System.Configuration.Install",
"System.Core",
"System.Data",
"System.Diagnostics.Tracing",
"System.DirectoryServices",
"System.Drawing",
"System.Management",
"System.Management.Automation",
"System.Numerics",
"System.Runtime.Serialization",
"System.ServiceModel.Internals",
"System.Transactions",
"System.Windows.Forms",
"System.Xaml",
"System.Xml",
"UIAutomationProvider",
"UIAutomationTypes",
"WindowsBase"
)

Foreach($Assembly in $AssemblyDictionary){

    $tsttemp = $Assembly.ToString()

    [reflection.assembly]::loadwithpartialname("$tsttemp") | Out-Null
}

$DateFormat = "yyyy-MM-dd HH:mm:ss"
$Date = [DateTime]::Now.ToString($dateFormat)

$TimestampFormat = "HH:mm:ss"
$Timestamp = [DateTime]::Now.ToString($TimestampFormat)

$LogDateFormat = "dd-MM-yyyy"
$LogDate = [DateTime]::Today.ToString($LogDateFormat)

$RootPath = "C:\Temp\"
$AppPath = "C:\Temp\AD-Tool\"
$LogFile = "log-$LogDate.txt"

#
# # # # # Form Structure Start # # # # #

#region Form structuring

$frmInitialScreen = New-Object system.Windows.Forms.Form
$frmInitialScreen.Text = 'AD-Tool Audit & Copy 2017'
$frmInitialScreen.Width = 640
$frmInitialScreen.Height = 710
$frmInitialScreen.Location = New-Object System.Drawing.Size(200,300)
$frmInitialScreen.StartPosition = "manual"
$frmInitialScreen.MaximumSize = New-Object system.drawing.size(640, 900)
$frmInitialScreen.MinimumSize = New-Object system.drawing.size(640, 710)


$lblSearchItems = New-Object system.windows.Forms.Label
$lblSearchItems.Text = 'Search:'
$lblSearchItems.AutoSize = $true
$lblSearchItems.Width = 25
$lblSearchItems.Height = 10
$lblSearchItems.location = New-Object system.drawing.size(5,20)
$lblSearchItems.Font = "Microsoft Sans Serif,10"
$lblSearchItems.Anchor = "Left, Top"
$frmInitialScreen.controls.Add($lblSearchItems)

$lblLeftSideBox = New-Object system.windows.Forms.Label
$lblLeftSideBox.Text = 'Source:'
$lblLeftSideBox.AutoSize = $true
$lblLeftSideBox.Width = 25
$lblLeftSideBox.Height = 10
$lblLeftSideBox.location = New-Object system.drawing.size(195,20)
$lblLeftSideBox.Font = "Microsoft Sans Serif,10"
$lblLeftSideBox.Anchor = "Left, Top"
$frmInitialScreen.controls.Add($lblLeftSideBox)

$lblRightSideBox = New-Object system.windows.Forms.Label
$lblRightSideBox.Text = 'Destination:'
$lblRightSideBox.AutoSize = $true
$lblRightSideBox.Width = 25
$lblRightSideBox.Height = 10
$lblRightSideBox.location = New-Object system.drawing.size(418,20)
$lblRightSideBox.Font = "Microsoft Sans Serif,10"
$lblRightSideBox.Anchor = "Left, Top"
$frmInitialScreen.controls.Add($lblRightSideBox)

$lblVerboseOutput = New-Object system.windows.Forms.Label
$lblVerboseOutput.Text = 'Verbose:'
$lblVerboseOutput.AutoSize = $true
$lblVerboseOutput.Width = 25
$lblVerboseOutput.Height = 10
$lblVerboseOutput.location = New-Object system.drawing.size(4,362)
$lblVerboseOutput.Font = "Microsoft Sans Serif,10"
$frmInitialScreen.controls.Add($lblVerboseOutput)

$lblFootnote = New-Object system.windows.Forms.Label
$lblFootnote.Text = "-This will require domain admin to copy users, unlock, or locate lockouts-" 
$lblFootnote.AutoSize = $true
$lblFootnote.Width = 25
$lblFootnote.Height = 10
$lblFootnote.location = New-Object system.drawing.size(4,656) # was 4 450
$lblFootnote.Font = "Microsoft Sans Serif,8"
$lblFootnote.Anchor = "Left, Bottom"
$frmInitialScreen.controls.Add($lblFootnote)

$txtGetFirstGroupInfo = New-Object system.windows.Forms.TextBox
$txtGetFirstGroupInfo.Width = 178
$txtGetFirstGroupInfo.Height = 20
$txtGetFirstGroupInfo.location = New-Object system.drawing.size(5,40)
$txtGetFirstGroupInfo.Font = "Microsoft Sans Serif,10"
$txtGetFirstGroupInfo.TabIndex = 0
$frmInitialScreen.controls.Add($txtGetFirstGroupInfo)
$txtGetFirstGroupInfo_KeyDown = {}
$txtGetFirstGroupInfo_KeyDown = [System.Windows.Forms.KeyEventHandler]{}
$txtGetFirstGroupInfo_KeyDown = [System.Windows.Forms.KeyEventHandler]{
    if ($_.KeyCode -eq 'Enter')
    {
        $btnShowFirstGroups.PerformClick()
    }
}
$txtGetFirstGroupInfo.Anchor = "Left, Top"
$txtGetFirstGroupInfo.Add_KeyDown($txtGetFirstGroupInfo_KeyDown)


$txtGetSecondGroupInfo = New-Object system.windows.Forms.TextBox
$txtGetSecondGroupInfo.Width = 178
$txtGetSecondGroupInfo.Height = 20
$txtGetSecondGroupInfo.location = New-Object system.drawing.size(5,105)
$txtGetSecondGroupInfo.Font = "Microsoft Sans Serif,10"
$txtGetSecondGroupInfo.TabIndex = 1
$frmInitialScreen.controls.Add($txtGetSecondGroupInfo)
$txtGetSecondGroupInfo_KeyDown = {}
$txtGetSecondGroupInfo_KeyDown = [System.Windows.Forms.KeyEventHandler]{}
$txtGetSecondGroupInfo_KeyDown = [System.Windows.Forms.KeyEventHandler]{
    if ($_.KeyCode -eq 'Enter')
    {
        $btnShowSecondGroups.PerformClick()
    }
}
$txtGetSecondGroupInfo.Anchor = "Left, Top"
$txtGetSecondGroupInfo.Add_KeyDown($txtGetSecondGroupInfo_KeyDown)

$txtGetUsernameInfo = New-Object system.windows.Forms.TextBox
$txtGetUsernameInfo.Width = 178
$txtGetUsernameInfo.Height = 20
$txtGetUsernameInfo.location = New-Object system.drawing.size(5,175)
$txtGetUsernameInfo.Font = "Microsoft Sans Serif,10"
$txtGetUsernameInfo.TabIndex = 2
$txtGetUsernameInfo.KeyUp
$frmInitialScreen.controls.Add($txtGetUsernameInfo)
$txtGetUsernameInfo_KeyDown = {}
$txtGetUsernameInfo_KeyDown = [System.Windows.Forms.KeyEventHandler]{}
$txtGetUsernameInfo_KeyDown = [System.Windows.Forms.KeyEventHandler]{
    if ($_.KeyCode -eq 'Enter')
    {
        $btnSearchByUsername.PerformClick()
    }
}
$txtGetUsernameInfo.Anchor = "Left, Top"
$txtGetUsernameInfo.Add_KeyDown($txtGetUsernameInfo_KeyDown)

#Buttons below are vertical

$btnShowFirstGroups = New-Object system.windows.Forms.Button
$btnShowFirstGroups.Text = 'Search Source Group'
$btnShowFirstGroups.Width = 180
$btnShowFirstGroups.Height = 30
$btnShowFirstGroups.location = New-Object system.drawing.size(4,65)
$btnShowFirstGroups.Font = "Microsoft Sans Serif,10"
$btnShowFirstGroups.TabStop = $false
$btnShowFirstGroups.Anchor = "Left, Top"
$frmInitialScreen.controls.Add($btnShowFirstGroups)

$btnShowSecondGroups = New-Object system.windows.Forms.Button
$btnShowSecondGroups.Text = 'Search Destination Group'
$btnShowSecondGroups.Width = 180
$btnShowSecondGroups.Height = 30
$btnShowSecondGroups.location = New-Object system.drawing.size(4,130)
$btnShowSecondGroups.Font = "Microsoft Sans Serif,10"
$btnShowSecondGroups.TabStop = $false
$btnShowSecondGroups.Anchor = "Left, Top"
$frmInitialScreen.controls.Add($btnShowSecondGroups)

$btnSearchByUsername = New-Object system.windows.Forms.Button
$btnSearchByUsername.Text = 'Search By Name'
$btnSearchByUsername.Width = 180
$btnSearchByUsername.Height = 30
$btnSearchByUsername.location = New-Object system.drawing.size(4,200)
$btnSearchByUsername.Font = "Microsoft Sans Serif,10"
$btnSearchByUsername.TabStop = $false
$btnSearchByUsername.Anchor = "Left, Top"
$frmInitialScreen.controls.Add($btnSearchByUsername)

$btnClearListBoxes = New-Object system.windows.Forms.Button
$btnClearListBoxes.Text = 'Clear'
$btnClearListBoxes.Width = 180
$btnClearListBoxes.Height = 30
$btnClearListBoxes.location = New-Object system.drawing.size(4,235)
$btnClearListBoxes.Font = "Microsoft Sans Serif,10"
$btnClearListBoxes.TabIndex = 3
$btnClearListBoxes.Anchor = "Left, Top"
$frmInitialScreen.controls.Add($btnClearListBoxes)


#Buttons below are horizontal

$btnUserInfo = New-Object system.windows.Forms.Button
$btnUserInfo.Text = 'User Info'
$btnUserInfo.Width = 150
$btnUserInfo.Height = 40
$btnUserInfo.location = New-Object system.drawing.size(4,270)
$btnUserInfo.Font = "Microsoft Sans Serif,10"
$btnUserInfo.TabIndex = 6
$frmInitialScreen.controls.Add($btnUserInfo)

$btnShowHelp = New-Object system.windows.Forms.Button
$btnShowHelp.Text = 'Show Help'
$btnShowHelp.Width = 150
$btnShowHelp.Height = 40
$btnShowHelp.location = New-Object system.drawing.size(4,310)
$btnShowHelp.Font = "Microsoft Sans Serif,10"
$btnShowHelp.TabIndex = 13
$frmInitialScreen.controls.Add($btnShowHelp)




$btnShowFirstUsers = New-Object system.windows.Forms.Button
$btnShowFirstUsers.Text = 'Show Group Members'
$btnShowFirstUsers.Width = 150
$btnShowFirstUsers.Height = 40
$btnShowFirstUsers.location = New-Object system.drawing.size(160,270)
$btnShowFirstUsers.Font = "Microsoft Sans Serif,10"
$btnShowFirstUsers.TabIndex = 8
$frmInitialScreen.controls.Add($btnShowFirstUsers)

$btnShowMemberships = New-Object system.windows.Forms.Button
$btnShowMemberships.Text = 'Show Users Groups'
$btnShowMemberships.Width = 150
$btnShowMemberships.Height = 40
$btnShowMemberships.location = New-Object system.drawing.size(160,310)
$btnShowMemberships.Font = "Microsoft Sans Serif,10"
$btnShowMemberships.TabIndex = 9
$frmInitialScreen.controls.Add($btnShowMemberships)




$btnCopyUsersOver = New-Object system.windows.Forms.Button
$btnCopyUsersOver.Text = 'Copy All Users'
$btnCopyUsersOver.Width = 150
$btnCopyUsersOver.Height = 40
$btnCopyUsersOver.location = New-Object system.drawing.size(315,270)
$btnCopyUsersOver.Font = "Microsoft Sans Serif,10"
$btnCopyUsersOver.TabIndex = 10
$frmInitialScreen.controls.Add($btnCopyUsersOver)

$btnCopySelectedUsers = New-Object system.windows.Forms.Button
$btnCopySelectedUsers.Text = 'Copy Selected Users'
$btnCopySelectedUsers.Width = 150
$btnCopySelectedUsers.Height = 40
$btnCopySelectedUsers.location = New-Object system.drawing.size(315,310)
$btnCopySelectedUsers.Font = "Microsoft Sans Serif,10"
$btnCopySelectedUsers.TabIndex = 7
$frmInitialScreen.controls.Add($btnCopySelectedUsers)




$btnGetLockoutBlame = New-Object system.windows.Forms.Button
$btnGetLockoutBlame.Text = 'Lockout Location'
$btnGetLockoutBlame.Width = 150
$btnGetLockoutBlame.Height = 40
$btnGetLockoutBlame.location = New-Object system.drawing.size(470,270)
$btnGetLockoutBlame.Font = "Microsoft Sans Serif,10"
$btnCopyUsersOver.TabIndex = 11
$frmInitialScreen.controls.Add($btnGetLockoutBlame)

$btnUnlockAccount = New-Object system.windows.Forms.Button
$btnUnlockAccount.Text = 'Unlock User'
$btnUnlockAccount.Width = 150
$btnUnlockAccount.Height = 40
$btnUnlockAccount.location = New-Object system.drawing.size(470,310)
$btnUnlockAccount.Font = "Microsoft Sans Serif,10"
$btnCopyUsersOver.TabIndex = 11
$frmInitialScreen.controls.Add($btnUnlockAccount)


#List boxes below

$lstShowFirstGroups = New-Object system.windows.Forms.ListBox
$lstShowFirstGroups.Text = "listView"
$lstShowFirstGroups.Width = 200
$lstShowFirstGroups.Height = 230
$lstShowFirstGroups.location = New-Object system.drawing.point(195,40)
$lstShowFirstGroups.TabIndex = 4
$lstShowFirstGroups.SelectionMode = "One"
$lstShowFirstGroups.Anchor = "Left, Right, Top"
$frmInitialScreen.controls.Add($lstShowFirstGroups)

$lstShowSecondGroups = New-Object system.windows.Forms.ListBox
$lstShowSecondGroups.Text = "listView"
$lstShowSecondGroups.Width = 200
$lstShowSecondGroups.Height = 230
$lstShowSecondGroups.location = New-Object system.drawing.point(419,40)
$lstShowSecondGroups.Anchor = "Right, Left, Top"
$lstShowSecondGroups.TabIndex = 5

$frmInitialScreen.controls.Add($lstShowSecondGroups)

$lstVerboseOutput = New-Object system.windows.Forms.ListBox
$lstVerboseOutput.Text = "listView"
$lstVerboseOutput.Width = 615
$lstVerboseOutput.Height = 270 # was 120
$lstVerboseOutput.location = New-Object system.drawing.point(4,383)
$lstVerboseOutput.TabStop = $false
$lstVerboseOutput.Anchor = "Top, Bottom, Left, Right"
$frmInitialScreen.controls.Add($lstVerboseOutput)

#endregion

# # # # # Form Structure Ended # # # # #
#



#region Connection test

#Initial test to check if the user has access to AD by trying to search the name of the current user.

$ErrorActionPreference = 'continue'
$CurrentUser = [Environment]::UserName
$CurrentMachine = [Environment]::MachineName
$CurrentTime = Get-Date -Format g

$RootCheck = Test-Path $RootPath
$AppCheck = Test-Path $AppPath

if($RootCheck -eq $true){
    if($AppCheck -eq $true){

        ("") | Out-File $AppPath$LogFile -Append
        ("-~-~--~-~-~-~-~-~-~-~") | Out-File $AppPath$LogFile -Append
        ("Launched AD-Tool at $Date") | Out-File $AppPath$LogFile -Append
        ("-~-~--~-~-~-~-~-~-~-~") | Out-File $AppPath$LogFile -Append
    }
    else{

        mkdir $AppPath
        ("") | Out-File $AppPath$LogFile -Append
        ("-~-~--~-~-~-~-~-~-~-~") | Out-File $AppPath$LogFile -Append
        ("Launched AD-Tool at $Date") | Out-File $AppPath$LogFile -Append
        ("-~-~--~-~-~-~-~-~-~-~") | Out-File $AppPath$LogFile -Append
    }
}
else{

    mkdir $RootPath
    mkdir $AppPath
    ("") | Out-File $AppPath$LogFile -Append
    ("-~-~--~-~-~-~-~-~-~-~") | Out-File $AppPath$LogFile -Append
    ("Launched AD-Tool at $Date") | Out-File $AppPath$LogFile -Append
    ("-~-~--~-~-~-~-~-~-~-~") | Out-File $AppPath$LogFile -Append
}

Try{

    Get-ADUser -Filter "Name -like '*$CurrentUser*'" | Select-Object name, samaccountname
    $Timestamp = [DateTime]::Now.ToString($TimestampFormat)
    ("$Timestamp - Connected to Active Directory as < $CurrentUser >")| Out-File $AppPath$LogFile -Append
    ("-~-") | Out-File $AppPath$LogFile -Append

}Catch{

    $BlankLine = ("------")
    $ConnectionFailed = ("Unable to connect to Active Directory.")
    $ConnectionFailed2 = ("Please ensure you have AD Services installed including the Powershell Module.")

    $lstVerboseOutput.Items.Add($BlankLine)
    $lstVerboseOutput.Items.Add($ConnectionFailed)
    $lstVerboseOutput.Items.Add($BlankLine)
    $lstVerboseOutput.Items.Add($ConnectionFailed2)
    $lstVerboseOutput.Items.Add($BlankLine)

    $Timestamp = [DateTime]::Now.ToString($TimestampFormat)
    ("$Timestamp - $ConnectionFailed")| Out-File $AppPath$LogFile -Append
    ("$Timestamp - $ConnectionFailed2")| Out-File $AppPath$LogFile -Append
    ("-~-") | Out-File $AppPath$LogFile -Append
}

#endregion



# # # # # Code Start # # # # 

#region Code for buttons, lists, text, etc

$btnShowFirstGroups.Add_Click({

    $Timestamp = [DateTime]::Now.ToString($TimestampFormat)
    ("") | Out-File $AppPath$LogFile -Append
    ("") | Out-File $AppPath$LogFile -Append
    ("$Timestamp - FUNCTION - Search First Group.")| Out-File $AppPath$LogFile -Append
    ("/") | Out-File $AppPath$LogFile -Append

    $lstShowFirstGroups.SelectionMode = "One"
    if($txtGetFirstGroupInfo.Text -eq ""){

        ClearVerboseOutput
        $VerboseMessage = ("You can't leave the source box blank.")
        $lstVerboseOutput.Items.Add($VerboseMessage)
    }
    else{
        Try{
            foreach($Item in $lstShowFirstGroups){
            $lstShowFirstGroups.Items.Clear()
            }
        }
        Catch{
            continue
        }
        Try{
            $GroupnameSearch = $txtGetFirstGroupInfo.Text
            $GroupList = Get-ADGroup -Filter "name -like '*$GroupnameSearch*'" | select SamAccountName

            ForEach($Group in $GroupList){
        
                $Groupname = $Group.SamAccountName
                $lstShowFirstGroups.Items.Add($GroupName)
                
                $Timestamp = [DateTime]::Now.ToString($TimestampFormat)
                ("$Timestamp - Group search for < $GroupnameSearch > returned - $Groupname.")| Out-File $AppPath$LogFile -Append
            }
        }Catch{

            ClearVerboseOutput
            $VerboseMessage = ("Couldn't import module for Active Directory.")
            $lstVerboseOutput.Items.Add($VerboseMessage)

            $Timestamp = [DateTime]::Now.ToString($TimestampFormat)
            ("$Timestamp - $VerboseMessage.")| Out-File $AppPath$LogFile -Append
        }
        ("-~-") | Out-File $AppPath$LogFile -Append    
    }
})


$btnShowSecondGroups.Add_Click({

    $Timestamp = [DateTime]::Now.ToString($TimestampFormat)
    ("") | Out-File $AppPath$LogFile -Append
    ("") | Out-File $AppPath$LogFile -Append
    ("$Timestamp - FUNCTION - Search Second Group.")| Out-File $AppPath$LogFile -Append
    ("/") | Out-File $AppPath$LogFile -Append

    if($txtGetSecondGroupInfo.Text -eq ""){

        ClearVerboseOutput
        $VerboseMessage = ("You can't leave the detination box blank.")
        $lstVerboseOutput.Items.Add($VerboseMessage)
    }
    else{
        Try{
            foreach($Item in $lstShowSecondGroups){

            $lstShowSecondGroups.Items.Clear()
            }
        }
        Catch{
            continue
        }
        Try{

            $GroupnameSearch2 = $txtGetSecondGroupInfo.Text
            $GroupList2 = Get-ADGroup -Filter "name -like '*$GroupnameSearch2*'" | select SamAccountName

            ForEach($Group2 in $GroupList2){
        
                $Groupname2 = $Group2.SamAccountName
                $lstShowSecondGroups.Items.Add($GroupName2)
                $Timestamp = [DateTime]::Now.ToString($TimestampFormat)
                ("$Timestamp - Group search for < $GroupnameSearch2 > returned - $Groupname2.")| Out-File $AppPath$LogFile -Append
            }
        }Catch{

            ClearVerboseOutput
            $VerboseMessage = ("Couldn't import module for Active Directory.")
            $lstVerboseOutput.Items.Add($VerboseMessage)

            $Timestamp = [DateTime]::Now.ToString($TimestampFormat)
            ("$Timestamp - $VerboseMessage.")| Out-File $AppPath$LogFile -Append
        }
        ("-~-") | Out-File $AppPath$LogFile -Append
    }
})


$btnSearchByUsername.Add_Click({

    $Timestamp = [DateTime]::Now.ToString($TimestampFormat)
    ("") | Out-File $AppPath$LogFile -Append
    ("") | Out-File $AppPath$LogFile -Append
    ("$Timestamp - FUNCTION - Search by Username.")| Out-File $AppPath$LogFile -Append
    ("/") | Out-File $AppPath$LogFile -Append

    $lstShowFirstGroups.SelectionMode = "MultiExtended"
    if($txtGetUsernameInfo.Text -eq ""){

        ClearVerboseOutput
        $VerboseMessage = ("You can't leave the username box blank.")
        $lstVerboseOutput.Items.Add($VerboseMessage)
    }
    else{
        Try{
            foreach($Item in $lstShowFirstGroups){

            $lstShowFirstGroups.Items.Clear()
            }
        }
        Catch{
            continue
        }
        Try{
            $UsernameSearch = $txtGetUsernameInfo.Text
            $UsernameSearch = $UsernameSearch -replace ".{1}$"
            $UserProfiles = Get-ADUser -Filter "Name -like '*$UsernameSearch*'" | Select-Object name, samaccountname
            
            foreach($User in $UserProfiles){

                $Username = $User.name
                if($Username -notmatch "tmp"){
                    if($Username -notmatch "tmpl"){

                        $lstShowFirstGroups.Items.Add($Username)
                        $Timestamp = [DateTime]::Now.ToString($TimestampFormat)
                        ("$Timestamp - User  search for < $UsernameSearch > returned - $Username.")| Out-File $AppPath$LogFile -Append
                    }
                }
            }
        }Catch{

            ClearVerboseOutput
            $VerboseMessage = ("Couldn't import module for Active Directory.")
            $lstVerboseOutput.Items.Add($VerboseMessage)

            $Timestamp = [DateTime]::Now.ToString($TimestampFormat)
            ("$Timestamp - $VerboseMessage.")| Out-File $AppPath$LogFile -Append
        }
        ("-~-") | Out-File $AppPath$LogFile -Append
    }
})


$btnClearListBoxes.Add_Click({
    
    foreach($Item in $lstShowFirstGroups){

        $lstShowFirstGroups.Items.Clear()
        $lstShowSecondGroups.Items.Clear()
    }

    $txtGetFirstGroupInfo.Text = ""
    $txtGetSecondGroupInfo.Text = ""
    $txtGetUsernameInfo.Text = ""
})


$btnGetLockoutBlame.Add_Click({

    $Timestamp = [DateTime]::Now.ToString($TimestampFormat)
    ("") | Out-File $AppPath$LogFile -Append
    ("") | Out-File $AppPath$LogFile -Append
    ("$Timestamp - FUNCTION - Search for Lockout.")| Out-File $AppPath$LogFile -Append
    ("/") | Out-File $AppPath$LogFile -Append

    ClearVerboseOutput
    $VerboseMessage = ("This requires using an admin account.")
    $lstVerboseOutput.Items.Add($VerboseMessage)
    
    Try{
        $DomainController = Get-ADDomainController -Discover -Service PrimaryDC
        $ComputerName = $DomainController.HostName

        $DC = $ComputerName.ToString()

        Try{
            if($lstShowFirstGroups.SelectedItem -eq $null){
                ClearVerboseOutput
                $VerboseMessage = ("You need to highlight a user to find a lockout. $ComputerName ")
                $lstVerboseOutput.Items.Add($VerboseMessage)    
            }
            else{
                $SelectedUser = $lstShowFirstGroups.SelectedItem
                $UserProfile = Get-ADUser -Filter "Name -eq '$SelectedUser'" | Select-Object name, samaccountname
                $UserProfileSAM = $UserProfile.samaccountname
            
                Try{
                    $Filter = "*[System[EventID=4740] and EventData[Data[@Name='TargetUserName']='$UserProfileSAM']]"
                    $Events = Get-WinEvent -ComputerName $ComputerName -Logname Security -FilterXPath $Filter
                    $LockoutOutput = $Events | Select-Object TimeCreated,
                                                            @{Name='User Name';Expression={$_.Properties[0].Value}},
                                                            @{Name='Source Host';Expression={$_.Properties[1].Value}}
                    ClearVerboseOutput
                    $lstVerboseOutput.Items.Add($LockoutOutput)
                    $Success = ("Task completed.")
                    $lstVerboseOutput.Items.Add($Success)

                    $Timestamp = [DateTime]::Now.ToString($TimestampFormat)
                    ("$Timestamp - $LockoutOutput.")| Out-File $AppPath$LogFile -Append
                    ("$Timestamp - $Success.")| Out-File $AppPath$LogFile -Append


                }Catch{
                    ClearVerboseOutput
                    $VerboseMessage = ("Failed to obtain lockout. Please ensure you are running as admin to complete this search.")
                    $lstVerboseOutput.Items.Add($VerboseMessage)
                    $lstVerboseOutput.Items.Add("")
                    $lstVerboseOutput.Items.Add("Error can be found in log file - C:\Temp\AD-Tool\$LogFile")

                    $Timestamp = [DateTime]::Now.ToString($TimestampFormat)
                    ("$Timestamp - $VerboseMessage.")| Out-File $AppPath$LogFile -Append
                    $_ | Out-File $AppPath$LogFile -Append
                    ("") | Out-File $AppPath$LogFile -Append
                }
            }
        }Catch{
            ClearVerboseOutput
            $VerboseMessage = ("Could not find user: < $SelectedUser >, please ensure you have highlighted a user and not a group.")
            $lstVerboseOutput.Items.Add($VerboseMessage)

            $Timestamp = [DateTime]::Now.ToString($TimestampFormat)
            ("$Timestamp - $VerboseMessage.")| Out-File $AppPath$LogFile -Append
        }
    }Catch{
        ClearVerboseOutput
        $VerboseMessage = ("Couldn't connect to a domain controller.")
        $lstVerboseOutput.Items.Add($VerboseMessage)

        $Timestamp = [DateTime]::Now.ToString($TimestampFormat)
        ("$Timestamp - $VerboseMessage.")| Out-File $AppPath$LogFile -Append
    }
    ("-~-") | Out-File $AppPath$LogFile -Append
})


$btnUnlockAccount.Add_Click({

    $Timestamp = [DateTime]::Now.ToString($TimestampFormat)
    ("") | Out-File $AppPath$LogFile -Append
    ("") | Out-File $AppPath$LogFile -Append
    ("$Timestamp - FUNCTION - Unlock a user account.")| Out-File $AppPath$LogFile -Append
    ("/") | Out-File $AppPath$LogFile -Append

    ClearVerboseOutput
    $VerboseMessage = ("This requires using an admin account.")
    $lstVerboseOutput.Items.Add($VerboseMessage)
    

    Try{
        if($lstShowFirstGroups.SelectedItem -eq $null){
            ClearVerboseOutput
            $VerboseMessage = ("You need to highlight a user to unlock their account.")
            $lstVerboseOutput.Items.Add($VerboseMessage)    
        }
        else{

            $SelectedUser = $lstShowFirstGroups.SelectedItem
            $UserProfile = Get-ADUser -Filter "Name -eq '$SelectedUser'" -Properties  name, samaccountname, LockedOut
            $UserProfileSAM = $UserProfile.samaccountname
            $UserProfileLocked = $UserProfile.LockedOut
            
            Try{

                If($UserProfileLocked -eq $false){
                
                    ClearVerboseOutput
                    $VerboseMessage = ("Cannot unlock the selected account because is not currently locked out.")
                    $lstVerboseOutput.Items.Add($VerboseMessage)

                    $Timestamp = [DateTime]::Now.ToString($TimestampFormat)
                    ("$Timestamp - The account for < $SelectedUser > was already unlocked.")| Out-File $AppPath$LogFile -Append

                }
                elseif($UserProfileLocked -eq $true){

                    Unlock-ADAccount -Identity $UserProfileSAM

                    $UserProfile = Get-ADUser -Filter "Name -eq '$SelectedUser'" -Properties  name, samaccountname, LockedOut
                    $UserProfileSAM = $UserProfile.samaccountname
                    $UserProfileLocked = $UserProfile.LockedOut

                    If($UserProfileLocked -eq $false){

                        ClearVerboseOutput
                        $VerboseMessage = ("Account has been unlocked.")
                        $lstVerboseOutput.Items.Add($VerboseMessage)

                        $Timestamp = [DateTime]::Now.ToString($TimestampFormat)
                        ("$Timestamp - The account for < $SelectedUser > has been unlocked.")| Out-File $AppPath$LogFile -Append
                    }
                    elseif($UserProfileLocked -eq $true){

                        ClearVerboseOutput
                        $VerboseMessage = ("Checking account . .")
                        $lstVerboseOutput.Items.Add($VerboseMessage)
                        $VerboseMessage = ("")
                        $lstVerboseOutput.Items.Add($VerboseMessage)
                        $VerboseMessage = ("Account is locked, either unlocking failed or account has been locked again instantly.")
                        $lstVerboseOutput.Items.Add($VerboseMessage)

                        $Timestamp = [DateTime]::Now.ToString($TimestampFormat)
                        ("$Timestamp - The account for < $SelectedUser > failed to be unlocked.")| Out-File $AppPath$LogFile -Append
                    }
                }
            }Catch{
                ClearVerboseOutput
                $VerboseMessage = ("Failed to unlock. Please ensure you are running as admin to perform this function.")
                $lstVerboseOutput.Items.Add($VerboseMessage)
                $lstVerboseOutput.Items.Add("")
                $lstVerboseOutput.Items.Add("Error can be found in log file - C:\Temp\AD-Tool\$LogFile")

                $Timestamp = [DateTime]::Now.ToString($TimestampFormat)
                ("$Timestamp - $VerboseMessage.")| Out-File $AppPath$LogFile -Append
                $_ | Out-File $AppPath$LogFile -Append
                ("") | Out-File $AppPath$LogFile -Append
            }
        }
    }Catch{
        ClearVerboseOutput
        $VerboseMessage = ("Could not find user: < $SelectedUser >, please ensure you have highlighted a user and not a group.")
        $lstVerboseOutput.Items.Add($VerboseMessage)

        $Timestamp = [DateTime]::Now.ToString($TimestampFormat)
        ("$Timestamp - $VerboseMessage.")| Out-File $AppPath$LogFile -Append
    }

    ("-~-") | Out-File $AppPath$LogFile -Append
})


$btnShowFirstUsers.Add_click({

    $Timestamp = [DateTime]::Now.ToString($TimestampFormat)
    ("") | Out-File $AppPath$LogFile -Append
    ("") | Out-File $AppPath$LogFile -Append
    ("$Timestamp - FUNCTION - Show Users in Group.")| Out-File $AppPath$LogFile -Append
    ("/") | Out-File $AppPath$LogFile -Append

    $lstShowFirstGroups.SelectionMode = "MultiExtended"
    if($lstShowFirstGroups.SelectedItem -eq $null){

        ClearVerboseOutput
        $VerboseMessage = ("You need to search for a group, or highlight a group from the first box.")
        $lstVerboseOutput.Items.Add($VerboseMessage)
    }
    else{
        Try{
            $SelectedGroup = $lstShowFirstGroups.SelectedItem

            foreach($Item in $lstShowFirstGroups){

                $lstShowFirstGroups.Items.Clear()
            }
        }Catch{
            continue
        }   
        Try{
            $Counter = 0
            $UserProfiles = Get-ADGroupMember -Identity $SelectedGroup -Recursive | Get-ADUser -Properties displayname, objectclass, name
            ForEach ($UserProfile in $UserProfiles){
 
                $Username = $UserProfile.name
                $ObjectClass = $UserProfile.objectclass
                $DisplayName = $UserProfile.displayname
                
                foreach($User in $UserProfile){

                    if($Username -notmatch "tmpl"){
                        if($Username -notmatch "tmp"){

                            $lstShowFirstGroups.Items.Add($Username)
                            $counter += 1
                        
                            $Timestamp = [DateTime]::Now.ToString($TimestampFormat)
                            ("$Timestamp - Group < $SelectedGroup > contains user - $Username.")| Out-File $AppPath$LogFile -Append
                        }
                    }
                }
        
                ClearVerboseOutput
                $VerboseMessage = ("$Counter users found.")
                $lstVerboseOutput.Items.Add($VerboseMessage)

            }
        }Catch{

            ClearVerboseOutput
            $VerboseMessage = ("Ensure you havn't highlighted a user already.")
            $lstVerboseOutput.Items.Add($VerboseMessage)
        }
    }
    ("-~-") | Out-File $AppPath$LogFile -Append       
})


$btnCopyUsersOver.Add_Click({

    $Timestamp = [DateTime]::Now.ToString($TimestampFormat)
    ("") | Out-File $AppPath$LogFile -Append
    ("") | Out-File $AppPath$LogFile -Append
    ("$Timestamp - FUNCTION - Copy All Users in First Group to Second Group.")| Out-File $AppPath$LogFile -Append
    ("/") | Out-File $AppPath$LogFile -Append

    Try{   
        $FirstSelectedGroup = $lstShowFirstGroups.SelectedItem
        $SecondSelectedGroup = $lstShowSecondGroups.SelectedItem

        if($FirstSelectedGroup -eq $null -and $SecondSelectedGroup -eq $null){

            ClearVerboseOutput
            $VerboseMessage = ("You need to search and select a group from each box.")
            $lstVerboseOutput.Items.Add($VerboseMessage)
        }
        elseif($FirstSelectedGroup -eq $null){
            
            ClearVerboseOutput
            $VerboseMessage = ("You haven't selected the first group.")
            $lstVerboseOutput.Items.Add($VerboseMessage)
        }
        elseif($SecondSelectedGroup -eq $null){
            
            ClearVerboseOutput
            $VerboseMessage = ("You haven't selected the second group")
            $lstVerboseOutput.Items.Add($VerboseMessage)
        }
        else{
            Try{

                ClearVerboseOutput
                $Group = Get-ADGroupMember -Identity $FirstSelectedGroup -Recursive | Get-ADUser -Properties displayname, objectclass, name, samaccountname
                $Count = 0

                ForEach ($GroupMember in $Group){

                    $Username = $GroupMember.samaccountname
                    $LifeName = $GroupMember.name

                    Try{
                        if($Lifename -notmatch "tmp"){
                            if($Lifename -notmatch "tmpl"){

                                $VerboseMessage = ("Would be copying $Lifename from $FirstSelectedGroup to $SecondSelectedGroup")
                                $lstVerboseOutput.Items.Add($VerboseMessage)
                                $Count += 1
                            }
                        }
                    }Catch{
                                    
                        $VerboseMessage = ("Could not copy $GroupMember to $SecondSelectedGroup")
                        $lstVerboseOutput.Items.Add($VerboseMessage)
                    }
                }
                # Confirm choices before copying any users.
                if([System.Windows.Forms.MessageBox]::Show("Copy $Count users?", "Confirm",[System.Windows.Forms.MessageBoxButtons]::OKCancel) -eq "OK"){

                    ClearVerboseOutput
                    $Count = 0
                                
                    foreach($GroupMember in $Group){

                        $Username = $GroupMember.samaccountname
                        $LifeName = $GroupMember.name

                        Try{
                            if($Lifename -notmatch "tmp"){
                                if($Lifename -notmatch "tmpl"){

                                    $VerboseMessage = ("Copying $Lifename to $SecondSelectedGroup")
                                    $lstVerboseOutput.Items.Add($VerboseMessage)

                                    Add-ADGroupmember -Identity $SecondSelectedGroup -Members $Username
                                    $Count += 1

                                    $Timestamp = [DateTime]::Now.ToString($TimestampFormat)
                                    ("$Timestamp - Copied < $LifeName > from < $FirstSelectedGroup > to - $SecondSelectedGroup.")| Out-File $AppPath$LogFile -Append
                                }
                            }
                        }Catch{   
                                     
                            $VerboseMessage = ("Could not copy $SelectedUser to $SecondSelectedGroup")
                            $lstVerboseOutput.Items.Add($VerboseMessage)

                            $Timestamp = [DateTime]::Now.ToString($TimestampFormat)
                            ("$Timestamp - FAILED to copy < $LifeName > from < $FirstSelectedGroup > to - $SecondSelectedGroup. Error:")| Out-File $AppPath$LogFile -Append
                            $_ | Out-File $AppPath$LogFile -Append
                            ("") | Out-File $AppPath$LogFile -Append
                        }
                        $Timestamp = [DateTime]::Now.ToString($TimestampFormat)
                        ("$Timestamp - Copied $Count users.")| Out-File $AppPath$LogFile -Append
                    }
                }
                else{

                    ClearVerboseOutput
                    $VerboseMessage = ("Cancelled. No users copied.")
                    $lstVerboseOutput.Items.Add($VerboseMessage)
                }
            }Catch{

                ClearVerboseOutput
                $VerboseMessage = ("Could not transfer users between groups.")
                $lstVerboseOutput.Items.Add($VerboseMessage)

                $Timestamp = [DateTime]::Now.ToString($TimestampFormat)
                ("$Timestamp - FAILED to copy users between groups.")| Out-File $AppPath$LogFile -Append
            }
        }
    }Catch{

        ClearVerboseOutput
        $VerboseMessage = ("Ensure you have selected a group in each box.")
        $lstVerboseOutput.Items.Add($VerboseMessage)
    }
    ("-~-") | Out-File $AppPath$LogFile -Append
})


$btnShowMemberships.Add_Click({

    $Timestamp = [DateTime]::Now.ToString($TimestampFormat)
    ("") | Out-File $AppPath$LogFile -Append
    ("") | Out-File $AppPath$LogFile -Append
    ("$Timestamp - FUNCTION - Show Which Groups the User Belongs to.")| Out-File $AppPath$LogFile -Append
    ("/") | Out-File $AppPath$LogFile -Append
    
    $SelectedUser = $lstShowFirstGroups.SelectedItem

    if($SelectedUser -eq $null){

        ClearVerboseOutput
        $VerboseMessage = ("You need to highlight a user to view their memberships.")
        $lstVerboseOutput.Items.Add($VerboseMessage)
    }
    else{
        Try{

            foreach($Item in $lstShowSecondGroups){

                $lstShowSecondGroups.Items.Clear()
            }
        }Catch{
            continue
        }
        Try{

            $UserProfile = Get-ADUser -Filter "Name -eq '$SelectedUser'" | Select-Object name, samaccountname
            $UserProfileSAM = $UserProfile.samaccountname
            $Memberships = Get-ADPrincipalGroupMembership -Identity $UserProfileSAM | select -ExpandProperty Name | Sort

            Foreach ($Group in $Memberships){

                $lstShowSecondGroups.Items.Add($Group)
                $Timestamp = [DateTime]::Now.ToString($TimestampFormat)
                ("$Timestamp - < $UserProfileSAM > belongs to - $Group.")| Out-File $AppPath$LogFile -Append

            }
        }Catch{

            ClearVerboseOutput
            $VerboseMessage = ("Could not display memeberships. Ensure you have selected a user and not a group.")
            $lstVerboseOutput.Items.Add($VerboseMessage)

            $Timestamp = [DateTime]::Now.ToString($TimestampFormat)
            ("$Timestamp - Failed to diplayed memberships, please ensure this a username is selected.")| Out-File $AppPath$LogFile -Append
        }
    }
    ("-~-") | Out-File $AppPath$LogFile -Append
})


$btnUserInfo.Add_Click({

    $Timestamp = [DateTime]::Now.ToString($TimestampFormat)
    ("") | Out-File $AppPath$LogFile -Append
    ("") | Out-File $AppPath$LogFile -Append
    ("$Timestamp - FUNCTION - Get details about selected user.")| Out-File $AppPath$LogFile -Append
    ("/") | Out-File $AppPath$LogFile -Append
    
    $SelectedUser = $lstShowFirstGroups.SelectedItem

    if($SelectedUser -eq $null){

        ClearVerboseOutput
        $VerboseMessage = ("You need to highlight a user to view their details.")
        $lstVerboseOutput.Items.Add($VerboseMessage)
    }
    else{
        Try{
            ClearVerboseOutput
            $Timestamp = [DateTime]::Now.ToString($TimestampFormat)
            ("$Timestamp - Results for - $SelectedUser")| Out-File $AppPath$LogFile -Append

            $ComputerProperties = Get-ADComputer -Filter "Description -like '*$SelectedUser*'" -Properties *
            $CompProperties = ("DNSHostName", "Description", "IPv4Address")

            $AllProperties = Get-ADUser -Filter "Name -eq '$SelectedUser'" -Properties *
            $Properties = ("DisplayName", "SamAccountName", "Title", "Department", "Description", "LockedOut", "logonCount", "LastLogonDate", "OfficePhone", "whenCreated", "whenChanged")

            Foreach ($Property in $CompProperties){

                Try{

                    $PropertyValue = $ComputerProperties.$Property

                    $PropertyObj = New-Object psobject -Property @{$Property = $PropertyValue}

                    $VerboseInfo = $PropertyObj | Format-List | Out-String

                    $lstVerboseOutput.Items.Add("$VerboseInfo")
                    $lstVerboseOutput.Items.Add("")
                    $Timestamp = [DateTime]::Now.ToString($TimestampFormat)
                    ("$Timestamp - $Property - $PropertyValue")| Out-File $AppPath$LogFile -Append
                
                }Catch{

                    $VerboseInfo = ("Couldn't obtain $Property for $SelectedUser")

                    $lstVerboseOutput.Items.Add("$VerboseInfo")
                    $lstVerboseOutput.Items.Add("")
                    $Timestamp = [DateTime]::Now.ToString($TimestampFormat)
                    ("$Timestamp - Couldn't obtain < $Property > for - $SelectedUser")| Out-File $AppPath$LogFile -Append
                }
            }
            Foreach ($Property in $Properties){

                $PropertyValue = $AllProperties.$Property


                $PropertyObj = New-Object psobject -Property @{$Property = $PropertyValue}

                $VerboseInfo = $PropertyObj | Format-List | Out-String

                $lstVerboseOutput.Items.Add("$VerboseInfo")
                $lstVerboseOutput.Items.Add("")
                $Timestamp = [DateTime]::Now.ToString($TimestampFormat)
                ("$Timestamp - $Property - $PropertyValue")| Out-File $AppPath$LogFile -Append
            }

            
        }Catch{

            ClearVerboseOutput
            $VerboseMessage = ("Could not display details. Ensure you have selected a user and not a group.")
            $lstVerboseOutput.Items.Add($VerboseMessage)

            $Timestamp = [DateTime]::Now.ToString($TimestampFormat)
            ("$Timestamp - Failed to diplayed details, please ensure this a username is selected.")| Out-File $AppPath$LogFile -Append
        }
    }
    ("-~-") | Out-File $AppPath$LogFile -Append
})


$btnShowHelp.Add_Click({

    ClearVerboseOutput

    $BlankLine = ("------")
    $InfoNumber0 = ("Permissions within this app are based on your user account.")
    $InfoNumber1 = ("Search for the source and destination groups using the search boxes.")
    $InfoNumber2 = ("After highlighting a group from each box, select 'Copy All Users' to copy all the users from the left to the right group.")
    $InfoNumber3 = ("You can also search for users by username or name which will be displayed in the first box.")
    $InfoNumber4 = ("Highlight users from the first box and choose 'Copy Selected Users' to copy only these users to the destination group.")
    $InfoNumber5 = ("Highlight a user and select 'Show Users Groups' to fill the second box with all the groups that the user belongs to.")
    $InfoNumber6 = ("Highlight a user and select 'Lockout Location' to find the machine that locked out their account.")
    $InfoNumber7 = ("Highlight a group in the source box and select 'Show Group Members' to display the users that are part of that group.")
    $InfoNumber8 = ("Highlight a user and select 'Get User Info' to see their machine name, AD info, IP Address etc")
    $lstVerboseOutput.Items.Add($BlankLine)
    $lstVerboseOutput.Items.Add($InfoNumber0)
    $lstVerboseOutput.Items.Add($BlankLine)
    $lstVerboseOutput.Items.Add($InfoNumber1)
    $lstVerboseOutput.Items.Add($BlankLine)
    $lstVerboseOutput.Items.Add($InfoNumber2)
    $lstVerboseOutput.Items.Add($BlankLine)
    $lstVerboseOutput.Items.Add($InfoNumber3)
    $lstVerboseOutput.Items.Add($BlankLine)
    $lstVerboseOutput.Items.Add($InfoNumber4)
    $lstVerboseOutput.Items.Add($BlankLine)
    $lstVerboseOutput.Items.Add($InfoNumber5)
    $lstVerboseOutput.Items.Add($BlankLine)
    $lstVerboseOutput.Items.Add($InfoNumber6)
    $lstVerboseOutput.Items.Add($BlankLine)
    $lstVerboseOutput.Items.Add($InfoNumber7)
    $lstVerboseOutput.Items.Add($BlankLine)
    $lstVerboseOutput.Items.Add($InfoNumber8)
})


$btnCopySelectedUsers.Add_Click({

    $Timestamp = [DateTime]::Now.ToString($TimestampFormat)
    ("") | Out-File $AppPath$LogFile -Append
    ("") | Out-File $AppPath$LogFile -Append
    ("$Timestamp - FUNCTION - Copy Selected Users to Second Group.") | Out-File $AppPath$LogFile -Append
    ("/") | Out-File $AppPath$LogFile -Append

    ClearVerboseOutput
    
    $SelectedUsers = $lstShowFirstGroups.SelectedItems
    $SecondSelectedGroup = $lstShowSecondGroups.SelectedItem

    $Count = 0

    if($SelectedUsers -eq $null){

        ClearVerboseOutput
        $VerboseMessage = ("You need to highlight at least one user to copy them to a group.")
        $lstVerboseOutput.Items.Add($VerboseMessage)
    }
    elseif($SecondSelectedGroup -eq $null){
            
        ClearVerboseOutput
        $VerboseMessage = ("You haven't selected a group to copy the users to.")
        $lstVerboseOutput.Items.Add($VerboseMessage)
    }
    else{
        Try{

            $SelectedUser = $SelectedUsers[0]
            $UserProfile = Get-ADUser -Filter "Name -eq '$SelectedUser'" | Select-Object name, samaccountname
            $username = $UserProfile.samaccountname

            if($Username -eq "" -or $username -eq $null){

                ClearVerboseOutput
                $VerboseMessage = ("Ensure you have selected a user and not a group.")
                $lstVerboseOutput.Items.Add($VerboseMessage)
            }
            else{
                foreach($SelectedUser in $SelectedUsers){

                    $UserProfile = Get-ADUser -Filter "Name -eq '$SelectedUser'" | Select-Object name, samaccountname
                    $Username = $UserProfile.samaccountname
                    $LifeName = $UserProfile.name

                    Try{

                        if($Lifename -notmatch "tmp"){
                            if($Lifename -notmatch "tmpl"){

                                $VerboseMessage = ("Would be copying $Lifename to $SecondSelectedGroup")
                                $lstVerboseOutput.Items.Add($VerboseMessage)
                                $Count += 1
                            }
                        }
                    }Catch{                
                        $VerboseMessage = ("Could not copy $SelectedUser to $SecondSelectedGroup")
                        $lstVerboseOutput.Items.Add($VerboseMessage)
                    }
                }
                #Ask the user to confirm before copying any users
                if([System.Windows.Forms.MessageBox]::Show("Copy $Count users?", "Confirm",[System.Windows.Forms.MessageBoxButtons]::OKCancel) -eq "OK"){

                    ClearVerboseOutput
                    $Count = 0
                                
                    foreach($SelectedUser in $SelectedUsers){

                        $UserProfile = Get-ADUser -Filter "Name -eq '$SelectedUser'" | Select-Object name, samaccountname
                        $Username = $UserProfile.samaccountname
                        $LifeName = $UserProfile.name

                        Try{
                            if($Lifename -notmatch "tmp"){
                                if($Lifename -notmatch "tmpl"){

                                    $VerboseMessage = ("Copying $Lifename to $SecondSelectedGroup")
                                    $lstVerboseOutput.Items.Add($VerboseMessage)
                                    Add-ADGroupmember -Identity $SecondSelectedGroup -Members $Username
                                    $Count += 1

                                    $Timestamp = [DateTime]::Now.ToString($TimestampFormat)
                                    ("$Timestamp - Copied < $LifeName > to - $SecondSelectedGroup.")| Out-File $AppPath$LogFile -Append
                                }
                            }
                        }Catch{
                            
                            $VerboseMessage = ("Could not copy $SelectedUser to $SecondSelectedGroup")
                            $lstVerboseOutput.Items.Add($VerboseMessage)

                            $Timestamp = [DateTime]::Now.ToString($TimestampFormat)
                            ("$Timestamp - FAILED to copy < $LifeName > to - $SecondSelectedGroup. Error:")| Out-File $AppPath$LogFile -Append
                            $_ | Out-File $AppPath$LogFile -Append
                            ("") | Out-File $AppPath$LogFile -Append
                        }
                    }
                }
                else{

                    ClearVerboseOutput
                    $VerboseMessage = ("Cancelled.")
                    $lstVerboseOutput.Items.Add($VerboseMessage)
                }
            }
        }Catch{

            ClearVerboseOutput
            $VerboseMessage = ("Ensure you have selected a user and not a group.")
            $lstVerboseOutput.Items.Add($VerboseMessage)
        }
    }
    ("-~-") | Out-File $AppPath$LogFile -Append   
})


Function ClearVerboseOutput(){

    Try{
        foreach($Item in $lstShowFirstGroups){

            $lstVerboseOutput.Items.Clear()
        }
    }Catch{
        continue
    }
}


#endregion

# # # # # Code End # # # # #
#

#Launch form
$frmInitialScreen.ShowDialog()