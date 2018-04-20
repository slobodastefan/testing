#
# Created by Bernard Welmers
# On March 11, 2016
#
# This script was created to automate the user account creation process at Knutson Construction.
#
# Pre-requisits
# This script needs Excel installed on the computer that it is running on
# IT also needs the RSAT tools installed.
#



# Stop script on errors (such as user already exists) so that it does not try and add groups or mailbox to existing user
$ErrorActionPreference = “Stop”


#This creates a list box that allows people to select items from the list.
Function ListBox ([array]$ListofGroups)
{
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $form = New-Object System.Windows.Forms.Form 
    $form.Text = "Group Selection form"
    $form.Size = New-Object System.Drawing.Size(600,800) 
    $form.StartPosition = "CenterScreen"

    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Location = New-Object System.Drawing.Point(75,720)
    $OKButton.Size = New-Object System.Drawing.Size(75,23)
    $OKButton.Text = "OK"
    $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.AcceptButton = $OKButton
    $form.Controls.Add($OKButton)

    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Location = New-Object System.Drawing.Point(150,720)
    $CancelButton.Size = New-Object System.Drawing.Size(75,23)
    $CancelButton.Text = "Cancel"
    $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.CancelButton = $CancelButton
    $form.Controls.Add($CancelButton)

    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(10,20) 
    $label.Size = New-Object System.Drawing.Size(280,20) 
    $label.Text = "Please make a selection from the list below:"
    $form.Controls.Add($label) 

    $listBox = New-Object System.Windows.Forms.ListBox 
    $listBox.Location = New-Object System.Drawing.Point(10,40) 
    $listBox.Size = '460,670' 
    $listbox.ColumnWidth = 225
    $listBox.MultiColumn = $True
    $listBox.SelectionMode = "MultiExtended"


    foreach ($ListofGroups in $ListofGroups)
        {
        [void] $listBox.Items.Add($ListofGroups.name)
        }


    $form.Controls.Add($listBox) 
    $form.Topmost = $True

    $result = $form.ShowDialog()



    if ($result -eq [System.Windows.Forms.DialogResult]::OK)
    {
        $ListofGroups = $listBox.SelectedItems
    }

    return ,$ListofGroups

}

#Take the Excel document and change it into a CSV File that can then be parsed quicker
function Import-Excel([string]$FilePath, [string]$SheetName = "")
{
    $csvFile = Join-Path $env:temp ("{0}.csv" -f (Get-Item -path $FilePath).BaseName)
    if (Test-Path -path $csvFile) { Remove-Item -path $csvFile }
 
    # convert Excel file to CSV file
    $xlCSVType = 6 # SEE: http://msdn.microsoft.com/en-us/library/bb241279.aspx
    $excelObject = New-Object -ComObject Excel.Application 
    $excelObject.Visible = $false
    $workbookObject = $excelObject.Workbooks.Open($FilePath)
    SetActiveSheet $workbookObject $SheetName | Out-Null
    $workbookObject.SaveAs($csvFile,$xlCSVType)
    $workbookObject.Saved = $true
    $workbookObject.Close()
 
     # cleanup
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbookObject) |
        Out-Null
    $excelObject.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excelObject) |
        Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
 
    # now import and return the data
    Import-Csv -path $csvFile
}

 
function FindSheet([Object]$workbook, [string]$name)
{
    $sheetNumber = 0
    for ($i=1; $i -le $workbook.Sheets.Count; $i++) {
        if ($name -eq $workbook.Sheets.Item($i).Name) { $sheetNumber = $i; break }
    }
    return $sheetNumber
}
 
function SetActiveSheet([Object]$workbook, [string]$name)
{
    if (!$name) { return }
    $sheetNumber = FindSheet $workbook $name
    if ($sheetNumber -gt 0) { $workbook.Worksheets.Item($sheetNumber).Activate() }
    return ($sheetNumber -gt 0)
}
 



#Check to see if a connection MS Exchagne (MEX2) has already been established in this PowerShell Session
if (!(Get-PSSession | Where { $_.ConfigurationName -eq "Microsoft.Exchange" })) 
{ 
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://ExchangeServer/PowerShell -Authentication Kerberos
    
    Import-PSSession $Session -ErrorAction Stop 
}

"Geting user information from Excel File"
$Users = Import-Excel -FilePath "\\ITScripts\UserCreation.xlsx" -SheetName "UserList"
   
"Now run through every user in the sheet and create their account, update their user directory, and create their mailbox"
foreach ($User in $Users)            
{            

    #Process the data in the sheet (get it ready to create the account)
    $Displayname = $User.Firstname + " " + $User.Lastname            
    $UserFirstname = $User.Firstname            
    $UserLastname = $User.Lastname 
    $SAM = $User.SamAccountName    
    #If SAM (Username) is empty then create a derived SAM        
    if (!$Sam) 
    {
        $SAM = $userfirstname.substring(0,1) + $Userlastname
    }
    
    #UPN is derived from the SAM (newer type of user account)
    $UPN = $SAM + "@internal.domain"         
    $Description = $User.Description            
    $Password = $User.Password            
	$Manager= $User.Manager

    #Setup user's home directory and "Z" path
    $TSPath = "\\server\tshome45\"+$SAM
    $ExpireDate = $user.AccountExpirationDate -as [datetime]
    if (!$ExpireDate)
    {  
        $ExpireDate = $null
    }

    #Add basic office address information
	if ($user.Office -eq "Rochester")
		{
		$StreetAddress = "XXX"
		$City = "Rochester"
		$State = "Minnesota"
		$Country = "USA"
		$PostalCode = "55901"
		$Fax = "XXX"
		$pager = "Main office number"
		$OU = "OU=People,DC=com"
        $Qpath = "\\server\users$\"+$sam

		}
	Elseif ($user.Office -eq "Iowa")
		{
		$StreetAddress = "XXX"
		$City = "Iowa City"
		$State = "Iowa"
		$Country = "USA"
		$PostalCode = "52240"
		$Fax = "XXX"
		$pager = "Main office number"
		$OU = "OU=People,DC=com"
        $Qpath = "\\server\users\"+$sam
		}
	Else
		{
		$StreetAddress = "XXX"
		$City = "Minneapolis"
		$State = "Minnesota"
		$Country = "USA"
		$PostalCode = "55426"
		$Fax = "XXX"
		$pager = "Main Office Number"
		$OU = "OU=People,DC=com"
        $Qpath = "\\server\tsprofiles$\"+ $sam
		}


"Now lets actually create the user"
    New-ADUser -Name "$Displayname" `
    -DisplayName "$Displayname" `
    -SamAccountName $SAM `
    -UserPrincipalName $UPN `
    -GivenName "$UserFirstname" `
    -Surname "$UserLastname" `
    -Description "$Description" `
    -AccountPassword (ConvertTo-SecureString $Password -AsPlainText -Force) `
    -Enabled $true  `
    -Path "$OU" `
    -CannotChangePassword ([System.Convert]::ToBoolean($User.CannotChangePassword)) `
    -PasswordNeverExpires ([System.Convert]::ToBoolean($User.PasswordNeverExpires)) `
    -ChangePasswordAtLogon ([System.Convert]::ToBoolean($User.ChangePasswordAtLogon)) `
    -AccountExpirationDate $ExpireDate `
    -Company "Knutson Construction" `
    -StreetAddress $StreetAddress `
    -City $City `
    -State $State `
    -Country $Country `
    -PostalCode $PostalCode `
    -Fax $Fax `
    -Office $user.Office `
    -Title $User.Title `
    -EmployeeID $User.EmployeeID `
    -Department $User.Department `
    -OfficePhone $User.OfficePhone `
    -MobilePhone $User.MobilePhone `
    -Manager $Manager `
    -OtherAttributes @{'Pager'= $Pager; 'msTSHomeDirectory' = $TSPath; 'msTSHomeDrive' = "Z:"} `
    -HomeDirectory $TSPath `
    -HomeDrive "Z:"`
	-HomePage "www.website.com"
#Other Attributes is the way to get "non standard" attributes added to the profile. It is possible to add more items using a ";" as delimiter

    "Create the mailbox for the user"
	Enable-Mailbox -identity $SAM.DistinguishedName -database FirstMailboxDatabase -retentionPolicy 3906f066-1625-4957-af72-3c438dc0f389

       
    "Create the TS Home45 directory"
    if(!(Test-Path -Path $TSPath)) 
        {
        New-Item -ItemType Directory -Force -Path $TSPath
        }
    
    "Give newly created person access to their TS Home 45 directory"
    $acl = get-acl $tspath
    $acl.setowner([System.Security.Principal.NTAccount] $SAM)
    set-acl $tspath $acl

    "Give newly created person access to their Q drive directory"
    if(!(Test-Path -Path $Qpath)) 
        {
            New-Item -ItemType Directory -Force -Path $Qpath
        }
    $acl = get-acl $Qpath
    $acl.setowner([System.Security.Principal.NTAccount] $SAM)
    set-acl $Qpath $acl


    "Create a list of all the Access Control List AD Groups and display them for selection"
    $ADGroupList = ListBox (Get-ADGroup -Filter {name -like "*"} -SearchBase "OU=SecurityGroups,DC=com" | sort)

    "add the user to each group that was selected"
    foreach ($item in $ADGroupList) 
    {
        add-adgroupmember $item $SAM
    }

    $ADGroupList = $Null

    "Create a list of all the Distribution Groups and display them for selection"
    $ADGroupList = ListBox (Get-ADGroup -Filter {name -like "*"} -SearchBase "OU=Distro-Departmental,DC=com" | sort)

    "add the user to each group that was selected"
    foreach ($item in $ADGroupList) 
    {
        add-adgroupmember $item $SAM
    }
}