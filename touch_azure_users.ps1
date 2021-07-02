Add-Type -AssemblyName PresentationFramework

# Test For Modules
if(-not(Get-Module ExchangeOnlineManagement -ListAvailable)){
    $null = [System.Windows.MessageBox]::Show('Please Install ExchangeOnlineManagement - view ITG for details')
    Exit
}
if(-not(Get-Module AzureAD -ListAvailable)){
    $null = [System.Windows.MessageBox]::Show('Please Install ExchangeOnlineManagement - view ITG for details')
    Exit
}

if(-not(Get-Module Microsoft.Online.SharePoint.PowerShell -ListAvailable)){
    $null = [System.Windows.MessageBox]::Show('Please Install ExchangeOnlineManagement - view ITG for details')
    Exit
}

### Start XAML and Reader to use WPF, as well as declare variables for use
[xml]$xaml = @"
<Window

  xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"

  xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"

  Title="Touch Users" Height="526.201" Width="525" ResizeMode="NoResize">

    <Grid ScrollViewer.HorizontalScrollBarVisibility="Auto" ScrollViewer.VerticalScrollBarVisibility="Auto">
        <TabControl Name="Tabs" HorizontalAlignment="Left" Height="487" Margin="10,0,0,0" VerticalAlignment="Top" Width="499">
            <TabItem Name="ResetTab" Header="Reset Password">
                <Grid Background="#FFE5E5E5">
                    <Label Content="Please Pick A User, Then Enter A Password" HorizontalAlignment="Left" Margin="10,10,0,0" VerticalAlignment="Top" Height="25" Width="473"/>
                    <TextBox Name="PasswordTextBox" HorizontalAlignment="Left" Height="25" Margin="10,130,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="473"/>
                    <Button Name="PasswordGoButton" Content="Go" HorizontalAlignment="Left" Margin="10,399,0,0" VerticalAlignment="Top" Width="473" Height="50" IsEnabled="False"/>
                    <TextBox Name="UserTextBox" HorizontalAlignment="Left" Height="23" Margin="258,55,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="225" IsReadOnly="True" Background="#FFC8C8C8"/>
                    <Button Name="UserButton" Content="Pick User" HorizontalAlignment="Left" Margin="10,40,0,0" VerticalAlignment="Top" Width="243" Height="54"/>
                    <Label Content="Enter Password Below, Go Button Activates At 8 Characters and User Selected" HorizontalAlignment="Left" Margin="10,99,0,0" VerticalAlignment="Top" Width="473"/>
                    <RichTextBox Name="PasswordRichTextBox" HorizontalAlignment="Left" Height="234" Margin="10,160,0,0" VerticalAlignment="Top" Width="473" Background="#FF646464" Foreground="Cyan" HorizontalScrollBarVisibility="Auto" VerticalScrollBarVisibility="Auto" IsReadOnly="True">
                        <FlowDocument/>
                    </RichTextBox>
                </Grid>
            </TabItem>
            <TabItem Name="CreateTab" Header="Create User">
                <Grid Background="#FFE5E5E5">
                    <Label Content="First Name" HorizontalAlignment="Left" Margin="10,10,0,0" VerticalAlignment="Top" Width="67"/>
                    <TextBox Name="FirstNameTextbox" HorizontalAlignment="Left" Height="23" Margin="10,41,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="168" TabIndex="0"/>
                    <Label Content="Last Name" HorizontalAlignment="Left" Margin="10,69,0,0" VerticalAlignment="Top" Width="67"/>
                    <TextBox Name="LastNameTextbox" HorizontalAlignment="Left" Height="23" Margin="10,100,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="168" TabIndex="1"/>
                    <Label Content="Username" HorizontalAlignment="Left" Margin="10,128,0,0" VerticalAlignment="Top" Width="67"/>
                    <TextBox Name="UsernameTextbox" HorizontalAlignment="Left" Height="23" Margin="10,159,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="168" TabIndex="2"/>
                    <Label Content="@" HorizontalAlignment="Left" Margin="10,187,0,0" VerticalAlignment="Top" Width="67"/>
                    <TextBox Name="CreatePasswordTextbox" HorizontalAlignment="Left" Height="23" Margin="10,276,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="168" TabIndex="4"/>
                    <Label Content="Password" HorizontalAlignment="Left" Margin="10,245,0,0" VerticalAlignment="Top" Width="67"/>
                    <ComboBox Name="DomainCombobox" HorizontalAlignment="Left" Margin="10,218,0,0" VerticalAlignment="Top" Width="168" TabIndex="3"/>
                    <Label Content="Usage Location" HorizontalAlignment="Left" Margin="10,304,0,0" VerticalAlignment="Top" Width="91"/>
                    <ComboBox Name="UsageLocationCombobox" HorizontalAlignment="Left" Margin="10,335,0,0" VerticalAlignment="Top" Width="168" TabIndex="5"/>
                    <Label Content="City" HorizontalAlignment="Left" Margin="343,10,0,0" VerticalAlignment="Top" Width="57"/>
                    <TextBox Name="CityTextbox" HorizontalAlignment="Left" Height="23" Margin="343,41,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="140" TabIndex="6"/>
                    <Label Content="State" HorizontalAlignment="Left" Margin="343,69,0,0" VerticalAlignment="Top" Width="57"/>
                    <TextBox Name="StateTextbox" HorizontalAlignment="Left" Height="23" Margin="343,100,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="140" TabIndex="6"/>
                    <Label Content="CustomAttribute1" HorizontalAlignment="Left" Margin="343,128,0,0" VerticalAlignment="Top" Width="107"/>
                    <TextBox Name="CustomAttribute1Textbox" HorizontalAlignment="Left" Height="23" Margin="343,159,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="140" TabIndex="8"/>
                    <Label Content="--Once you have filled in&#xD;&#xA;the required details, click &#xD;&#xA;Create User.&#xD;&#xA;--You will be prompted&#xD;&#xA;for Licenses and Groups to &#xD;&#xA;add.&#xD;&#xA;--The button will&#xD;&#xA;activate when the left &#xD;&#xA;side is filled in, the right&#xD;&#xA;side is not required&#xD;&#xA;for all tenants." HorizontalAlignment="Left" VerticalAlignment="Top" Margin="183,10,0,0" Height="203" Width="155"/>
                    <RichTextBox Name="CreateRichTextBox" HorizontalAlignment="Left" Height="87" Margin="10,362,0,0" VerticalAlignment="Top" Width="473" Background="#FF646464" Foreground="Cyan" IsReadOnly="True">
                        <FlowDocument/>
                    </RichTextBox>
                    <Button Name="CreateGoButton" Content="Create User" HorizontalAlignment="Left" Margin="183,218,0,0" VerticalAlignment="Top" Width="155" Height="139" IsEnabled="False" TabIndex="12"/>
                </Grid>
            </TabItem>
            <TabItem Name="TerminateTab" Header="Terminate User">
                <Grid Background="#FFE5E5E5">
                    <Label Content="Please Select Options Below for User Termination and press Terminate User.  You will&#xD;&#xA;be prompted to select a user and who to share to." HorizontalAlignment="Left" Margin="10,10,0,0" VerticalAlignment="Top" Width="473" Height="43"/>
                    <GroupBox Header="Share OneDrive?" HorizontalAlignment="Left" Height="80" Margin="243,67,0,0" VerticalAlignment="Top" Width="240">
                        <StackPanel HorizontalAlignment="Left" Height="74" Margin="10,10,-2,-13" VerticalAlignment="Top" Width="281">
                            <RadioButton Name="OneDriveNoRadioButton" Content="No" TabIndex="3"/>
                            <RadioButton Name="OneDriveSameRadioButton" Content="To Same User As Shared Mailbox" IsChecked="True" TabIndex="4"/>
                            <RadioButton Name="OneDriveDifferentRadioButton" Content="To Different User As Shared Mailbox" TabIndex="5"/>
                        </StackPanel>
                    </GroupBox>
                    <GroupBox Header="Standard Options" HorizontalAlignment="Left" Height="80" Margin="10,67,0,0" VerticalAlignment="Top" Width="228">
                        <StackPanel HorizontalAlignment="Left" Height="74" Margin="10,10,-2,-13" VerticalAlignment="Top" Width="208">
                            <CheckBox Name="ConvertCheckbox" Content="Convert to Shared Mailbox?" HorizontalAlignment="Left" VerticalAlignment="Top" RenderTransformOrigin="-1.855,-1.274" IsChecked="True" TabIndex="0"/>
                            <CheckBox Name="RemoveLicensesCheckbox" Content="Remove All Licenses?" HorizontalAlignment="Left" VerticalAlignment="Top" RenderTransformOrigin="-1.855,-1.274" IsChecked="True" TabIndex="1"/>
                            <CheckBox Name="ShareMailboxCheckbox" Content="Share the Mailbox?" HorizontalAlignment="Left" VerticalAlignment="Top" RenderTransformOrigin="-1.855,-1.274" IsChecked="True" TabIndex="2"/>
                        </StackPanel>
                    </GroupBox>
                    <Button Name="RemoveGoButton" Content="Terminate User" HorizontalAlignment="Left" Margin="10,397,0,0" VerticalAlignment="Top" Width="473" Height="52"/>
                    <RichTextBox Name="RemoveRichTextBox" HorizontalAlignment="Left" Height="240" Margin="10,152,0,0" VerticalAlignment="Top" Width="473" IsReadOnly="True" Background="#FF646464">
                        <FlowDocument/>
                    </RichTextBox>
                </Grid>
            </TabItem>
        </TabControl>
    </Grid>

</Window>
"@

$reader = (New-Object System.Xml.XmlNodeReader $xaml)
Try{
    $UserForm = [Windows.Markup.XamlReader]::Load($reader)
}
Catch{
    Write-Host "Unable to load Windows.Markup.XamlReader.  Some possible causes for this problem include: .NET Framework is missing PowerShell must be launched with PowerShell -sta, invalid XAML code was encountered."
    Exit
}

#Create Variables For Use In Script Automatically
$xaml.SelectNodes("//*[@Name]") | ForEach-Object {Set-Variable -Name ($_.Name) -Value $UserForm.FindName($_.Name)}
### End XAML and Variables from XAML

# Create Functions For Color Changing Output
Function Write-PasswordRichTextBox {
    Param(
        [string]$text,
        [string]$color = "Cyan"
    )
    $RichTextRange = New-Object System.Windows.Documents.TextRange( 
        $PasswordRichTextBox.Document.ContentEnd,$PasswordRichTextBox.Document.ContentEnd ) 
    $RichTextRange.Text = $text
    $RichTextRange.ApplyPropertyValue( ( [System.Windows.Documents.TextElement]::ForegroundProperty ), $color )  
}

Function Write-CreateRichTextBox {
    Param(
        [string]$text,
        [string]$color = "Cyan"
    )
    $RichTextRange = New-Object System.Windows.Documents.TextRange( 
        $CreateRichTextBox.Document.ContentEnd,$CreateRichTextBox.Document.ContentEnd ) 
    $RichTextRange.Text = $text
    $RichTextRange.ApplyPropertyValue( ( [System.Windows.Documents.TextElement]::ForegroundProperty ), $color )  
}

Function Write-RemoveRichTextBox {
    Param(
        [string]$text,
        [string]$color = "Cyan"
    )
    $RichTextRange = New-Object System.Windows.Documents.TextRange( 
        $RemoveRichTextBox.Document.ContentEnd,$RemoveRichTextBox.Document.ContentEnd ) 
    $RichTextRange.Text = $text
    $RichTextRange.ApplyPropertyValue( ( [System.Windows.Documents.TextElement]::ForegroundProperty ), $color )  
}

### Start Password Tab Functionality
$PasswordTextBox.Add_TextChanged({
    if (($PasswordTextBox.Text.Length -ge 8) -and ($UserTextBox.Text.Length -ge 2)){
        $PasswordGoButton.IsEnabled = $true
    }
    else{
        $PasswordGoButton.IsEnabled = $false
    }
})

$UserButton.Add_Click({
    $tempuser = Get-AzureADUser -all $true | Out-GridView -Outputmode Single
    $UserTextBox.Text = $tempuser.UserPrincipalName
})

$PasswordGoButton.Add_Click({
    $securepassword = ConvertTo-SecureString -String $PasswordTextBox.Text -AsPlainText -Force
    Try{
        Set-AzureADUserPassword -ObjectID $UserTextBox.Text -Password $securepassword -ForceChangePasswordNextLogin $false -ErrorAction Stop
        Write-PasswordRichTextBox("SUCCESS:  $($UserTextBox.Text)'s password has been reset to $($PasswordTextBox.Text)`r")
        $PasswordRichTextBox.ScrollToEnd()
        $UserTextbox.Text = ""
        $PasswordTextBox.Text = ""
    }Catch{
        $message = $_.Exception.Message
        if ($_.Exception.ErrorContent.Message.Value) {
            $message = $_.Exception.ErrorContent.Message.Value
        }
        Write-PasswordRichTextBox("$message`rFAILURE:  Please review above and try again`r") -Color "Red"
        $PasswordRichTextBox.ScrollToEnd()
    }
})
### End Password Tab Functionality

### Start User Creation Tab Functionality



### End User Creation Tab Functionality

### Start User Termination Tab Functionality
#Test And Connect To Microsoft Exchange Online If Needed
$RemoveGoButton.Add_Click({
    Try {
        Get-Mailbox -ErrorAction Stop | Out-Null
    }Catch {
        Connect-ExchangeOnline
    }

    #Pull All Azure AD Users and Store In Hash Table Instead Of Calling Get-AzureADUser Multiple Times
    Write-RemoveRichTextBox("Pulling Users To Store In a Hash Table")
    $allUsers = @{}    
    foreach ($user in Get-AzureADUser -All $true){ $allUsers[$user.UserPrincipalName] = $user }
    Write-RemoveRichTextBox("Hash Table Filled")

    #Request Username(s) To Be Terminated From Script Runner (Hold Ctrl To Select Multiples)
    $usernames = $allUsers.Values | Where-Object {$_.AccountEnabled } | Sort-Object DisplayName | Select-Object -Property DisplayName,UserPrincipalName | Out-Gridview -Passthru -Title "Please select the user(s) to be terminated" | Select-Object -ExpandProperty UserPrincipalName
    
    ##### Start User(s) Loop #####
    foreach ($username in $usernames) {
        $UserInfo = $allusers[$username]
        #Request User(s) To Share Mailbox With When Grant Access Is Selected
        if ($GrantMailboxCheckBox.Checked -eq $true) {
            $sharedMailboxUser = $allUsers.Values | Where-Object {$_.AccountEnabled } | Sort-Object DisplayName | Select-Object -Property DisplayName,UserPrincipalName | Out-GridView -Title "Please select the user(s) to share the $username Shared Mailbox with" -OutputMode Single | Select-Object -ExpandProperty UserPrincipalName
        }
    }
    
    #Block Sign In Of User/Force Sign Out Within 60 Minutes
    Set-AzureADUser -ObjectID $UserInfo.ObjectId -AccountEnabled $false
    Write-RemoveRichTextBox("Sign in Blocked for $($UserInfo.ObjectID)")

    #Remove All Group Memberships
    Write-RemoveRichTextBox("Removing all group memberships, skipping Dynamic groups as they cannot be removed this way")
    $memberships = Get-AzureADUserMembership -ObjectId $username | Where-Object {$_.ObjectType -ne "Role"}| Select-Object DisplayName,ObjectId
    foreach ($membership in $memberships) { 
            $group = Get-AzureADMSGroup -ID $membership.ObjectId
            if ($group.GroupTypes -contains 'DynamicMembership') {
                Write-RemoveRichTextBox("Skipping $($group.Displayname) as it is dynamic")
            }
            else{
                Try{
                    Remove-AzureADGroupMember -ObjectId $membership.ObjectId -MemberId $UserInfo.ObjectId -ErrorAction Stop
                }Catch{
                    Write-RemoveRichTextBox("Could not remove from group $($group.name).  Error:  $_.Message") -color "Yellow"
                }
            }
        }
    Write-RemoveRichTextBox("All non-dynamic groups removed, please check your Downloads folder for the file, it will also open automatically at end of user termination")

    #Convert To Shared Mailbox And Hide From GAL When Convert Is Selected, Must Be Done Before Removing Licenses
    if ($ConvertCheckBox.Checked -eq $true) {
        Write-RemoveRichTextBox("Converting $username to Shared Mailbox and Hiding from GAL")
        Set-Mailbox $username -Type Shared -HiddenFromAddressListsEnabled $true
        Write-RemoveRichTextBox("Mailbox for $username converted to Shared, address hidden from GAL")
    }

    #Grant Access To Shared Mailbox When Grant CheckBox Is Selected
    if ($ShareMailboxCheckBox.Checked -eq $true) {
        Add-MailboxPermission -Identity $username -User $SharedMailboxUser -AccessRights FullAccess -InheritanceType All
        Add-RecipientPermission -Identity $username -Trustee $SharedMailboxUser -AccessRights SendAs -Confirm:$False
        Write-Verbose -Message "Access granted to the $username Shared Mailbox to $sharedMailboxUser"
    }

    #Remove All Licenses When Remove Licenses Is Selected
    if ($RemoveLicensesCheckBox.Checked -eq $true) {
        $licenses = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicenses
        if($UserInfo.assignedlicenses){
            $licenses.RemoveLicenses = $UserInfo.assignedlicenses.SkuId
            Set-AzureADUserLicense -ObjectId $UserInfo.ObjectId -AssignedLicenses $licenses
        }
        Write-RemoveRichTextBox("All licenses have been removed")
    }

    #Test And Connect To Sharepoint Online If Needed
    if ($OneDriveNoRadioButton.IsChecked -ne $true) {
        $domainPrefix = ((Get-AzureADDomain | Where-Object Name -match "\.onmicrosoft\.com")[0].Name -split '\.')[0]
        $AdminSiteUrl = "https://$domainPrefix-admin.sharepoint.com"
        Try{
            Get-SPOSite -ErrorAction Stop | Out-Null
        }Catch{
            Write-Verbose -Message "Connecting to SharePoint Online"
            Connect-SPOService -Url $AdminSiteURL
        }
    }

    #Share OneDrive With Same User as Shared Mailbox
    if ($OneDriveSameRadioButton.IsChecked -eq $true) {
        #Pull OneDriveSiteURL Dynamically And Grant Access
        $OneDriveSiteURL = Get-SPOSite -Filter "Owner -eq $($UserInfo.UserPrincipalName)" -IncludePersonalSite $true | Select-Object -ExpandProperty Url            

        #Add User Receiving Access To Terminated User's OneDrive, Add The Access Link To CSV File For Copying
        Set-SPOUser -Site $OneDriveSiteUrl -LoginName $SharedMailboxUser -IsSiteCollectionAdmin $True
        Write-RemoveRichTextBox("OneDrive Data Shared with $SharedMailboxUser successfully, link to copy and give to Manager is $OneDriveSiteURL")
    }
    #Share OneDrive With Different User(s) than Shared Mailbox
    elseif ($OneDriveDifferentRadioButton.IsChecked -eq $true) {
        $SharedOneDriveUser = $allusers.Values | Where-Object {$_.AccountEnabled } | Sort-Object Displayname | Select-Object -Property DisplayName,UserPrincipalName | Out-GridView -Title "Please select the user(s) to share the Mailbox and OneDrive with" -OutputMode Single | Select-Object -ExpandProperty UserPrincipalName
        $SharedOneDriveUser = $allusers.Values | Sort-Object Displayname | Select-Object -Property DisplayName,UserPrincipalName | Out-GridView -Title "Please select the user to share the OneDrive with" -OutputMode Single | Select-Object -ExpandProperty UserPrincipalName
        
        #Pull Object ID Needed For User Receiving Access To OneDrive And OneDriveSiteURL Dynamically
        $OneDriveSiteURL = Get-SPOSite -Filter "Owner -eq $($UserInfo.UserPrincipalName)" -IncludePersonalSite $true | Select-Object -ExpandProperty Url            

        #Add User Receiving Access To Terminated User's OneDrive, Add The Access Link To CSV File For Copying
        Write-Verbose -Message "Adding $SharedOneDriveUser to OneDrive folder for access to files"
        Set-SPOUser -Site $OneDriveSiteUrl -LoginName $SharedOneDriveUser -IsSiteCollectionAdmin $True
        Write-Verbose "OneDrive Data Shared with $SharedOneDriveUser successfully, link to copy and provide to trustee is $OneDriveSiteURL"
    }

    #Export Groups Removed and OneDrive URL to CSV
    [pscustomobject]@{
        GroupsRemoved    = $memberships.DisplayName -join ','
        OneDriveSiteURL = $OneDriveSiteURL
    } | Export-Csv -Path c:\users\$env:USERNAME\Downloads\$(get-date -f yyyy-MM-dd)_info_on_$username.csv -NoTypeInformation

    #Open Created CSV File At End Of Loop For Ease Of Copying OneDrive URL To Give
    Start-Process c:\users\$env:USERNAME\Downloads\$(get-date -f yyyy-MM-dd)_info_on_$username.csv
})

$ConvertCheckbox.Add_Checked({
    $ShareMailboxCheckBox.IsEnabled = $true
})

$ConvertCheckbox.Add_Unchecked({
    $ShareMailboxCheckBox.IsChecked = $false
    $ShareMailboxCheckBox.IsEnabled = $false
})

$ShareMailboxCheckBox.Add_Checked({ 
    $OneDriveSameRadioButton.IsEnabled = $true
})

$ShareMailboxCheckBox.Add_Unchecked({
    $OneDriveSameRadioButton.IsEnabled = $false
    $OneDriveNoRadioButton.IsChecked = $true
})



### End User Termination Tab Functionality


$UserForm.Add_Loaded({
    Try{
    Get-AzureADUser -ErrorAction Stop | Out-Null
    }Catch{
        Connect-AzureAD
    }
})

$null = $UserForm.ShowDialog()
