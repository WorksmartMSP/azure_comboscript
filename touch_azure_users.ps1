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
        <Grid.RowDefinitions>
            <RowDefinition/>
        </Grid.RowDefinitions>
        <TabControl Name="Tabs" HorizontalAlignment="Left" Height="487" Margin="10,0,0,0" VerticalAlignment="Top" Width="499">
            <TabItem Name="ResetTab" Header="Reset Password">
                <Grid Background="#FFE5E5E5">
                    <Label Content="Please Pick A User, Then Enter A Password" HorizontalAlignment="Left" Margin="10,10,0,0" VerticalAlignment="Top" Height="25" Width="243"/>
                    <TextBox Name="PasswordTextBox" HorizontalAlignment="Left" Height="25" Margin="10,130,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="473" TabIndex="1"/>
                    <Button Name="PasswordGoButton" Content="Go" HorizontalAlignment="Left" Margin="10,399,0,0" VerticalAlignment="Top" Width="473" Height="50" IsEnabled="False" TabIndex="0"/>
                    <TextBox Name="UserTextBox" HorizontalAlignment="Left" Height="23" Margin="258,55,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="225" IsReadOnly="True" Background="#FFC8C8C8"/>
                    <Button Name="UserButton" Content="Pick User" HorizontalAlignment="Left" Margin="10,40,0,0" VerticalAlignment="Top" Width="243" Height="54"/>
                    <Label Content="Enter Password Below, Go Button Activates At 8 Characters and User Selected" HorizontalAlignment="Left" Margin="10,99,0,0" VerticalAlignment="Top" Width="473"/>
                    <RichTextBox Name="PasswordRichTextBox" HorizontalAlignment="Left" Height="214" Margin="10,180,0,0" VerticalAlignment="Top" Width="473" Background="#FF646464" Foreground="Cyan" HorizontalScrollBarVisibility="Auto" VerticalScrollBarVisibility="Auto" IsReadOnly="True">
                        <FlowDocument/>
                    </RichTextBox>
                    <CheckBox Name="PasswordResetCheckbox" Content="Force Reset of Password on Login?" HorizontalAlignment="Left" Margin="10,160,0,0" VerticalAlignment="Top" RenderTransformOrigin="-1.221,-1.703" Width="461"/>
                    <Button Name="PasswordReconnectButton" Content="Reconnect/Change Tenants" HorizontalAlignment="Left" Margin="258,10,0,0" VerticalAlignment="Top" Width="225" Height="25"/>
                </Grid>
            </TabItem>
            <TabItem Name="MailboxTab" Header="Mailboxes">
                <Grid Background="#FFE5E5E5">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition/>
                    </Grid.ColumnDefinitions>
                    <Button Name="MailboxReconnectButton" Content="Reconnect/Change Tenants" HorizontalAlignment="Left" Margin="10,10,0,0" VerticalAlignment="Top" Width="473" Height="50"/>
                    <Label Content="When you click Add or Remove, you may select multiple users by holding control.  The&#xD;&#xA;same applies to the mailboxes, but it will impact ALL the users you have selected.&#xD;&#xA;Please use caution if you have multiple users with different needed mailboxes." HorizontalAlignment="Left" Margin="10,65,0,0" VerticalAlignment="Top" Height="65" Width="473"/>
                    <Button Name="MailboxAddButton" Content="Add User(s) to Mailbox(es)" HorizontalAlignment="Left" Margin="10,349,0,0" VerticalAlignment="Top" Width="230" Height="100"/>
                    <Button Name="MailboxRemoveButton" Content="Remove User(s) from Mailbox(es)" HorizontalAlignment="Left" Margin="253,349,0,0" VerticalAlignment="Top" Width="230" Height="100"/>
                    <RichTextBox Name="MailboxRichTextBox" HorizontalAlignment="Left" Height="209" Margin="10,135,0,0" VerticalAlignment="Top" Width="473" Foreground="Cyan" Background="#FF646464" IsReadOnly="True">
                        <FlowDocument/>
                    </RichTextBox>
                </Grid>
            </TabItem>
            <TabItem Name="GroupTab" Header="Groups">
                <Grid Background="#FFE5E5E5">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition/>
                    </Grid.ColumnDefinitions>
                    <Button Name="GroupReconnectButton" Content="Reconnect/Change Tenants" HorizontalAlignment="Left" Margin="10,10,0,0" VerticalAlignment="Top" Width="473" Height="50"/>
                    <Label Content="When you click Add or Remove, you may select multiple users by holding control.  The&#xA;same applies to the groups, but it will impact ALL the users you have selected.&#xA;Please use caution if you have multiple users with different needed groups." HorizontalAlignment="Left" Margin="10,65,0,0" VerticalAlignment="Top" Height="65" Width="473"/>
                    <Button Name="GroupRemoveButton" Content="Remove User(s) from Group(s)" HorizontalAlignment="Left" Margin="253,349,0,0" VerticalAlignment="Top" Width="230" Height="100"/>
                    <Button Name="GroupAddButton" Content="Add User(s) to Group(s)" HorizontalAlignment="Left" Margin="10,349,0,0" VerticalAlignment="Top" Width="230" Height="100"/>
                    <RichTextBox Name="GroupRichTextBox" HorizontalAlignment="Left" Height="209" Margin="10,135,0,0" VerticalAlignment="Top" Width="473" IsReadOnly="True" Background="#FF646464">
                        <FlowDocument/>
                    </RichTextBox>
                </Grid>
            </TabItem>
            <TabItem Name="CalendarTab" Header="Calendars">
                <Grid Background="#FFE5E5E5">
                    <Button Name="CalendarReconnectButton" Content="Reconnect/Change Tenants" HorizontalAlignment="Left" Margin="10,10,0,0" VerticalAlignment="Top" Width="473" Height="20"/>
                    <GroupBox Header="Select Needed Permissions - http://worksmart.link/7f for Permissions Information" HorizontalAlignment="Left" Height="171" Margin="10,108,0,0" VerticalAlignment="Top" Width="473">
                        <StackPanel HorizontalAlignment="Left" Height="151" Margin="10,10,-2,-13" VerticalAlignment="Top" Width="453">
                            <RadioButton Name="AuthorRadioButton" Content="Author - CI, DOI, EOI, FV, RI"/>
                            <RadioButton Name="ContributorRadioButton" Content="Contributor - CI, FV"/>
                            <RadioButton Name="EditorRadioButton" Content="Editor - CI, DAI, DOI, EAI, EOI, FV, RI"/>
                            <RadioButton Name="NoneRadioButton" Content="None (View Only) - FV" IsChecked="True"/>
                            <RadioButton Name="NonEditingAuthorRadioButton" Content="NonEditing Author - CI, DOI, FV, RI"/>
                            <RadioButton Name="OwnerRadioButton" Content="Owner - CI, CS, DAI, DOI, EAI, EOI, FC, FO, FV, RI"/>
                            <RadioButton Name="PublishingAuthorRadioButton" Content="Publishing Author - CI, CS, DOI, EOI, FV, RI"/>
                            <RadioButton Name="PublishingEditorRadioButton" Content="Publishing Editor - CI, CS, DAI, DOI, EAI, EOI, FV, RI"/>
                            <RadioButton Name="ReviewerRadioButton" Content="Reviewer - FV, RI"/>
                        </StackPanel>
                    </GroupBox>
                    <Button Name="CalendarGoButton" Content="Update User(s)' Permissions on Calendar(s)" HorizontalAlignment="Left" Margin="10,399,0,0" VerticalAlignment="Top" Width="473" Height="50"/>
                    <RichTextBox Name="CalendarRichTextBox" HorizontalAlignment="Left" Height="90" Margin="10,304,0,0" VerticalAlignment="Top" Width="473" Background="#FF646464" IsReadOnly="True">
                        <FlowDocument/>
                    </RichTextBox>
                    <CheckBox Name="AvailabilityOnlyCheckbox" Content="Availability Only" HorizontalAlignment="Left" Margin="10,284,0,0" VerticalAlignment="Top"/>
                    <CheckBox Name="LimitedDetailsCheckbox" Content="LimitedDetails - Subject/Location" HorizontalAlignment="Left" Margin="288,284,0,0" VerticalAlignment="Top"/>
                    <Label Content="The permissions below are shorthand, see link for more details.&#xD;&#xA;Create Items, Create Subfolders, Delete All Items, Delete Owned Items, Edit All Items,&#xD;&#xA;Edit Owned Items, Folder Contact, Folder Owner, Folder Visible, Read Items" HorizontalAlignment="Left" Margin="10,35,0,0" VerticalAlignment="Top" Height="68" Width="473"/>
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
                    <Label Content="City" HorizontalAlignment="Left" Margin="343,41,0,0" VerticalAlignment="Top" Width="57"/>
                    <TextBox Name="CityTextbox" HorizontalAlignment="Left" Height="23" Margin="343,72,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="140" TabIndex="7"/>
                    <Label Content="State" HorizontalAlignment="Left" Margin="343,100,0,0" VerticalAlignment="Top" Width="57"/>
                    <TextBox Name="StateTextbox" HorizontalAlignment="Left" Height="23" Margin="343,131,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="140" TabIndex="8"/>
                    <Label Content="CustomAttribute1" HorizontalAlignment="Left" Margin="343,218,0,0" VerticalAlignment="Top" Width="107"/>
                    <TextBox Name="CustomAttribute1Textbox" HorizontalAlignment="Left" Height="23" Margin="343,249,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="140" TabIndex="10"/>
                    <Label Content="--Once you have filled in&#xA;the required details, click &#xA;Create User.&#xA;&#xD;&#xA;--The button will&#xA;activate when the left is&#xD;&#xA;filled in, the right is not&#xD;&#xA;required for all tenants.&#xD;&#xA;&#xD;&#xA;--You will be prompted to &#xD;&#xA;connect to Exchange&#xD;&#xA;Online, then for Licenses&#xD;&#xA;and Groups to add." HorizontalAlignment="Left" VerticalAlignment="Top" Margin="183,41,0,0" Height="230" Width="155"/>
                    <RichTextBox Name="CreateRichTextBox" HorizontalAlignment="Left" Height="67" Margin="10,382,0,0" VerticalAlignment="Top" Width="473" Background="#FF646464" Foreground="Cyan" IsReadOnly="True">
                        <FlowDocument/>
                    </RichTextBox>
                    <Button Name="CreateGoButton" Content="Create User" HorizontalAlignment="Left" Margin="183,277,0,0" VerticalAlignment="Top" Width="300" Height="100" IsEnabled="False" TabIndex="11"/>
                    <Label Content="Country" HorizontalAlignment="Left" Margin="343,159,0,0" VerticalAlignment="Top"/>
                    <TextBox Name="CountryTextbox" HorizontalAlignment="Left" Height="23" Margin="343,190,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="140" TabIndex="9"/>
                    <Button Name="CreateReconnectButton" Content="Reconnect/Change Tenants" HorizontalAlignment="Left" Margin="183,10,0,0" VerticalAlignment="Top" Width="300" Height="26"/>
                    <CheckBox Name="CreateResetPasswordCheckbox" Content="Reset Password on Login?" HorizontalAlignment="Left" Margin="10,362,0,0" VerticalAlignment="Top" TabIndex="6"/>
                </Grid>
            </TabItem>
            <TabItem Name="TerminateTab" Header="Terminate User">
                <Grid Background="#FFE5E5E5">
                    <Label Content="Please Select Options Below for User Termination and press Terminate User.  You will&#xD;&#xA;be prompted to select a user and who to share to." HorizontalAlignment="Left" Margin="10,65,0,0" VerticalAlignment="Top" Width="473" Height="43"/>
                    <GroupBox Header="Share OneDrive?" HorizontalAlignment="Left" Height="80" Margin="243,113,0,0" VerticalAlignment="Top" Width="240">
                        <StackPanel HorizontalAlignment="Left" Height="74" Margin="10,10,-2,-13" VerticalAlignment="Top" Width="281">
                            <RadioButton Name="OneDriveNoRadioButton" Content="No" TabIndex="3"/>
                            <RadioButton Name="OneDriveSameRadioButton" Content="To Same User As Shared Mailbox" IsChecked="True" TabIndex="4"/>
                            <RadioButton Name="OneDriveDifferentRadioButton" Content="To Different User As Shared Mailbox" TabIndex="5"/>
                        </StackPanel>
                    </GroupBox>
                    <GroupBox Header="Standard Options" HorizontalAlignment="Left" Height="80" Margin="10,113,0,0" VerticalAlignment="Top" Width="228">
                        <StackPanel HorizontalAlignment="Left" Height="74" Margin="10,10,-2,-13" VerticalAlignment="Top" Width="208">
                            <CheckBox Name="ConvertCheckbox" Content="Convert to Shared Mailbox?" HorizontalAlignment="Left" VerticalAlignment="Top" RenderTransformOrigin="-1.855,-1.274" IsChecked="True" TabIndex="0"/>
                            <CheckBox Name="RemoveLicensesCheckbox" Content="Remove All Licenses?" HorizontalAlignment="Left" VerticalAlignment="Top" RenderTransformOrigin="-1.855,-1.274" IsChecked="True" TabIndex="1"/>
                            <CheckBox Name="ShareMailboxCheckbox" Content="Share the Mailbox?" HorizontalAlignment="Left" VerticalAlignment="Top" RenderTransformOrigin="-1.855,-1.274" IsChecked="True" TabIndex="2"/>
                        </StackPanel>
                    </GroupBox>
                    <Button Name="RemoveGoButton" Content="Terminate User" HorizontalAlignment="Left" Margin="10,397,0,0" VerticalAlignment="Top" Width="473" Height="52"/>
                    <RichTextBox Name="RemoveRichTextBox" HorizontalAlignment="Left" Height="194" Margin="10,198,0,0" VerticalAlignment="Top" Width="473" IsReadOnly="True" Background="#FF646464">
                        <FlowDocument/>
                    </RichTextBox>
                    <Button Name="TerminateReconnectButton" Content="Reconnect/ChangeTenants" HorizontalAlignment="Left" Margin="10,10,0,0" VerticalAlignment="Top" Width="473" Height="50"/>
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
Function Write-MailboxRichTextBox {
    Param(
        [string]$text,
        [string]$color = "Cyan"
    )
    $RichTextRange = New-Object System.Windows.Documents.TextRange( 
        $MailboxRichTextBox.Document.ContentEnd,$MailboxRichTextBox.Document.ContentEnd ) 
    $RichTextRange.Text = $text
    $RichTextRange.ApplyPropertyValue( ( [System.Windows.Documents.TextElement]::ForegroundProperty ), $color )  
}

Function Write-GroupRichTextBox {
    Param(
        [string]$text,
        [string]$color = "Cyan"
    )
    $RichTextRange = New-Object System.Windows.Documents.TextRange( 
        $GroupRichTextBox.Document.ContentEnd,$GroupRichTextBox.Document.ContentEnd ) 
    $RichTextRange.Text = $text
    $RichTextRange.ApplyPropertyValue( ( [System.Windows.Documents.TextElement]::ForegroundProperty ), $color )  
}

Function Write-CalendarRichTextBox {
    Param(
        [string]$text,
        [string]$color = "Cyan"
    )
    $RichTextRange = New-Object System.Windows.Documents.TextRange( 
        $CalendarRichTextBox.Document.ContentEnd,$CalendarRichTextBox.Document.ContentEnd ) 
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

Function Set-Comboboxes {
    foreach($UsageLocation in $UsageLocations.keys)
    {
        $null = $UsageLocationComboBox.Items.Add($usagelocation)
    }
    $UsageLocationComboBox.SelectedIndex = 0
    
    $domainComboBox.Items.Clear()
    foreach($domain in Get-AzureADDomain){
        $null = $domainComboBox.Items.add($domain.Name)
    }
    $DomainComboBox.SelectedIndex = 0
}

# Friendly Name Lookup Table
$SkuToFriendly = @{
    "c42b9cae-ea4f-4ab7-9717-81576235ccac" = "DevPack E5 (No Windows or Audio)"
    "8f0c5670-4e56-4892-b06d-91c085d7004f" = "APP CONNECT IW"
    "0c266dff-15dd-4b49-8397-2bb16070ed52" = "Microsoft 365 Audio Conferencing"
    "2b9c8e7c-319c-43a2-a2a0-48c5c6161de7" = "AZURE ACTIVE DIRECTORY BASIC"
    "078d2b04-f1bd-4111-bbd4-b4b1b354cef4" = "AZURE ACTIVE DIRECTORY PREMIUM P1"
    "84a661c4-e949-4bd2-a560-ed7766fcaf2b" = "AZURE ACTIVE DIRECTORY PREMIUM P2"
    "c52ea49f-fe5d-4e95-93ba-1de91d380f89" = "AZURE INFORMATION PROTECTION PLAN 1"
    "295a8eb0-f78d-45c7-8b5b-1eed5ed02dff" = "COMMON AREA PHONE"
    "47794cd0-f0e5-45c5-9033-2eb6b5fc84e0" = "COMMUNICATIONS CREDITS"
    "ea126fc5-a19e-42e2-a731-da9d437bffcf" = "DYNAMICS 365 CUSTOMER ENGAGEMENT PLAN ENTERPRISE EDITION"
    "749742bf-0d37-4158-a120-33567104deeb" = "DYNAMICS 365 FOR CUSTOMER SERVICE ENTERPRISE EDITION"
    "cc13a803-544e-4464-b4e4-6d6169a138fa" = "DYNAMICS 365 FOR FINANCIALS BUSINESS EDITION"
    "8edc2cf8-6438-4fa9-b6e3-aa1660c640cc" = "DYNAMICS 365 FOR SALES AND CUSTOMER SERVICE ENTERPRISE EDITION"
    "1e1a282c-9c54-43a2-9310-98ef728faace" = "DYNAMICS 365 FOR SALES ENTERPRISE EDITION"
    "f2e48cb3-9da0-42cd-8464-4a54ce198ad0" = "DYNAMICS 365 FOR SUPPLY CHAIN MANAGEMENT"
    "8e7a3d30-d97d-43ab-837c-d7701cef83dc" = "DYNAMICS 365 FOR TEAM MEMBERS ENTERPRISE EDITION"
    "338148b6-1b11-4102-afb9-f92b6cdc0f8d" = "DYNAMICS 365 P1 TRIAL FOR INFORMATION WORKERS"
    "b56e7ccc-d5c7-421f-a23b-5c18bdbad7c0" = "DYNAMICS 365 TALENT: ONBOARD"
    "7ac9fe77-66b7-4e5e-9e46-10eed1cff547" = "DYNAMICS 365 TEAM MEMBERS"
    "ccba3cfe-71ef-423a-bd87-b6df3dce59a9" = "DYNAMICS 365 UNF OPS PLAN ENT EDITION"
    "efccb6f7-5641-4e0e-bd10-b4976e1bf68e" = "ENTERPRISE MOBILITY + SECURITY E3"
    "b05e124f-c7cc-45a0-a6aa-8cf78c946968" = "ENTERPRISE MOBILITY + SECURITY E5"
    "4b9405b0-7788-4568-add1-99614e613b69" = "EXCHANGE ONLINE (PLAN 1)"
    "19ec0d23-8335-4cbd-94ac-6050e30712fa" = "EXCHANGE ONLINE (PLAN 2)"
    "ee02fd1b-340e-4a4b-b355-4a514e4c8943" = "EXCHANGE ONLINE ARCHIVING FOR EXCHANGE ONLINE"
    "90b5e015-709a-4b8b-b08e-3200f994494c" = "EXCHANGE ONLINE ARCHIVING FOR EXCHANGE SERVER"
    "7fc0182e-d107-4556-8329-7caaa511197b" = "EXCHANGE ONLINE ESSENTIALS (ExO P1 BASED)"
    "e8f81a67-bd96-4074-b108-cf193eb9433b" = "EXCHANGE ONLINE ESSENTIALS"
    "80b2d799-d2ba-4d2a-8842-fb0d0f3a4b82" = "EXCHANGE ONLINE KIOSK"
    "cb0a98a8-11bc-494c-83d9-c1b1ac65327e" = "EXCHANGE ONLINE POP"
    "061f9ace-7d42-4136-88ac-31dc755f143f" = "INTUNE"
    "b17653a4-2443-4e8c-a550-18249dda78bb" = "Microsoft 365 A1"
    "4b590615-0888-425a-a965-b3bf7789848d" = "MICROSOFT 365 A3 FOR FACULTY"
    "7cfd9a2b-e110-4c39-bf20-c6a3f36a3121" = "MICROSOFT 365 A3 FOR STUDENTS"
    "e97c048c-37a4-45fb-ab50-922fbf07a370" = "MICROSOFT 365 A5 FOR FACULTY"
    "46c119d4-0379-4a9d-85e4-97c66d3f909e" = "MICROSOFT 365 A5 FOR STUDENTS"
    "cdd28e44-67e3-425e-be4c-737fab2899d3" = "MICROSOFT 365 APPS FOR BUSINESS"
    "b214fe43-f5a3-4703-beeb-fa97188220fc" = "MICROSOFT 365 APPS FOR BUSINESS"
    "c2273bd0-dff7-4215-9ef5-2c7bcfb06425" = "MICROSOFT 365 APPS FOR ENTERPRISE"
    "2d3091c7-0712-488b-b3d8-6b97bde6a1f5" = "MICROSOFT 365 AUDIO CONFERENCING FOR GCC"
    "3b555118-da6a-4418-894f-7df1e2096870" = "MICROSOFT 365 BUSINESS BASIC"
    "dab7782a-93b1-4074-8bb1-0e61318bea0b" = "MICROSOFT 365 BUSINESS BASIC"
    "f245ecc8-75af-4f8e-b61f-27d8114de5f3" = "MICROSOFT 365 BUSINESS STANDARD"
    "ac5cef5d-921b-4f97-9ef3-c99076e5470f" = "MICROSOFT 365 BUSINESS STANDARD - PREPAID LEGACY"
    "cbdc14ab-d96c-4c30-b9f4-6ada7cdc1d46" = "MICROSOFT 365 BUSINESS PREMIUM"
    "11dee6af-eca8-419f-8061-6864517c1875" = "MICROSOFT 365 DOMESTIC CALLING PLAN (120 Minutes)"
    "05e9a617-0261-4cee-bb44-138d3ef5d965" = "MICROSOFT 365 E3"
    "06ebc4ee-1bb5-47dd-8120-11324bc54e06" = "Microsoft 365 E5"
    "d61d61cc-f992-433f-a577-5bd016037eeb" = "Microsoft 365 E3_USGOV_DOD"
    "ca9d1dd9-dfe9-4fef-b97c-9bc1ea3c3658" = "Microsoft 365 E3_USGOV_GCCHIGH"
    "184efa21-98c3-4e5d-95ab-d07053a96e67" = "Microsoft 365 E5 Compliance"
    "26124093-3d78-432b-b5dc-48bf992543d5" = "Microsoft 365 E5 Security"
    "44ac31e7-2999-4304-ad94-c948886741d4" = "Microsoft 365 E5 Security for EMS E5"
    "44575883-256e-4a79-9da4-ebe9acabe2b2" = "Microsoft 365 F1"
    "66b55226-6b4f-492c-910c-a3b7a3c9d993" = "Microsoft 365 F3"
    "f30db892-07e9-47e9-837c-80727f46fd3d" = "MICROSOFT FLOW FREE"
    "e823ca47-49c4-46b3-b38d-ca11d5abe3d2" = "MICROSOFT 365 G3 GCC"
    "e43b5b99-8dfb-405f-9987-dc307f34bcbd" = "MICROSOFT 365 PHONE SYSTEM"
    "d01d9287-694b-44f3-bcc5-ada78c8d953e" = "MICROSOFT 365 PHONE SYSTEM FOR DOD"
    "d979703c-028d-4de5-acbf-7955566b69b9" = "MICROSOFT 365 PHONE SYSTEM FOR FACULTY"
    "a460366a-ade7-4791-b581-9fbff1bdaa85" = "MICROSOFT 365 PHONE SYSTEM FOR GCC"
    "7035277a-5e49-4abc-a24f-0ec49c501bb5" = "MICROSOFT 365 PHONE SYSTEM FOR GCCHIGH"
    "aa6791d3-bb09-4bc2-afed-c30c3fe26032" = "MICROSOFT 365 PHONE SYSTEM FOR SMALL AND MEDIUM BUSINESS"
    "1f338bbc-767e-4a1e-a2d4-b73207cc5b93" = "MICROSOFT 365 PHONE SYSTEM FOR STUDENTS"
    "ffaf2d68-1c95-4eb3-9ddd-59b81fba0f61" = "MICROSOFT 365 PHONE SYSTEM FOR TELSTRA"
    "b0e7de67-e503-4934-b729-53d595ba5cd1" = "MICROSOFT 365 PHONE SYSTEM_USGOV_DOD"
    "985fcb26-7b94-475b-b512-89356697be71" = "MICROSOFT 365 PHONE SYSTEM_USGOV_GCCHIGH"
    "440eaaa8-b3e0-484b-a8be-62870b9ba70a" = "MICROSOFT 365 PHONE SYSTEM - VIRTUAL USER"
    "2347355b-4e81-41a4-9c22-55057a399791" = "MICROSOFT 365 SECURITY AND COMPLIANCE FOR FLW"
    "726a0894-2c77-4d65-99da-9775ef05aad1" = "MICROSOFT BUSINESS CENTER"
    "111046dd-295b-4d6d-9724-d52ac90bd1f2" = "MICROSOFT DEFENDER FOR ENDPOINT"
    "906af65a-2970-46d5-9b58-4e9aa50f0657" = "MICROSOFT DYNAMICS CRM ONLINE BASIC"
    "d17b27af-3f49-4822-99f9-56a661538792" = "MICROSOFT DYNAMICS CRM ONLINE"
    "ba9a34de-4489-469d-879c-0f0f145321cd" = "MS IMAGINE ACADEMY"
    "2c21e77a-e0d6-4570-b38a-7ff2dc17d2ca" = "MICROSOFT INTUNE DEVICE FOR GOVERNMENT"
    "dcb1a3ae-b33f-4487-846a-a640262fadf4" = "MICROSOFT POWER APPS PLAN 2 TRIAL"
    "e6025b08-2fa5-4313-bd0a-7e5ffca32958" = "MICROSOFT INTUNE SMB"
    "1f2f344a-700d-42c9-9427-5cea1d5d7ba6" = "MICROSOFT STREAM"
    "16ddbbfc-09ea-4de2-b1d7-312db6112d70" = "MICROSOFT TEAM (FREE)"
    "710779e8-3d4a-4c88-adb9-386c958d1fdf" = "MICROSOFT TEAMS EXPLORATORY"
    "a4585165-0533-458a-97e3-c400570268c4" = "Office 365 A5 for faculty"
    "ee656612-49fa-43e5-b67e-cb1fdf7699df" = "Office 365 A5 for students"
    "1b1b1f7a-8355-43b6-829f-336cfccb744c" = "Office 365 Advanced Compliance"
    "4ef96642-f096-40de-a3e9-d83fb2f90211" = "Microsoft Defender for Office 365 (Plan 1)"
    "18181a46-0d4e-45cd-891e-60aabd171b4e" = "OFFICE 365 E1"
    "6634e0ce-1a9f-428c-a498-f84ec7b8aa2e" = "OFFICE 365 E2"
    "6fd2c87f-b296-42f0-b197-1e91e994b900" = "OFFICE 365 E3"
    "189a915c-fe4f-4ffa-bde4-85b9628d07a0" = "OFFICE 365 E3 DEVELOPER"
    "b107e5a3-3e60-4c0d-a184-a7e4395eb44c" = "Office 365 E3_USGOV_DOD"
    "aea38a85-9bd5-4981-aa00-616b411205bf" = "Office 365 E3_USGOV_GCCHIGH"
    "1392051d-0cb9-4b7a-88d5-621fee5e8711" = "OFFICE 365 E4"
    "c7df2760-2c81-4ef7-b578-5b5392b571df" = "OFFICE 365 E5"
    "26d45bd9-adf1-46cd-a9e1-51e9a5524128" = "OFFICE 365 E5 WITHOUT AUDIO CONFERENCING"
    "4b585984-651b-448a-9e53-3b10f069cf7f" = "OFFICE 365 F3"
    "535a3a29-c5f0-42fe-8215-d3b9e1f38c4a" = "OFFICE 365 G3 GCC"
    "04a7fb0d-32e0-4241-b4f5-3f7618cd1162" = "OFFICE 365 MIDSIZE BUSINESS"
    "bd09678e-b83c-4d3f-aaba-3dad4abd128b" = "OFFICE 365 SMALL BUSINESS"
    "fc14ec4a-4169-49a4-a51e-2c852931814b" = "OFFICE 365 SMALL BUSINESS PREMIUM"
    "e6778190-713e-4e4f-9119-8b8238de25df" = "ONEDRIVE FOR BUSINESS (PLAN 1)"
    "ed01faf2-1d88-4947-ae91-45ca18703a96" = "ONEDRIVE FOR BUSINESS (PLAN 2)"
    "87bbbc60-4754-4998-8c88-227dca264858" = "POWERAPPS AND LOGIC FLOWS"
    "a403ebcc-fae0-4ca2-8c8c-7a907fd6c235" = "POWER BI (FREE)"
    "45bc2c81-6072-436a-9b0b-3b12eefbc402" = "POWER BI FOR OFFICE 365 ADD-ON"
    "f8a1db68-be16-40ed-86d5-cb42ce701560" = "POWER BI PRO"
    "a10d5e58-74da-4312-95c8-76be4e5b75a0" = "PROJECT FOR OFFICE 365"
    "776df282-9fc0-4862-99e2-70e561b9909e" = "PROJECT ONLINE ESSENTIALS"
    "09015f9f-377f-4538-bbb5-f75ceb09358a" = "PROJECT ONLINE PREMIUM"
    "2db84718-652c-47a7-860c-f10d8abbdae3" = "PROJECT ONLINE PREMIUM WITHOUT PROJECT CLIENT"
    "53818b1b-4a27-454b-8896-0dba576410e6" = "PROJECT ONLINE PROFESSIONAL"
    "f82a60b8-1ee3-4cfb-a4fe-1c6a53c2656c" = "PROJECT ONLINE WITH PROJECT FOR OFFICE 365"
    "beb6439c-caad-48d3-bf46-0c82871e12be" = "PROJECT PLAN 1"
    "1fc08a02-8b3d-43b9-831e-f76859e04e1a" = "SHAREPOINT ONLINE (PLAN 1)"
    "a9732ec9-17d9-494c-a51c-d6b45b384dcb" = "SHAREPOINT ONLINE (PLAN 2)"
    "b8b749f8-a4ef-4887-9539-c95b1eaa5db7" = "SKYPE FOR BUSINESS ONLINE (PLAN 1)"
    "d42c793f-6c78-4f43-92ca-e8f6a02b035f" = "SKYPE FOR BUSINESS ONLINE (PLAN 2)"
    "d3b4fe1f-9992-4930-8acb-ca6ec609365e" = "SKYPE FOR BUSINESS PSTN DOMESTIC AND INTERNATIONAL CALLING"
    "0dab259f-bf13-4952-b7f8-7db8f131b28d" = "SKYPE FOR BUSINESS PSTN DOMESTIC CALLING"
    "54a152dc-90de-4996-93d2-bc47e670fc06" = "SKYPE FOR BUSINESS PSTN DOMESTIC CALLING (120 Minutes)"
    "4016f256-b063-4864-816e-d818aad600c9" = "TOPIC EXPERIENCES"
    "de3312e1-c7b0-46e6-a7c3-a515ff90bc86" = "TELSTRA CALLING FOR O365"
    "4b244418-9658-4451-a2b8-b5e2b364e9bd" = "VISIO ONLINE PLAN 1"
    "c5928f49-12ba-48f7-ada3-0d743a3601d5" = "VISIO ONLINE PLAN 2"
    "4ae99959-6b0f-43b0-b1ce-68146001bdba" = "VISIO PLAN 2 FOR GCC"
    "cb10e6cd-9da4-4992-867b-67546b1db821" = "WINDOWS 10 ENTERPRISE E3"
    "6a0f6da5-0b87-4190-a6ae-9bb5a2b9546a" = "WINDOWS 10 ENTERPRISE E3"
    "488ba24a-39a9-4473-8ee5-19291e71b002" = "Windows 10 Enterprise E5"
    "6470687e-a428-4b7a-bef2-8a291ad947c9" = "WINDOWS STORE FOR BUSINESS"
}
#Usage Location Lookup Table
$UsageLocations=@{
    "United States" = "US"
    "United Kingdom" = "UK"
}

#Function to Check If Mailbox Exists Before Touching It
function MailboxExistCheck {
    Param(
            [string]$user
        )
    Clear-Variable MailboxExistsCheck -ErrorAction SilentlyContinue
    #Start Mailbox Check Wait Loop
    while ($MailboxExistsCheck -ne "YES") {
        try {
            Get-Mailbox $user -ErrorAction Stop
            $MailboxExistsCheck = "YES"
        }
        catch {
            $null = [System.Windows.MessageBox]::Show("Mailbox Does Not Exist, Waiting 60 Seconds and Trying Again")
            Start-Sleep -Seconds 60
            $MailboxExistsCheck = "NO"
        }
    }#End Mailbox Check Wait Loop    
}

#Verification Check to enable OK button on Create User Page
function CheckAllBoxes{
    if ( $CreatePasswordTextbox.Text.Length -and ($domainComboBox.SelectedIndex -ge 0) -and $usernameTextbox.Text.Length -and $firstnameTextbox.Text.Length -and $lastnameTextbox.Text.Length )
    {
        $CreateGoButton.IsEnabled = $true
    }
    else {
        $CreateGoButton.IsEnabled = $false
    }
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

$UserTextBox.Add_TextChanged({
    if (($PasswordTextBox.Text.Length -ge 8) -and ($UserTextBox.Text.Length -ge 2)){
        $PasswordGoButton.IsEnabled = $true
    }
    else{
        $PasswordGoButton.IsEnabled = $false
    }
})

$PasswordReconnectButton.Add_Click({
    Disconnect-ExchangeOnline -Confirm:$false -InformationAction Ignore -ErrorAction SilentlyContinue
    Try{
        Disconnect-SPOOnline -ErrorAction SilentlyContinue
    }
    Catch{
        #Do Nothing If Not Connected to SPO, Not Needed For Password Reset
    }
    Connect-AzureAD
    Set-Comboboxes
})

$UserButton.Add_Click({
    $tempuser = Get-AzureADUser -all $true | Out-GridView -Outputmode Single
    $UserTextBox.Text = $tempuser.UserPrincipalName
})

$PasswordGoButton.Add_Click({
    $securepassword = ConvertTo-SecureString -String $PasswordTextBox.Text -AsPlainText -Force
    Try{
        if($PasswordResetCheckbox.IsChecked -eq $true){
            Set-AzureADUserPassword -ObjectID $UserTextBox.Text -Password $securepassword -ForceChangePasswordNextLogin $true -ErrorAction Stop
        }
        else{
            Set-AzureADUserPassword -ObjectID $UserTextBox.Text -Password $securepassword -ForceChangePasswordNextLogin $false -ErrorAction Stop
        }

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

### Start Mailbox Tab Functionality
$MailboxReconnectButton.Add_Click({
    Try{
        Disconnect-SPOOnline -ErrorAction SilentlyContinue
    }
    Catch{
        #Do Nothing If Not Connected to SPO, Not Needed For Mailboxes
    }
    Connect-AzureAD
    Connect-ExchangeOnline -ShowBanner:$false
    Set-Comboboxes
})

$MailboxAddButton.Add_Click({
    Try {
        Get-Mailbox -ErrorAction SilentlyContinue | Out-Null
    }
    Catch {
        Connect-ExchangeOnline -ShowBanner:$false
    }
    Clear-Variable Mailboxes -ErrorAction SilentlyContinue
    Clear-Variable SharedMailboxUsers -ErrorAction SilentlyContinue
    $SharedMailboxUsers = Get-AzureADUser -All $true | Where-Object {$_.AccountEnabled } |  Select-Object Displayname, UserPrincipalName | Sort-Object Displayname | Out-GridView -Title "Select User to Grant Access - Hold Ctrl for Multiple" -Passthru
    if($SharedMailboxUsers){
        $Mailboxes = Get-Mailbox -ResultSize Unlimited | Select-Object DisplayName,UserPrincipalName | Out-GridView -Title "Select Mailbox to Share - Hold Ctrl for Multiple" -Passthru
        if($Mailboxes){
            foreach($Mailbox in $Mailboxes){
                foreach($SharedMailboxUser in $SharedMailboxUsers){
                    Add-MailboxPermission -Identity $Mailbox.UserPrincipalName -User $SharedMailboxUser.UserPrincipalName -AccessRights FullAccess -InheritanceType All | Out-Null
                    Add-RecipientPermission -Identity $Mailbox.UserPrincipalName -Trustee $SharedMailboxUser.UserPrincipalName -AccessRights SendAs -Confirm:$False | Out-Null
                    Write-MailboxRichTextBox("$($SharedMailboxUser.Displayname) has been added to $($Mailbox.DisplayName)'s Mailbox`r")
                }
            }
        }else{
            Write-MailboxRichTextBox("No Mailbox Selected`r") -Color "Red"    
        }
    }else{
        Write-MailboxRichTextBox("No User Selected`r") -Color "Red"
    }
})

$MailboxRemoveButton.Add_Click({

})
### End Mailbox Tab Functionality

### Start Group Tab Functionality
$GroupReconnectButton.Add_Click({
    Disconnect-ExchangeOnline -Confirm:$false -InformationAction Ignore -ErrorAction SilentlyContinue
    Try{
        Disconnect-SPOOnline -ErrorAction SilentlyContinue
    }
    Catch{
        #Do Nothing If Not Connected to SPO, Not Needed For Groups
    }
    Connect-AzureAD
    Set-Comboboxes
})

$GroupAddButton.Add_Click({
    Try{
        Get-AzureADUser -ErrorAction SilentlyContinue | Out-Null
    }
    Catch{
        Connect-AzureAD
    }
    # Pull User ObjectID and Group ObjectID to add member to all groups selected, skipping dynamic
    Clear-Variable Users -ErrorAction SilentlyContinue
    Clear-Variable roups -ErrorAction SilentlyContinue
    
    $Users = Get-AzureADUser -All $true | Out-GridView -Title "Select User - Hold Ctrl for Multiple" -PassThru
    if($Users){
        foreach($User in $Users){
            $Groups = Get-AzureADMSGroup -All $true | Where-Object {$_.GroupTypes -notcontains "DynamicMembership"} | Select-Object DisplayName,Description,Id | Sort-Object DisplayName | Out-GridView -Passthru -Title "Hold Ctrl to select multiple groups"
            if ($Groups){
                foreach($Group in $Groups){
                    Add-AzureADGroupMember -ObjectId $Group.Id -RefObjectId $User.ObjectID
                    Write-GroupRichTextBox("$($User.DisplayName) added to $($Group.Displayname)")
                }
            }else{
                Write-GroupRichTextBox("Group Selection Cancelled`r") -Color "Red"
            }
        }
    }else{
        Write-GroupRichTextBox("User Selection Cancelled`r") -Color "Red"
    }
})

$GroupRemoveButton.Add_Click({
    Try{
        Get-AzureADUser -ErrorAction SilentlyContinue | Out-Null
    }
    Catch{
        Connect-AzureAD
    }
    # Pull User ObjectID and Group ObjectID to remove member to all groups selected, skipping dynamic
    Clear-Variable users -ErrorAction SilentlyContinue
    Clear-Variable groups -ErrorAction SilentlyContinue
    
    $Users = Get-AzureADUser -All $true | Out-GridView -Title "Select User - Hold Ctrl for Multiple" -PassThru
    foreach($User in $Users){
        $Groups = Get-AzureADUserMembership -ObjectId $user.ObjectId | Where-Object {($_.ObjectType -ne "Role") -and ($_.GroupTypes -notcontains "DynamicMembership")} | Select-Object DisplayName,ObjectId | Sort-Object Displayname | Out-GridView -Title "Select Group - Hold Ctrl for Multiple" -PassThru
        foreach ($Group in $Groups) { 
            Try
            {
                Remove-AzureADGroupMember -ObjectId $Group.ObjectId -MemberId $User.ObjectId
                Write-GroupRichTextBox ("Removed $($user.Displayname) from $($Group.Displayname)`r")
            }
            catch
            {
                Write-GroupRichTextBox ("Could not remove from group $($Group.Displayname).  Error:  $_.Message") -Color "Red"
            }
        }
    }
})
### End Group Tab Functionality

### Start Calendar Tab Functionality
$CalendarReconnectButton.Add_Click({
    Try{
        Disconnect-SPOOnline -ErrorAction SilentlyContinue
    }
    Catch{
        #Do Nothing If Not Connected to SPO, Not Needed For Creation
    }
    Connect-AzureAD
    Connect-ExchangeOnline -ShowBanner:$false
    Set-Comboboxes
})

$CalendarGoButton.Add_Click({

})
### End Calendar Tab Functionality



### Start User Creation Tab Functionality
$firstnameTextbox.Add_TextChanged({
    CheckAllBoxes
})

$lastnameTextbox.Add_TextChanged({
    CheckAllBoxes
})

$usernameTextbox.Add_TextChanged({
    CheckAllBoxes
})

$CreatePasswordTextbox.Add_TextChanged({
    CheckAllBoxes
})

$DomainCombobox.Add_SelectionChanged({
    CheckAllBoxes
})

$UsageLocationCombobox.Add_SelectionChanged({
    CheckAllBoxes
})

$CreateReconnectButton.Add_Click({
    Try{
        Disconnect-SPOOnline -ErrorAction SilentlyContinue
    }
    Catch{
        #Do Nothing If Not Connected to SPO, Not Needed For Creation
    }
    Connect-AzureAD
    Connect-ExchangeOnline -ShowBanner:$false
    Set-Comboboxes
})

$CreateGoButton.Add_Click({
    Clear-Variable AvailableLicenseCheck -ErrorAction SilentlyContinue
    Try{
        Get-AzureADUser -ErrorAction SilentlyContinue | Out-Null
    }
    Catch{
        Connect-AzureAD
    }
    Try{
        Get-Mailbox -ErrorAction SilentlyContinue | Out-Null
    }
    Catch{
        Connect-ExchangeOnline -ShowBanner:$false
    }
    $Licenses =  Get-AzureADSubscribedSku | Select-Object -Property Sku*,ConsumedUnits -ExpandProperty PrepaidUnits
    foreach($License in $Licenses){
        $TempSkuCheck = $skuToFriendly.Item("$($License.SkuID)")
        if($TempSkuCheck){
            $License.SkuPartNumber = $skuToFriendly.Item("$($License.SkuID)")
        }
        else{
            $null = [System.Windows.MessageBox]::Show("Please Submit a Github Issue for Non-Matching SkuPartNumber $($License.SkuID) - $($License.SkuPartNumber): https://github.com/mrobinson-ws/touch_azure_users/issues")
        }
    }
    while($AvailableLicenseCheck -ne $true){
        Clear-Variable SelectedLicenses -ErrorAction SilentlyContinue
        $SelectedLicenses = $Licenses | Sort-Object SkuPartNumber | Out-GridView -Passthru -Title "Hold Ctrl For Multiple Licenses"
        foreach($SelectedLicense in $SelectedLicenses){
            if($SelectedLicense.Enabled-$SelectedLicense.ConsumedUnits -ge 1){
                $Available = $SelectedLicense.Enabled-$SelectedLicense.ConsumedUnits
                $null = [System.Windows.MessageBox]::Show("You have $Available available $($SelectedLicense.SkuPartNumber) licenses")
                $AvailableLicenseCheck = $true
            }
            elseif($SelectedLicense.Enabled-$SelectedLicense.ConsumedUnits -le 0){
                $null = [System.Windows.MessageBox]::Show("You do not have any $($SelectedLicense.SkuPartNumber) licenses to assign, please acquire licenses and try again","License Check","OKCancel","Warning")
                $AvailableLicenseCheck = $false
            }
        }
        if(-not($SelectedLicenses)){
            Break
        }
    }

    if ($AvailableLicenseCheck -eq $true) {
        $PasswordProfile = New-Object -TypeName Microsoft.Open.AzureAD.Model.PasswordProfile
        $PasswordProfile.Password = $CreatePasswordTextbox.text
        $UPN = $usernameTextbox.Text + "@" + $domainCombobox.Text
        $displayname = $firstnameTextbox.text + " " + $lastnameTextbox.Text
        $usageloc = $UsageLocations[$UsageLocationComboBox.Text]
               
        #Test if User with matching UPN already exists, then exits
        try {
            Write-CreateRichTextBox("Testing If Username Does Not Exist`r")
            Get-AzureAdUSer -ObjectID $UPN -ErrorAction Stop | Out-Null
            $UserExists = $true
        }
        #Otherwise, creates user and assigns licenses selected in first step
        catch {
            Write-CreateRichTextBox("Username Does Not Exist, Creating User and Assigning Licenses`r")
            $UserExists = $false
        }

        if($UserExists -ne $true){
            New-AzureADUser -DisplayName $displayname -PasswordProfile $PasswordProfile -UserPrincipalName $UPN -AccountEnabled $true -MailNickName "$($usernameTextbox.Text)" -UsageLocation $usageloc
            foreach($SelectedLicense in $SelectedLicenses){
                $AssignedLicense = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicense
                $AssignedLicense.SkuID = $SelectedLicense.SkuID
                $Licenses = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicenses 
                $Licenses.AddLicenses = $AssignedLicense
                Set-AzureADUserLicense -ObjectID $UPN -AssignedLicenses $Licenses
            }
            if([string]::IsNullOrwhiteSpace($cityTextbox.Text) -eq $false){
                Set-AzureADUser -ObjectId $UPN -City $cityTextbox.Text
            }
            if([string]::IsNullOrWhiteSpace($stateTextbox.Text) -eq $false){
                Set-AzureADUser -ObjectID $UPN -State $stateTextbox.Text
            }
            if([string]::IsNullOrWhiteSpace($countryTextbox.Text) -eq $false){
                Set-AzureADUser -ObjectID $UPN -State $countryTextbox.Text
            }
            if([string]::IsNullOrWhiteSpace($CustomAttribute1Textbox.Text) -eq $false){
                MailboxExistCheck($UPN)
                Set-Mailbox $UPN -CustomAttribute1 $CustomAttribute1Textbox.Text
            }
        
            $user = Get-AzureADUser -ObjectID $UPN
            MailboxExistCheck($UPN)
            $Groups = Get-AzureADMSGroup -All $true | Where-Object {$_.GroupTypes -notcontains "DynamicMembership"} | Select-Object DisplayName,Description,Id | Sort-Object DisplayName | Out-GridView -Passthru -Title "Hold Ctrl to select multiple groups" | Select-Object -ExpandProperty Id
            if ($Groups){
                foreach($group in $Groups){
                    Add-AzureADGroupMember -ObjectId $group -RefObjectId $user.ObjectID
                    Write-CreateRichTextBox("Added $($user.DisplayName) to $($group.Displayname)`r")
                }
                
            }
            else {
                Write-CreateRichTextBox("No Groups Selected`r") -Color "Yellow"
            }
            $CustomAttribute1Textbox.Text = ""
            $cityTextbox.Text = ""
            $stateTextbox.Text = ""
            $countryTextBox.Text = ""
            $firstnameTextbox.Text = ""
            $lastnameTextbox.Text = ""
            $usernameTextbox.Text = ""
            $CreatePasswordTextbox.Text = ""
        }
        Else{
            $ExistingUser = Get-AzureADUser -ObjectID $UPN
            Write-CreateRichTextBox("Username $UPN Exists as $($ExistingUser.DisplayName), Please Review and Try Again`r") -Color "Red"
        }
    }
Clear-Variable Licenses -ErrorAction SilentlyContinue
Clear-Variable SelectedLicenses -ErrorAction SilentlyContinue
Clear-Variable AvailableLicenseCheck -ErrorAction SilentlyContinue
Clear-Variable Groups -ErrorAction SilentlyContinue
Clear-Variable UserExists -ErrorAction SilentlyContinue
})
### End User Creation Tab Functionality

### Start User Termination Tab Functionality
$TerminateReconnectButton.Add_Click({
    Connect-AzureAD
    Connect-ExchangeOnline -ShowBanner:$false
    $domainPrefix = ((Get-AzureADDomain | Where-Object Name -match "\.onmicrosoft\.com")[0].Name -split '\.')[0]
    $AdminSiteUrl = "https://$domainPrefix-admin.sharepoint.com"
    Connect-SPOService -Url $AdminSiteURL
    Set-Comboboxes
})

$RemoveGoButton.Add_Click({
    Try{
        Get-AzureADUser -ErrorAction Stop | Out-Null
    }Catch{
        Connect-AzureAD
    }
    Try{
        Get-Mailbox -ErrorAction Stop | Out-Null
    }Catch{
        Connect-ExchangeOnline -ShowBanner:$false
    }
    
    #Pull All Azure AD Users and Store In Hash Table Instead Of Calling Get-AzureADUser Multiple Times
    $allUsers = @{}    
    foreach ($user in Get-AzureADUser -All $true){ $allUsers[$user.UserPrincipalName] = $user }
    
    #Request Username(s) To Be Terminated From Script Runner (Hold Ctrl To Select Multiples)
    $username = $allUsers.Values | Where-Object {$_.AccountEnabled } | Sort-Object DisplayName | Select-Object -Property DisplayName,UserPrincipalName | Out-Gridview -OutputMode Single -Title "Please select the user(s) to be terminated" | Select-Object -ExpandProperty UserPrincipalName
    
    ##### Start User(s) Loop #####
    if($username){
        $UserInfo = $allusers[$username]

        #Block Sign In Of User/Force Sign Out Within 60 Minutes
        Set-AzureADUser -ObjectID $UserInfo.ObjectId -AccountEnabled $false
        Write-RemoveRichTextBox("Sign in Blocked for $($UserInfo.ObjectID)`r")

        #Remove All Group Memberships
        $memberships = Get-AzureADUserMembership -ObjectId $username | Where-Object {$_.ObjectType -ne "Role"}| Select-Object DisplayName,ObjectId
        foreach ($membership in $memberships) { 
                $group = Get-AzureADMSGroup -ID $membership.ObjectId
                if ($group.GroupTypes -contains 'DynamicMembership') {
                    Write-RemoveRichTextBox("Skipped $($group.Displayname) as it is dynamic`r") -Color "Yellow"
                }
                else{
                    Try{
                        Remove-AzureADGroupMember -ObjectId $membership.ObjectId -MemberId $UserInfo.ObjectId -ErrorAction Stop
                    }Catch{
                        Write-RemoveRichTextBox("Could not remove from group $($group.name).  Error:  $_.Message`r") -color "Yellow"
                    }
                }
            }
        Write-RemoveRichTextBox("All non-dynamic groups removed, please check your Downloads folder for the file, it will also open automatically at end of user termination`r")

        #Convert To Shared Mailbox And Hide From GAL When Convert Is Selected, Must Be Done Before Removing Licenses
        if ($ConvertCheckBox.IsChecked -eq $true) {
            Set-Mailbox $username -Type Shared -HiddenFromAddressListsEnabled $true
            Write-RemoveRichTextBox("Mailbox for $username converted to Shared, address hidden from GAL`r")
        }

        #Grant Access To Shared Mailbox When Grant CheckBox Is Selected
        if ($ShareMailboxCheckBox.IsChecked -eq $true) {
            $sharedMailboxUser = $allUsers.Values | Where-Object {$_.AccountEnabled } | Sort-Object DisplayName | Select-Object -Property DisplayName,UserPrincipalName | Out-GridView -Title "Please select the user(s) to share the $username Shared Mailbox with" -OutputMode Single | Select-Object -ExpandProperty UserPrincipalName
            if($sharedMailboxUser){
               Add-MailboxPermission -Identity $username -User $SharedMailboxUser -AccessRights FullAccess -InheritanceType All
                Add-RecipientPermission -Identity $username -Trustee $SharedMailboxUser -AccessRights SendAs -Confirm:$False
                Write-RemoveRichTextBox("Access granted to the $username Shared Mailbox to $sharedMailboxUser`r")
            }
            else{
                Write-RemoveRichTextBox("Cancelled Sharing of Mailbox`r") -Color "Red"
            }
        }

        #Remove All Licenses When Remove Licenses Is Selected
        if ($RemoveLicensesCheckBox.IsChecked -eq $true) {
            $licenses = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicenses
            if($UserInfo.assignedlicenses){
                $licenses.RemoveLicenses = $UserInfo.assignedlicenses.SkuId
                Set-AzureADUserLicense -ObjectId $UserInfo.ObjectId -AssignedLicenses $licenses
            }
            Write-RemoveRichTextBox("All licenses have been removed`r")
        }

        #Test And Connect To Sharepoint Online If Needed
        if ($OneDriveNoRadioButton.IsChecked -ne $true) {
            $domainPrefix = ((Get-AzureADDomain | Where-Object Name -match "\.onmicrosoft\.com")[0].Name -split '\.')[0]
            $AdminSiteUrl = "https://$domainPrefix-admin.sharepoint.com"
            Try{
                Get-SPOSite -ErrorAction Stop | Out-Null
            }Catch{
                Connect-SPOService -Url $AdminSiteURL
            }
        }

        #Share OneDrive With Same User as Shared Mailbox
        if ($OneDriveSameRadioButton.IsChecked -eq $true) {
            #Pull OneDriveSiteURL Dynamically And Grant Access
            $OneDriveSiteURL = Get-SPOSite -Filter "Owner -eq $($UserInfo.UserPrincipalName)" -IncludePersonalSite $true | Select-Object -ExpandProperty Url            

            #Add User Receiving Access To Terminated User's OneDrive, Add The Access Link To CSV File For Copying
            Set-SPOUser -Site $OneDriveSiteUrl -LoginName $SharedMailboxUser -IsSiteCollectionAdmin $True
            Write-RemoveRichTextBox("OneDrive Data Shared with $SharedMailboxUser successfully, link to copy and give to Manager is $OneDriveSiteURL`r")
        }
        #Share OneDrive With Different User(s) than Shared Mailbox
        elseif ($OneDriveDifferentRadioButton.IsChecked -eq $true) {
            $SharedOneDriveUser = $allusers.Values | Sort-Object Displayname | Select-Object -Property DisplayName,UserPrincipalName | Out-GridView -Title "Please select the user to share the OneDrive with" -OutputMode Single | Select-Object -ExpandProperty UserPrincipalName
            
            #Pull Object ID Needed For User Receiving Access To OneDrive And OneDriveSiteURL Dynamically
            $OneDriveSiteURL = Get-SPOSite -Filter "Owner -eq $($UserInfo.UserPrincipalName)" -IncludePersonalSite $true | Select-Object -ExpandProperty Url            

            #Add User Receiving Access To Terminated User's OneDrive, Add The Access Link To CSV File For Copying
            Set-SPOUser -Site $OneDriveSiteUrl -LoginName $SharedOneDriveUser -IsSiteCollectionAdmin $True
            Write-RemoveRichTextBox("OneDrive Data Shared with $SharedOneDriveUser successfully, link to copy and provide to trustee is $OneDriveSiteURL`r")
        }

        #Export Groups Removed and OneDrive URL to CSV
        [pscustomobject]@{
            GroupsRemoved    = $memberships.DisplayName -join ','
            OneDriveSiteURL = $OneDriveSiteURL
        } | Export-Csv -Path c:\users\$env:USERNAME\Downloads\$(get-date -f yyyy-MM-dd)_info_on_$username.csv -NoTypeInformation

        #Open Created CSV File At End Of Loop For Ease Of Copying OneDrive URL To Give
        Start-Process c:\users\$env:USERNAME\Downloads\$(get-date -f yyyy-MM-dd)_info_on_$username.csv
    }
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
    
    Set-Comboboxes
})

$null = $UserForm.ShowDialog()
