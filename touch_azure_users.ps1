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
        <TabControl HorizontalAlignment="Left" Height="487" Margin="10,0,0,0" VerticalAlignment="Top" Width="499">
            <TabItem Header="Reset Password">
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
            <TabItem Header="Create User">
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
                    <Button Content="Create User" HorizontalAlignment="Left" Margin="183,218,0,0" VerticalAlignment="Top" Width="155" Height="139" IsEnabled="False" TabIndex="12"/>
                </Grid>
            </TabItem>
            <TabItem Header="Terminate User" Margin="-2,-2,-2,0">
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
                    <Button Content="Button" HorizontalAlignment="Left" Margin="10,397,0,0" VerticalAlignment="Top" Width="473" Height="52"/>
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

# Create Functions For Color Changing Messages
Function WritePasswordRichTextBox {
    Param(
        [string]$text,
        [string]$color = "Cyan"
    )

    $RichTextRange = New-Object System.Windows.Documents.TextRange( 
        $PasswordRichTextBox.Document.ContentEnd,$PasswordRichTextBox.Document.ContentEnd ) 
    $RichTextRange.Text = $text
    $RichTextRange.ApplyPropertyValue( ( [System.Windows.Documents.TextElement]::ForegroundProperty ), $color )  

}

Try{
    Get-AzureADUser -ErrorAction Stop
}Catch{
    Connect-AzureAD
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
        WritePasswordRichTextBox("SUCCESS:  $($UserTextBox.Text)'s password has been reset to $($PasswordTextBox.Text)`r")
        $PasswordRichTextBox.ScrollToEnd()
        $UserTextbox.Text = ""
        $PasswordTextBox.Text = ""
    }Catch{
        $message = $_.Exception.Message
        if ($_.Exception.ErrorContent.Message.Value) {
            $message = $_.Exception.ErrorContent.Message.Value
        }
        WritePasswordRichTextBox("$message`rFAILURE:  Please review above and try again`r") -Color "Red"
        $PasswordRichTextBox.ScrollToEnd()
    }
})
### End Password Tab Functionality

### Start User Creation Tab Functionality



### End User Creation Tab Functionality

### Start User Termination Tab Functionality



### End User Termination Tab Functionality




$null = $UserForm.ShowDialog()