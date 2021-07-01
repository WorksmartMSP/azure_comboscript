Add-Type -AssemblyName PresentationFramework

if(-not(Get-Module ExchangeOnlineManagement -ListAvailable)){
    $null = [System.Windows.MessageBox]::Show('Please Install ExchangeOnlineManagement - view ITG for details https://worksmart.itglue.com/2503920/docs/7777752')
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
                    <Label Content="Please Pick A User, Then Enter A Password, Go Button Activates Once Selected" HorizontalAlignment="Left" Margin="10,10,0,0" VerticalAlignment="Top" Height="25" Width="473"/>
                    <TextBox Name="PasswordTextBox" HorizontalAlignment="Left" Height="25" Margin="10,130,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="473"/>
                    <Button Name="PasswordGoButton" Content="Go" HorizontalAlignment="Left" Margin="10,399,0,0" VerticalAlignment="Top" Width="473" Height="50" IsEnabled="False"/>
                    <TextBox Name="UserTextBox" HorizontalAlignment="Left" Height="23" Margin="258,55,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="225" IsReadOnly="True" Background="#FFC8C8C8"/>
                    <Button Name="UserButton" Content="Pick User" HorizontalAlignment="Left" Margin="10,40,0,0" VerticalAlignment="Top" Width="243" Height="54"/>
                    <Label Content="Enter Password Below, Go Button Activates At 8 Characters" HorizontalAlignment="Left" Margin="10,99,0,0" VerticalAlignment="Top" Width="473"/>
                    <RichTextBox Name="PasswordRichTextBox" HorizontalAlignment="Left" Height="234" Margin="10,160,0,0" VerticalAlignment="Top" Width="473" Background="#FF646464" Foreground="Cyan" HorizontalScrollBarVisibility="Auto" VerticalScrollBarVisibility="Auto">
                        <FlowDocument/>
                    </RichTextBox>
                </Grid>
            </TabItem>
            <TabItem Header="Create User">
                <Grid Background="#FFE5E5E5"/>
            </TabItem>
            <TabItem Header="Terminate User" Margin="-2,-2,-2,0">
                <Grid Background="#FFE5E5E5"/>
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

$null = $UserForm.ShowDialog()