param
(
  [String]
  [Parameter(Mandatory)]
  $ExchServer
)


$null = [reflection.assembly]::LoadWithPartialName('System.Windows.Forms')


$Search_Name = Get-Date -Format 'dd/MM/yyyy hh:mm:ss tt'

$QueryBuilderForm = New-Object -TypeName System.Windows.Forms.Form
$QueryBuilderForm.Size = '400,400'
$QueryBuilderForm.Text = 'Spam Email - Search and Removal Tool'
$QueryBuilderForm.TopMost = $true


$DateLabel = New-Object -TypeName System.Windows.Forms.Label
$DateLabel.Text = 'Days to search'
$DateLabel.Location = '20,20'
$QueryBuilderForm.Controls.Add($DateLabel)


$DateTextBox = New-Object -TypeName System.Windows.Forms.TextBox
$DateTextBox.Location = '125,20'
$DateTextBox.Size = '200,23'
$QueryBuilderForm.Controls.Add($DateTextBox)

$SenderLabel = New-Object -TypeName System.Windows.Forms.Label
$SenderLabel.Text = 'Senders Address'
$SenderLabel.Location = '20,60'
$QueryBuilderForm.Controls.Add($SenderLabel)


$SenderTextBox = New-Object -TypeName System.Windows.Forms.TextBox
$SenderTextBox.Location = '125,60'
$SenderTextBox.Size = '200,23'
$QueryBuilderForm.Controls.Add($SenderTextBox)

$SubjectLabel = New-Object -TypeName System.Windows.Forms.Label
$SubjectLabel.Text = 'Email Subject'
$SubjectLabel.Location = '20,100'
$QueryBuilderForm.Controls.Add($SubjectLabel)


$SubjectTextBox = New-Object -TypeName System.Windows.Forms.TextBox
$SubjectTextBox.Location = '125,100'
$SubjectTextBox.Size = '200,23'
$QueryBuilderForm.Controls.Add($SubjectTextBox)

$PhraseLabel = New-Object -TypeName System.Windows.Forms.Label
$PhraseLabel.Text = 'Search Phrase'
$PhraseLabel.Location = '20,140'
$QueryBuilderForm.Controls.Add($PhraseLabel)


$PhraseTextBox = New-Object -TypeName System.Windows.Forms.TextBox
$PhraseTextBox.Location = '125,140'
$PhraseTextBox.Size = '200,23'
$QueryBuilderForm.Controls.Add($PhraseTextBox)

$UsageLabel1 = New-Object -TypeName System.Windows.Forms.Label
$UsageLabel1.Location = '20,175'
$UsageLabel1.Font = [System.Drawing.Font]::new("Microsoft Sans Serif", 8, [System.Drawing.FontStyle]::Bold)
$UsageLabel1.Text = 'User Guide:'
$QueryBuilderForm.Controls.Add($UsageLabel1)

$UsageLabel2 = New-Object -TypeName System.Windows.Forms.Label
$UsageLabel2.Location = '20,205'
$UsageLabel2.Size = '220,80'
$UsageLabel2.Text = 'Method 1.'+"`n"+'Complete any combination of the first three fields.'+"`n`n"+'Method 2.'+"`n"+'Enter a search phrase into the fourth field'
$QueryBuilderForm.Controls.Add($UsageLabel2)

$SearchButton = New-Object -TypeName System.Windows.Forms.Button
$SearchButton.Text = 'Search'
$SearchButton.Location = '100,300'
$SearchButton.add_click({
    Try
    {
      Get-MailboxDatabase | Out-Null
    }
    Catch
    {
      $Acc = Get-Credential
      $Sess = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$ExchServer/Powershell/ -Credential $Acc -Authentication Kerberos
      Import-PSSession -Session $Sess -DisableNameChecking
    }
  
    #region Query Builder
    if(($DateTextBox.Text).Length -gt 0)
    {
      $Query = "sent>={0:MM/dd/yyyy} AND`n" -f (Get-Date).AddDays(-$DateTextBox.Text)+(Get-Date -Format "'sent'<=MM/dd/yyyy")
    }
  
    if($Query -eq $null)
    {
      if(($SenderTextBox.Text).Length -gt 0)
      {
        $Query = ('from:"{0}"' -f $SenderTextBox.Text)
      }
    }
    else
    {
      if(($SenderTextBox.Text).Length -gt 0)
      {
        $Query = $Query +' AND' +"`n" + ('from:"{0}"' -f $SenderTextBox.Text)
      }
    }
  
    if($Query -eq $null)
    {
      if(($SubjectTextBox.Text).Length -gt 0)
      {
        $Query = ('subject:"{0}"' -f $SubjectTextBox.Text)
      }
    }
    else
    {
      if(($SubjectTextBox.Text).Length -gt 0)
      {
        $Query = $Query +' AND' +"`n" + ('subject:"{0}"' -f $SubjectTextBox.Text)
      }
    }
    
    if($Query -eq $null)
    {
      if(($PhraseTextBox.Text).Length -gt 0)
      {
        $Query = ('"{0}"' -f $PhraseTextBox.Text)
      }
    }
    else
    {
      if(($PhraseTextBox.Text).Length -gt 0)
      {
        $Query = $Query +' AND' +"`n" + ('{0}' -f $PhraseTextBox.Text)
      }
    }
  
    Write-Host 'Search Query'
    Write-Host ''
    Write-Host $Query
    #endregion
    New-ComplianceSearch -Name $Search_Name -ExchangeLocation All -ContentMatchQuery $Query
    Start-ComplianceSearch -Identity $Search_Name
    do
    {
      Start-Sleep -Seconds 5
    }until((Get-ComplianceSearch -Identity $Search_Name).Status -eq 'Completed')

    $QueryBuilderForm.Close()
    $QueryBuilderForm.DialogResult = 'Ok'
})
$QueryBuilderForm.Controls.Add($SearchButton)

$CancelButton = New-Object -TypeName System.Windows.Forms.Button
$CancelButton.Text = 'Cancel'
$CancelButton.Location = '210,300'
$QueryBuilderForm.CancelButton = $CancelButton
$QueryBuilderForm.Controls.Add($CancelButton)


$QueryBuilderForm.ShowDialog()



if($QueryBuilderForm.DialogResult -eq 'Cancel'){
    exit
}


if((Get-ComplianceSearch -Identity $Search_Name).Items -gt 0){
do
{
  $Message = (New-ComplianceSearchAction -SearchName $Search_Name -Preview | Select-Object -ExpandProperty Results) -replace ';', "`n" -replace ',', "`n`n"
}
Until($Message -ne '{}')
}else{
    [System.Windows.MessageBox]::Show('The search had no results')
    Remove-ComplianceSearch -Identity $Search_Name -Confirm:$false
    break
}


$Split = (($Message -Replace 'Location:', 'LocationLocationLocation:') -split 'LocationLocation') -replace '{', '' -replace '}', ''
$Values = New-Object -TypeName System.Collections.ArrayList

foreach ($item in $Split)
{
  $lines = $item -split "`n"
  $Values += New-Object -TypeName psobject -Property @{
    Recipient = ($lines | Select-String -Pattern 'Location:') -replace 'Location: ', ''
    Sender    = ($lines | Select-String -Pattern 'Sender:') -replace 'Sender: ', ''
    Subject   = ($lines | Select-String -Pattern 'Subject:') -replace 'Subject: ', ''
    Received  = ($lines | Select-String -Pattern 'Received Time:') -replace 'Received Time: ', ''
    Link      = ($lines | Select-String -Pattern 'Data Link:') -replace 'Data Link: ', ''
  }
}
$Values = $Values | Where-Object -FilterScript {
  $_.Sender -ne '{}'
}
$Values |
Select-Object -Property Sender, Recipient, Subject, Received, Link |
Out-GridView
    
$ConfirmationForm = New-Object -TypeName System.Windows.Forms.Form
$ConfirmationForm.Text = 'Remove'
$ConfirmationForm.Size = '300,200'
$RemoveMessageLabel = New-Object -TypeName System.Windows.Forms.Label
$RemoveMessageLabel.Text = 'Would you like to remove the emails found in this search?'
$RemoveMessageLabel.TextAlign = 'TopCenter'
$RemoveMessageLabel.Size = '150,50'
$RemoveMessageLabel.Location = '65,10'
$ConfirmButton = New-Object -TypeName System.Windows.Forms.Button
$ConfirmButton.Text = 'Remove'
$ConfirmButton.Location = '60,50'
$ConfirmButton.Add_Click({
    New-ComplianceSearchAction -SearchName $Search_Name -Purge -PurgeType SoftDelete -Confirm:$false
    $SearchActionName = $Search_Name+'_Purge'
    do
    {
      Start-Sleep -Seconds 5
    }until((Get-ComplianceSearchAction -Identity $SearchActionName).Status -eq 'Completed')    
    
    Remove-ComplianceSearch -Identity $Search_Name -Confirm:$false
    $ConfirmationForm.Close()
    $ConfirmationForm.DialogResult = 'Ok'

})
$ConfirmationForm.Controls.Add($ConfirmButton)
$CancelButton2 = New-Object -TypeName System.Windows.Forms.Button
$CancelButton2.Text = 'Cancel'
$CancelButton2.Location = '150,50'
$CancelButton2.add_click({
    Remove-ComplianceSearch -Identity $Search_Name -Confirm:$false
    $ConfirmationForm.Close()
})
$ConfirmationForm.CancelButton = $CancelButton2
$ConfirmationForm.Controls.Add($CancelButton2)
$ConfirmationForm.Controls.Add($RemoveMessageLabel)
$ConfirmationForm.TopMost = $true

$WarningLabel1 = New-Object -TypeName System.Windows.Forms.Label
$WarningLabel1.Font = [System.Drawing.Font]::new("Microsoft Sans Serif", 8, [System.Drawing.FontStyle]::Bold)
$WarningLabel1.ForeColor = 'RED'
$WarningLabel1.Text = 'WARNING'
$WarningLabel1.Location = '110,80'
$WarningLabel1.Size = '75,18'
$ConfirmationForm.Controls.Add($WarningLabel1)

$WarningLabel2 = New-Object -TypeName System.Windows.Forms.Label
$WarningLabel2.Text = 'Pressing the Remove button above will remove all emails that are shown on the list'
$WarningLabel2.Location = '40,100'
$WarningLabel2.Size = '200,100'
$WarningLabel2.TextAlign = 'TopCenter'
$ConfirmationForm.Controls.Add($WarningLabel2)

$ConfirmationForm.ShowDialog()
