<#-----[task]_Exchange-mailbox-change-primary-email-address-----#>
# Connect to Exchange
try {
    $adminSecurePassword = ConvertTo-SecureString -String "$ExchangeAdminPassword" -AsPlainText -Force
    $adminCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $ExchangeAdminUsername, $adminSecurePassword
    $sessionOption = New-PSSessionOption -SkipCACheck -SkipCNCheck #-SkipRevocationCheck
    $exchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $exchangeConnectionUri -Credential $adminCredential -SessionOption $sessionOption -Authentication Kerberos -ErrorAction Stop #-AllowRedirection
    HID-Write-Status -Message "Successfully connected to Exchange using the URI [$exchangeConnectionUri]" -Event Success
} catch {
    HID-Write-Status -Message "Error connecting to Exchange using the URI [$exchangeConnectionUri]" -Event Error
    HID-Write-Status -Message "$($_.Exception.Message)" -Event Error
    HID-Write-Summary -Message "Failed to connect to Exchange using the URI [$exchangeConnectionUri]" -Event Failed
    throw $_
}

try {
    $newPrimaryMail = $emailAddress
    HID-Write-Status -Message "$($isPrimary.gettype())" -Event Success
    HID-Write-Status -Message "$($isPrimary)" -Event Success
    if ($isPrimary -eq "true") {
        HID-Write-Status -Message "Successfully set primary emailaddress to [$newPrimaryMail] for mailbox [$UserPrincipalName]" -Event Success
        HID-Write-Status -Message "No changes where made. The selected primary address is the same as the existing primary address" -Event Success
        HID-Write-Summary -Message "Successfully set primary emailaddress to [$newPrimaryMail] for mailbox [$UserPrincipalName]" -Event Success
        return
    }

    $ParamsGetMailbox = @{
        Identity = $UserPrincipalName
    }
    $mailBoxes = Invoke-Command -Session $exchangeSession -ScriptBlock {
        Param ($ParamsGetMailbox)
        Get-Mailbox @ParamsGetMailbox
    } -ArgumentList $ParamsGetMailbox


    $list = New-Object System.Collections.ArrayList
    foreach ($address in $mailBoxes.EmailAddresses) {
        $prefix = $address.Split(":")[0]
        $mail = $address.Split(":")[1]
        if ($mail.ToLower() -eq $newPrimaryMail.ToLower()) {
            $address = "SMTP:" + $mail
        } else {
            $address = $prefix.ToLower() + ":" + $mail
        }
        $list.Add($address)
    }

    $ParamsSetMailbxox = @{
        Identity                  = $UserPrincipalName
        EmailAddresses            = $list
        EmailAddressPolicyEnabled = $false
    }
    $null = Invoke-Command -Session $exchangeSession -ScriptBlock {
        Param ($ParamsSetMailbxox)
        Set-Mailbox @ParamsSetMailbxox
    } -ArgumentList $ParamsSetMailbxox

    HID-Write-Status -Message "Successfully set primary emailaddress to [$newPrimaryMail] for mailbox [$UserPrincipalName]" -Event Success
    HID-Write-Summary -Message "Successfully set primary emailaddress to [$newPrimaryMail] for mailbox [$UserPrincipalName]" -Event Success

} catch {
    HID-Write-Status -Message "Failed to set the primary emailaddress to [$newPrimaryMail] for mailbox [$UserPrincipalName]" -Event Error
    throw $_
}

# Disconnect from Exchange
try {
    Remove-PSSession -Session $exchangeSession -Confirm:$false -ErrorAction Stop
    HID-Write-Status -Message "Successfully disconnected from Exchange" -Event Success
} catch {
    HID-Write-Status -Message "Error disconnecting from Exchange" -Event Error
    HID-Write-Status -Message "Error at line: $($_.InvocationInfo.ScriptLineNumber - 79): $($_.Exception.Message)" -Event Error
    if ($debug -eq $true) {
        HID-Write-Status -Message "$($_.Exception)" -Event Error
    }
    HID-Write-Summary -Message "Failed to disconnect from Exchange" -Event Failed
    throw $_
}
<#----- Exchange On-Premises: End -----#>

