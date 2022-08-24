$VerbosePreference = "SilentlyContinue"
$InformationPreference = "Continue"
$WarningPreference = "Continue"

# variables configured in form
$UserPrincipalName = $form.gridmailbox.Userprincipalname
$isPrimary = $form.grid.IsPrimary
$newPrimaryMail = $form.grid.EmailAddress

# Connect to Exchange
try {
    $adminSecurePassword = ConvertTo-SecureString -String "$ExchangeAdminPassword" -AsPlainText -Force
    $adminCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $ExchangeAdminUsername, $adminSecurePassword
    $sessionOption = New-PSSessionOption -SkipCACheck -SkipCNCheck -SkipRevocationCheck
    $exchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $exchangeConnectionUri -Credential $adminCredential -SessionOption $sessionOption -ErrorAction Stop 
    #-AllowRedirection
    $session = Import-PSSession $exchangeSession -DisableNameChecking -AllowClobber
    Write-Information "Successfully connected to Exchange using the URI [$exchangeConnectionUri]" 
    
    $Log = @{
        Action            = "UpdateAccount" # optional. ENUM (undefined = default) 
        System            = "Exchange On-Premise" # optional (free format text) 
        Message           = "Successfully connected to Exchange using the URI [$exchangeConnectionUri]" # required (free format text) 
        IsError           = $false # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
        TargetDisplayName = $exchangeConnectionUri # optional (free format text) 
        TargetIdentifier  = $([string]$session.GUID) # optional (free format text) 
    }
    #send result back  
    Write-Information -Tags "Audit" -MessageData $log
}
catch {
    Write-Error "Error connecting to Exchange using the URI [$exchangeConnectionUri]. Error: $($_.Exception.Message)"
    $Log = @{
        Action            = "UpdateAccount" # optional. ENUM (undefined = default) 
        System            = "Exchange On-Premise" # optional (free format text) 
        Message           = "Failed to connect to Exchange using the URI [$exchangeConnectionUri]." # required (free format text) 
        IsError           = $true # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
        TargetDisplayName = $exchangeConnectionUri # optional (free format text) 
        TargetIdentifier  = $([string]$session.GUID) # optional (free format text) 
    }
    #send result back  
    Write-Information -Tags "Audit" -MessageData $log
}

if ($isPrimary -eq "true") {
    Write-Information "No changes where made. The selected primary address [$newPrimaryMail] is the same as the existing primary address"        
    $Log = @{
        Action            = "UpdateAccount" # optional. ENUM (undefined = default) 
        System            = "Exchange On-Premise" # optional (free format text) 
        Message           = "No changes where made. The selected primary address [$newPrimaryMail] is the same as the existing primary address" # required (free format text) 
        IsError           = $false # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
        TargetDisplayName = $UserPrincipalName # optional (free format text) 
        TargetIdentifier  = $newPrimaryMail # optional (free format text) 
    }
    #send result back  
    Write-Information -Tags "Audit" -MessageData $log
}
else {
    try {
    
    
    

        $ParamsGetMailbox = @{
            Identity = $UserPrincipalName
        }
        $mailBoxes = Invoke-Command -Session $exchangeSession -ErrorAction Stop -ScriptBlock {
            Param ($ParamsGetMailbox)
            Get-Mailbox @ParamsGetMailbox
        } -ArgumentList $ParamsGetMailbox
    
        $list = New-Object System.Collections.ArrayList
        foreach ($address in $mailBoxes.EmailAddresses) {
            $prefix = $address.Split(":")[0]
            $mail = $address.Split(":")[1]
            if ($mail.ToLower() -eq $newPrimaryMail.ToLower()) {
                $address = "SMTP:" + $mail
            }
            else {
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

        Write-Information "Successfully set primary emailaddress to [$newPrimaryMail] for mailbox [$UserPrincipalName]"
        $Log = @{
            Action            = "UpdateAccount" # optional. ENUM (undefined = default) 
            System            = "Exchange On-Premise" # optional (free format text) 
            Message           = "Successfully set primary emailaddress to [$newPrimaryMail] for mailbox [$UserPrincipalName]." # required (free format text) 
            IsError           = $false # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
            TargetDisplayName = $newPrimaryMail # optional (free format text) 
            TargetIdentifier  = $([string]$mailBoxes.GUID) # optional (free format text) 
        }
        #send result back  
        Write-Information -Tags "Audit" -MessageData $log    

    }
    catch {
        Write-Error "Failed to set the primary emailaddress to [$newPrimaryMail] for mailbox [$UserPrincipalName]. Error: $($_.Exception.Message)"
        $Log = @{
            Action            = "UpdateAccount" # optional. ENUM (undefined = default) 
            System            = "Exchange On-Premise" # optional (free format text) 
            Message           = "Failed to set the primary emailaddress to [$newPrimaryMail] for mailbox [$UserPrincipalName]." # required (free format text) 
            IsError           = $true # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
            TargetDisplayName = $newPrimaryMail # optional (free format text) 
            TargetIdentifier  = $([string]$mailBoxes.GUID) # optional (free format text) 
        }
        #send result back  
        Write-Information -Tags "Audit" -MessageData $log 
    }
}

# Disconnect from Exchange
try {
    Remove-PsSession -Session $exchangeSession -Confirm:$false -ErrorAction Stop
    Write-Information "Successfully disconnected from Exchange using the URI [$exchangeConnectionUri]"     
    $Log = @{
        Action            = "UpdateAccount" # optional. ENUM (undefined = default) 
        System            = "Exchange On-Premise" # optional (free format text) 
        Message           = "Successfully disconnected from Exchange using the URI [$exchangeConnectionUri]" # required (free format text) 
        IsError           = $false # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
        TargetDisplayName = $exchangeConnectionUri # optional (free format text) 
        TargetIdentifier  = $([string]$session.GUID) # optional (free format text) 
    }
    #send result back  
    Write-Information -Tags "Audit" -MessageData $log
}
catch {
    Write-Error "Error disconnecting from Exchange.  Error: $($_.Exception.Message)"
    $Log = @{
        Action            = "UpdateAccount" # optional. ENUM (undefined = default) 
        System            = "Exchange On-Premise" # optional (free format text) 
        Message           = "Failed to disconnect from Exchange using the URI [$exchangeConnectionUri]." # required (free format text) 
        IsError           = $true # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
        TargetDisplayName = $exchangeConnectionUri # optional (free format text) 
        TargetIdentifier  = $([string]$session.GUID) # optional (free format text) 
    }
    #send result back  
    Write-Information -Tags "Audit" -MessageData $log
}
<#----- Exchange On-Premises: End -----#>

