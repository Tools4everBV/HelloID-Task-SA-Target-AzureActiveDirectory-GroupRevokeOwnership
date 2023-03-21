# HelloID-Task-SA-Target-AzureActiveDirectory-GroupRevokeOwnership
##################################################################
# Form mapping
$formObject = @{
    groupId     = $form.groupId
    ownersToRevoke = $form.ownersToRevoke  
}
try {
    Write-Information "Executing AzureActiveDirectory action: [GroupRevokeOwnership] for: [$($formObject.groupId)]"
    Write-Information "Retrieving Microsoft Graph AccessToken for tenant: [$AADTenantID]"
    $splatTokenParams = @{
        Uri         = "https://login.microsoftonline.com/$($AADTenantID)/oauth2/token"
        ContentType = 'application/x-www-form-urlencoded'
        Method      = 'POST'
        Body        = @{                                                                                                                         
            grant_type    = 'client_credentials'
            client_id     = $AADAppID
            client_secret = $AADAppSecret
            resource      = 'https://graph.microsoft.com'
        }
    }
    $accessToken = (Invoke-RestMethod @splatTokenParams).access_token

    $headers = [System.Collections.Generic.Dictionary[string, string]]::new()
    $headers.Add("Authorization", "Bearer $($accessToken)")
    $headers.Add("Content-Type", "application/json")

    foreach ($owner in $formObject.ownersToRevoke){
        try{
            $splatRevokeOwnerFromGroup = @{
                Uri         = "https://graph.microsoft.com/v1.0/groups/$($formObject.groupId)/owners/$($owner.userId)/`$ref"
                ContentType = 'application/json'
                Method      = 'DELETE'
                Headers     = $headers
            }
            $null = Invoke-RestMethod @splatRevokeOwnerFromGroup

            $auditLog = @{
                Action            = 'UpdateResource'
                System            = 'AzureActiveDirectory'
                TargetIdentifier  = $formObject.groupId
                TargetDisplayName = $formObject.groupId
                Message           = "AzureActiveDirectory action: [GroupRevokeOwnership] from group [$($formObject.groupId)] for: [$($owner.userPrincipalName)] executed successfully"
                IsError           = $false
            }

            Write-Information -Tags 'Audit' -MessageData $auditLog
            Write-Information "AzureActiveDirectory action: [GroupRevokeOwnership] from group [$($formObject.groupId)] for: [$($owner.userPrincipalName)] executed successfully"
        }catch{
            $ex = $_
            if (-not[string]::IsNullOrEmpty($ex.ErrorDetails)) {
                $errorExceptionDetails = ($ex.ErrorDetails | ConvertFrom-Json).error.Message
            }else {
                $errorExceptionDetails = $ex.Exception.Message
            }

            if (($ex.Exception.Response) -and ($Ex.Exception.Response.StatusCode -eq 404)) {
                # 404 indicates already removed
                   $auditLog = @{
                    Action            = 'UpdateResource'
                    System            = 'AzureActiveDirectory'
                    TargetIdentifier  = $formObject.groupId
                    TargetDisplayName = $formObject.groupId
                    Message           = "AzureActiveDirectory action: [GroupRevokeOwnership from group [$($formObject.groupId))] ] for: [$($owner.userPrincipalName)] executed successfully. Note that the account was not a owner"
                    IsError           = $false
                }
                Write-Information -Tags 'Audit' -MessageData $auditLog
                Write-Information "AzureActiveDirectory action: [GroupRevokeOwnership from group [$($formObject.groupId))] ] for: [$($owner.userPrincipalName)] executed successfully.  Note that the account was not a owner"
            }else {
                $auditLog = @{
                    Action            = 'UpdateResource'
                    System            = 'AzureActiveDirectory'
                    TargetIdentifier  = $formObject.groupId
                    TargetDisplayName = $formObject.groupId
                    Message           = "Could not execute AzureActiveDirectory action: [GroupRevokeOwnership] from group [$($formObject.groupId)] for: [$($owner.userPrincipalName)], error: $($errorExceptionDetails)"
                    IsError           = $true
                }
                Write-Information -Tags "Audit" -MessageData $auditLog
                Write-Error "Could not execute AzureActiveDirectory action: [GroupRevokeOwnership] from group [$($formObject.groupId)] for: [$($owner.userPrincipalName)], error: $($errorExceptionDetails)"
            }
        }
    }
}
catch {
    $ex = $_

    if (-not[string]::IsNullOrEmpty($ex.ErrorDetails)) {
        $errorExceptionDetails = ($ex.ErrorDetails | ConvertFrom-Json).error.Message
    }else {
        $errorExceptionDetails = $ex.Exception.Message
    }

    $auditLog = @{
        Action            = 'UpdateResource'
        System            = 'AzureActiveDirectory'
        TargetIdentifier  = $formObject.groupId
        TargetDisplayName = $formObject.groupId
        Message           = "Could not execute AzureActiveDirectory action: [GroupRevokeOwnership] from group [$($formObject.groupId)] for: [$($owner.userPrincipalName)], error: $($errorExceptionDetails)"
        IsError           = $true
    }
    Write-Information -Tags "Audit" -MessageData $auditLog
    Write-Error "Could not execute AzureActiveDirectory action: [GroupRevokeOwnership] from group [$($formObject.groupId)] for: [$($owner.userPrincipalName)], error: $($errorExceptionDetails)"
}
##################################################################