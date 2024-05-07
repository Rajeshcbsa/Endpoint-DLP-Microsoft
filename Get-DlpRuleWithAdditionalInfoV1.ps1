###  Get DLP Rule info information types ( With Additonal protocol info)  ##############################################################
$DateTimeStamp = Get-Date -Format "MM-dd-yyyy HH-mm-ss"
$FolderPath = Get-Location
$filenamep = $FolderPath.ToString() + "\Rulesstatus_" + $DateTimeStamp + ".csv"
$dlppolicy = Get-DlpCompliancePolicy
$dlppolicylist = $dlppolicy|Where-Object -Property EndpointDlpLocation -NE ""|Select-Object Priority,DisplayName,Mode,EndpointDlpLocation,EndpointDlpLocationException
#$dlpRules = Get-DlpComplianceRule
#$dlpRules | Where-Object ParentPolicyName -eq "USB-Monitoring"| Select-Object *
#$dlpRules |Select-Object DisplayName,Policy,ParentPolicyName,Mode,CreatedBy,WhenCreated,LastModifiedBy,WhenChanged,ExchangeObjectId,GenerateAlert
$resultsOut = @()
#$dlppolicy = Get-DlpCompliancePolicy -policy $_.DisplayName
$dlppolicylist | ForEach-Object {
        $n = $_.DisplayName
        $m = $_.Mode
        $pri = $_.Priority
        $scopes = $_.EndpointDlpLocation
        $scopesExcption = $_.EndpointDlpLocationException
        write-host $_.DisplayName        
        $dlpRules = Get-DlpComplianceRule -Policy $_.DisplayName
        $dlpRules |ForEach-Object{        
            if($_.ParentPolicyName -eq $n ){
            #write-host $_.DisplayName " " $n
                # $PolicyName= $_.ParentPolicyName
                # $RuleName= $_.DisplayName
                # $RuleMode= $_.Mode                        
                # $CreatedBy= $_.CreatedBy
                # $WhenCreated= $_.WhenCreated
                # $LastModifiedBy= $_.LastModifiedBy
                # $WhenChanged= $_.WhenChanged
                # $PolicyId= $_.Policy
                # $ExchangeObjectId= $_.ExchangeObjectId
                # $GenerateAlert= $_.GenerateAlert
                # $Protocols =  $_.EndpointDlpRestrictions
                $CloudEgress = "False"
                $RemovableMedia = "False"
                $NetworkShare = "False"
                $UnallowedApps = "False"
                $Print = "False"
                $RemoteDesktopServices = "False"
                $UnallowedBluetoothTransferApps = "False"
                $PasteToBrowser = "False"

                if($_.EndpointDlpRestrictions -ne $null){

                    Write-Host $_.EndpointDlpRestrictions.count
                    for($p=0;$p -le $_.EndpointDlpRestrictions.count; $p++)
                    {
                        #Write-Host $p
                        #Write-Host $_.EndpointDlpRestrictions[$p].setting
                        #Write-Host $_.EndpointDlpRestrictions[$p].value
                        if($_.EndpointDlpRestrictions[$p].setting -eq "CloudEgress"){
                            $CloudEgress = $_.EndpointDlpRestrictions[$p].value
                        }
                        if($_.EndpointDlpRestrictions[$p].setting -eq "RemovableMedia"){
                            $RemovableMedia = $_.EndpointDlpRestrictions[$p].value                                                
                        }
                        if($_.EndpointDlpRestrictions[$p].setting -eq "NetworkShare"){
                            $NetworkShare = $_.EndpointDlpRestrictions[$p].value                                                
                        }
                        if($_.EndpointDlpRestrictions[$p].setting -eq "UnallowedApps"){
                            $UnallowedApps = $_.EndpointDlpRestrictions[$p].value                                                
                        }
                        if($_.EndpointDlpRestrictions[$p].setting -eq "Print"){
                            $Print = $_.EndpointDlpRestrictions[$p].value                                                
                        }
                        if($_.EndpointDlpRestrictions[$p].setting -eq "RemoteDesktopServices"){
                            $RemoteDesktopServices = $_.EndpointDlpRestrictions[$p].value                                                
                        }
                        if($_.EndpointDlpRestrictions[$p].setting -eq "UnallowedBluetoothTransferApps"){
                            $UnallowedBluetoothTransferApps = $_.EndpointDlpRestrictions[$p].value                                                
                        }
                        if($_.EndpointDlpRestrictions[$p].setting -eq "PasteToBrowser"){
                            $PasteToBrowser = $_.EndpointDlpRestrictions[$p].value                                                
                        }
                    }
                }
                    Write-Host "PolicyName " $p
                        Write-Host "PolicyName " $_.ParentPolicyName
                        Write-Host "RuleName " $_.DisplayName
                        Write-Host "cloudegeress "  $CloudEgress        
                        Write-Host "removeable " $RemovableMedia
                        Write-Host "networkshare " $NetworkShare
                        Write-Host "Unallowedapp " $UnallowedApps
                        Write-Host "Print " $Print
                        Write-Host "RemoteDesktopServices " $RemoteDesktopServices
                        Write-Host "UnallowedBluetoothTransferApps " $UnallowedBluetoothTransferApps
                        Write-Host "PasteToBrowser " $PasteToBrowser
                #        Read-Host "waiting"              

                # Write-Host $_.Setting

                    $resultsOut += [PSCustomObject]@{
                    #Id = $json.Id
                    Priority = $pri
                    PolicyName= $_.ParentPolicyName
                    RuleName= $_.DisplayName
                    CloudEgress =  $CloudEgress        
                    RemovableMedia = $RemovableMedia
                    NetworkShare = $NetworkShare
                    UnallowedApps = $UnallowedApps
                    Print = $Print
                    RemoteDesktopServices = $RemoteDesktopServices
                    UnallowedBluetoothTransferApps = $UnallowedBluetoothTransferApps
                    PasteToBrowser = $PasteToBrowser
                    RuleMode= $_.Mode
                    PolicyMode = $m                    
                    Scopes = $scopes
                    ScopesExcption = $scopesExcption
                    GenerateAlert= $_.GenerateAlert
                    CreatedBy= $_.CreatedBy
                    WhenCreated= $_.WhenCreated
                    LastModifiedBy= $_.LastModifiedBy
                    WhenChanged= $_.WhenChanged
                    PolicyId= $_.Policy
                    ExchangeObjectId= $_.ExchangeObjectId
                    }              
            }          
        }
}
    #$resultsOut = ""
#write-host $resultsOut
$resultsOut | Export-Csv -Path $filenamep -NoTypeInformation -Append

######################################################################################################