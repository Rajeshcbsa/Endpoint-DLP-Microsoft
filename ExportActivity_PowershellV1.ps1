param 
(
    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string] $UserPrincipalName,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string] $EndTime,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string] $StartTime,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string] $OutputFormat,

    [Parameter(Mandatory = $false)]
    [string] $PageCookie,

    [Parameter(Mandatory = $false)]
    [int] $PageSize,

    [Parameter(Mandatory = $false)]
    [string[]] $Filter1,

    [Parameter(Mandatory = $false)]
    [string[]] $Filter2,

    [Parameter(Mandatory = $false)]
    [string[]] $Filter3,

    [Parameter(Mandatory = $false)]
    [string[]] $Filter4,

    [Parameter(Mandatory = $false)]
    [string[]] $Filter5,

    [Parameter(Mandatory = $false)]
    [string] $FileName,

    [Parameter(Mandatory = $false)]
    [string] $DetailsFileName,

    [Parameter(Mandatory = $false)]
    [string] $FolderPath,

    [Parameter(Mandatory =  $false)]
    [int] $RecordsCount,
    
    [Switch] $Details
)


if (!$FolderPath)
{
    $FolderPath = Get-Location
}

$DateTimeStamp = Get-Date -Format "MM-dd-yyyy HH-mm-ss"
if (!$FileName)
{
    $FileName = "Export-ActivityExplorerData " + $DateTimeStamp
}

if ($OutputFormat -eq 'csv')
{
    $FileName = $FileName + '.csv'
} 
else
{
    $FileName = $FileName + '.json'
}
$FilePath = $FolderPath + '\' + $FileName
New-Item -Path $FolderPath -Name $FileName
if ($OutputFormat -eq 'json')
{
    Add-Content $FilePath '[' -NoNewline
}

if ($Details -or $DetailsFileName)
{
    if(!$DetailsFileName)
    {
        $DetailsFileName = "Export-ActivityExplorerData Details " + $DateTimeStamp + '.txt'  
    }
    $DetailsFilePath = $FolderPath + '\' + $DetailsFileName
    New-Item -Path $FolderPath -Name $DetailsFileName
}

$stopwatch = [System.Diagnostics.Stopwatch]::new()
$stopwatch.Start()
if ($UserPrincipalName)
{
    Import-Module ExchangeOnlineManagement
    Connect-IPPSSession -UserPrincipalName $UserPrincipalName
}
 
[System.Collections.ArrayList]$FiltersList = @()
if($Filter1)
{
    $FiltersList.Add($Filter1)
}
if($Filter2)
{
    $FiltersList.Add($Filter2)
}
if($Filter3)
{
    $FiltersList.Add($Filter3)
}
if($Filter4)
{
    $FiltersList.Add($Filter4)
}
if($Filter5)
{
    $FiltersList.Add($Filter5)
}

$RecordsExported = 0

if ($RecordsCount)
{
    if (!$PageSize)
    {
        Write-Error -Message "Please provide PageSize with RecordsCount"
    }
    else
    {
        if ($RecordsCount -lt $PageSize)
        {
            Write-Error -Message "RecordsCount to be returned should be multiple of PageSize"
        }
        else
        {
            if ($RecordsCount % $PageSize -ne 0)
            {
                $NumberOfQueries = [Math]::Ceiling($RecordsCount / $PageSize)
                $NewReccordsCount = $count * $PageSize
                Write-Host "RecordsCount to be returned should be multiple of PageSize. Returning " + $NewReccordsCount + " instead of " + $RecordsCount + "."
                $RecordsCount = $NewReccordsCount
            }
            else
            {
                $NumberOfQueries = $RecordsCount / $PageSize
            }
            for ($QueryCounter = 0; $QueryCounter -lt $NumberOfQueries; $QueryCounter++)
            {
                try
                {
                    if ($FiltersList.Count -eq 0)
                    {
                        $res = Export-ActivityExplorerData -StartTime $startTime -EndTime $endTime -PageSize $pageSize -PageCookie $pageCookie -OutputFormat json
                    }
                    elseif ($FiltersList.Count -eq 1)
                    {
                        $res = Export-ActivityExplorerData -StartTime $startTime -EndTime $endTime -PageSize $pageSize -PageCookie $pageCookie -OutputFormat json -Filter1 $FiltersList[0]
                    }
                    elseif ($FiltersList.Count -eq 2)
                    {
                        $res = Export-ActivityExplorerData -StartTime $startTime -EndTime $endTime -PageSize $pageSize -PageCookie $pageCookie -OutputFormat json -Filter1 $FiltersList[0] -Filter2 $FiltersList[1]
                    }
                    elseif ($FiltersList.Count -eq 3)
                    {
                        $res = Export-ActivityExplorerData -StartTime $startTime -EndTime $endTime -PageSize $pageSize -PageCookie $pageCookie -OutputFormat json -Filter1 $FiltersList[0] -Filter2 $FiltersList[1] -Filter3 $FiltersList[2]
                    }
                    elseif ($FiltersList.Count -eq 4)
                    {
                        $res = Export-ActivityExplorerData -StartTime $startTime -EndTime $endTime -PageSize $pageSize -PageCookie $pageCookie -OutputFormat json -Filter1 $FiltersList[0] -Filter2 $FiltersList[1] -Filter3 $FiltersList[2] -Filter4 $FiltersList[3]
                    }
                    elseif ($FiltersList.Count -eq 5)
                    {
                        $res = Export-ActivityExplorerData -StartTime $startTime -EndTime $endTime -PageSize $pageSize -PageCookie $pageCookie -OutputFormat json -Filter1 $FiltersList[0] -Filter2 $FiltersList[1] -Filter3 $FiltersList[2] -Filter4 $FiltersList[3] -Filter5 $FiltersList[4]
                    }
                    $pageCookie = $res.WaterMark

                    $RecordsExported = $RecordsExported + $res.RecordCount

                    Write-host "Total Records to be exported : "$res.TotalResultCount


                    if ($OutputFormat -eq "csv")
                    {
                        
                        $results = @()
                        $d = $res.ResultData | ConvertFrom-Json


                        foreach ($x in $d) {
 
                            $json = $x #| ConvertFrom-Json
                            if($json.SensitiveInfoTypeData.SensitiveInformationDetectionsInfo.DetectedValues)
                                {
                                   $DetectedValues = $json.SensitiveInfoTypeData.SensitiveInformationDetectionsInfo.DetectedValues[0].Name.ToString()
                                   $SensitiveInfoTypeId = $json.SensitiveInfoTypeData[0].SensitiveInfoTypeId.ToString()
                                }
                                else
                                {
                                   $DetectedValues = ""
                                   $SensitiveInfoTypeId = ""
                                }

                           
                            # Parse the string into a DateTime object
                            $datetime = [DateTime]::ParseExact($json.Happened, "yyyy-MM-ddTHH:mm:ssZ", [System.Globalization.CultureInfo]::InvariantCulture)
                            $formattedDateTime = $datetime.ToString("yyyy-MM-dd HH:mm:ss")

                            
                            $results += [PSCustomObject]@{
                                #Id = $json.Id
                                Activity = $json.Activity
                                Happened = $formattedDateTime
                                User = $json.User
                                FilePath = $json.FilePath
                                ClientIP = $json.ClientIP
                                FileExtension = $json.FileExtension
                                Platform = $json.Platform
                                Application = $json.Application
                                ProcessName = $json.ProcessName
                                DeviceName = $json.DeviceName
                                EnforcementMode = $json.EnforcementMode
                                TargetDomain = $json.TargetDomain
                                TargetFilePath = $json.TargetFilePath
                                FileSize = $json.FileSize
                                PolicyName = $json.PolicyMatchInfo.PolicyName
                                RuleName = $json.PolicyMatchInfo.RuleName
                                EndpointOperation = $json.EndpointOperation
                                StorageName = $json.EvidenceFile.StorageName
                                FullUrl = $json.EvidenceFile.FullUrl
                                SensitiveInfoTypeData = $json.SensitiveInfoTypeData[0].SensitiveInfoTypeId
                                DetectedValues = $DetectedValues
                                #SensitiveInfoTypeData = $json.SensitiveInfoTypeData[0].SensitiveInfoTypeId
                                SensitiveInfoTypeId = $SensitiveInfoTypeId
                              }
 
                           }
                           #$FilePath = "d:\test\Export-ActivityExplorerData 01-13-2024 23-41-28.csv"
                           $results | Export-Csv -Path $FilePath -NoTypeInformation -Append
                    }
                    else
                    {
                        if ($QueryCounter -eq 0)
                        {
                            if ($res.RecordCount -ne 0)
                            {
                                Add-Content $FilePath ($res.ResultData.Substring(1, ($res.ResultData.Length-2))) -NoNewline   
                            }
                            Write-Host ('Total records to be exported: ' + $RecordsCount + '.')
                        }
                        else
                        {
                            if ($res.RecordCount -ne 0)
                            {
                                Add-Content $FilePath (',' + $res.ResultData.Substring(1, ($res.ResultData.Length-2))) -NoNewline
                            }
                        }
                    } 

                    Write-Host ("$RecordsExported records are exported. Time Elapsed: " + $stopwatch.Elapsed.TotalSeconds + "s.")

                    if ($Details -or $DetailsFileName)
                    {
                        Add-Content $DetailsFilePath ('Query ' + ($QueryCounter + 1) + ' Execution Details')
                        Add-Content $DetailsFilePath ('DataType: ' + $res.DataType)
                        Add-Content $DetailsFilePath ('Database: ' + $res.Database)
                        Add-Content $DetailsFilePath ('WaterMark: ' + $res.WaterMark)
                        Add-Content $DetailsFilePath ('LastPage: ' + $res.LastPage)
                        Add-Content $DetailsFilePath ('TotalResultCount: ' + $res.TotalResultCount)
                        Add-Content $DetailsFilePath ('RecordCount: ' + $res.RecordCount)
                        Add-Content $DetailsFilePath ('ErrorData: ' + $res.ErrorData)
                        Add-Content $DetailsFilePath ('ResultCode: ' + $res.ResultCode)
                        Add-Content $DetailsFilePath ''
                    }         
                }
                catch
                {
                    Write-Error $_.message
                    break
                }
            }
            if ($OutputFormat -eq "json")
            {
                Add-Content $FilePath ']'
            }
        }
    }
}
else
{
    If (!$PageSize)
    {
        $PageSize = 5000
    }
    $FirstQuery = 1  
    Do
    {
        try
        {
            if ($FiltersList.Count -eq 0)
            {
                $res = Export-ActivityExplorerData -StartTime $startTime -EndTime $endTime -PageSize $pageSize -PageCookie $pageCookie -OutputFormat json
            }
            elseif ($FiltersList.Count -eq 1)
            {
                $res = Export-ActivityExplorerData -StartTime $startTime -EndTime $endTime -PageSize $pageSize -PageCookie $pageCookie -OutputFormat json -Filter1 $FiltersList[0]
            }
            elseif ($FiltersList.Count -eq 2)
            {
                $res = Export-ActivityExplorerData -StartTime $startTime -EndTime $endTime -PageSize $pageSize -PageCookie $pageCookie -OutputFormat json -Filter1 $FiltersList[0] -Filter2 $FiltersList[1]
            }
            elseif ($FiltersList.Count -eq 3)
            {
                $res = Export-ActivityExplorerData -StartTime $startTime -EndTime $endTime -PageSize $pageSize -PageCookie $pageCookie -OutputFormat json -Filter1 $FiltersList[0] -Filter2 $FiltersList[1] -Filter3 $FiltersList[2]
            }
            elseif ($FiltersList.Count -eq 4)
            {
                Write-Host "Filter4 - -----------------------------------------------------"
                $res = Export-ActivityExplorerData -StartTime $startTime -EndTime $endTime -PageSize $pageSize -PageCookie $pageCookie -OutputFormat json -Filter1 $FiltersList[0] -Filter2 $FiltersList[1] -Filter3 $FiltersList[2] -Filter4 $FiltersList[3]
            }
            elseif ($FiltersList.Count -eq 5)
            {
                $res = Export-ActivityExplorerData -StartTime $startTime -EndTime $endTime -PageSize $pageSize -PageCookie $pageCookie -OutputFormat json -Filter1 $FiltersList[0] -Filter2 $FiltersList[1] -Filter3 $FiltersList[2] -Filter4 $FiltersList[3] -Filter5 $FiltersList[4]
            }
            $pageCookie = $res.WaterMark

            $RecordsExported = $RecordsExported + $res.RecordCount

            if ($OutputFormat -eq "csv")
            {
                
                $results = @()
                        $d = $res.ResultData | ConvertFrom-Json


                        foreach ($x in $d) {
 
                            $json = $x #| ConvertFrom-Json
                            if($json.SensitiveInfoTypeData.SensitiveInformationDetectionsInfo.DetectedValues)
                                {
                                   $DetectedValues = $json.SensitiveInfoTypeData.SensitiveInformationDetectionsInfo.DetectedValues[1].Name.ToString()
                                   $SensitiveInfoTypeId = $json.SensitiveInfoTypeData[0].SensitiveInfoTypeId.ToString()

                                   $json.SensitiveInfoTypeData.SensitiveInformationDetectionsInfo.DetectedValues.Count
                                }
                                else
                                {
                                   $DetectedValues = ""
                                   $SensitiveInfoTypeId = ""
                                }

                          

                            # Parse the string into a DateTime object
                            $datetime = [DateTime]::ParseExact($json.Happened, "yyyy-MM-ddTHH:mm:ssZ", [System.Globalization.CultureInfo]::InvariantCulture)
                            $formattedDateTime = $datetime.ToString("yyyy-MM-dd HH:mm:ss")

                            $results += [PSCustomObject]@{
                                #Id = $json.Id
                                Activity = $json.Activity
                                Happened = $formattedDateTime
                                User = $json.User
                                FilePath = $json.FilePath
                                ClientIP = $json.ClientIP
                                FileExtension = $json.FileExtension
                                Platform = $json.Platform
                                Application = $json.Application
                                ProcessName = $json.ProcessName
                                DeviceName = $json.DeviceName
                                EnforcementMode = $json.EnforcementMode
                                TargetDomain = $json.TargetDomain
                                TargetFilePath = $json.TargetFilePath
                                FileSize = $json.FileSize
                                PolicyName = $json.PolicyMatchInfo.PolicyName
                                RuleName = $json.PolicyMatchInfo.RuleName
                                EndpointOperation = $json.EndpointOperation
                                StorageName = $json.EvidenceFile.StorageName
                                FullUrl = $json.EvidenceFile.FullUrl
                                SensitiveInfoTypeData = $json.SensitiveInfoTypeData[0].SensitiveInfoTypeId
                                DetectedValues = $DetectedValues
                                #SensitiveInfoTypeData = $json.SensitiveInfoTypeData[0].SensitiveInfoTypeId
                                SensitiveInfoTypeId = $SensitiveInfoTypeId
                              }
 
                           }
                           $results | Export-Csv -Path $FilePath -NoTypeInformation -Append
 
            }
            else
            {
                if ($FirstQuery -eq 1)
                {
                    if ($res.RecordCount -ne 0)
                        {
                            Add-Content $FilePath ($res.ResultData.Substring(1, ($res.ResultData.Length-2))) -NoNewline   
                        }
                    Write-Host ('Total records to be exported: ' + $res.TotalResultCount + '.')
                }
                else
                {
                    if ($res.RecordCount -ne 0)
                        {
                            Add-Content $FilePath (',' + $res.ResultData.Substring(1, ($res.ResultData.Length-2))) -NoNewline
                        }
                }
            }

            Write-Host ("$RecordsExported records are exported. Time Elapsed: " + $stopwatch.Elapsed.TotalSeconds + "s.")

            if ($Details -or $DetailsFileName)
            {
                Add-Content $DetailsFilePath ('Query ' + $FirstQuery + ' Execution Details')
                Add-Content $DetailsFilePath ('DataType: ' + $res.DataType)
                Add-Content $DetailsFilePath ('Database: ' + $res.Database)
                Add-Content $DetailsFilePath ('WaterMark: ' + $res.WaterMark)
                Add-Content $DetailsFilePath ('LastPage: ' + $res.LastPage)
                Add-Content $DetailsFilePath ('TotalResultCount: ' + $res.TotalResultCount)
                Add-Content $DetailsFilePath ('RecordCount: ' + $res.RecordCount)
                Add-Content $DetailsFilePath ('ErrorData: ' + $res.ErrorData)
                Add-Content $DetailsFilePath ('ResultCode: ' + $res.ResultCode)
                Add-Content $DetailsFilePath ''
            }
            $FirstQuery = $FirstQuery + 1
        }
        catch
        {
            Write-Error $_.message
            break
        }
    }While(!($res.LastPage))
    if ($OutputFormat -eq "json")
    {
        Add-Content $FilePath ']'
    }
}

Write-Host ("Total time taken to export: " + $stopwatch.Elapsed.TotalSeconds + "s.")

$DisconnectSession = Read-Host -Prompt "Press Y to disconnect from the session otherwise press any other key"

if ($DisconnectSession -eq 'Y')
{
    Disconnect-ExchangeOnline
}

#Step 1 – save file as “ExportActivityExplorerData.ps1”

#Connect-IPPSSession

#.\ExportActivity_PowershellV1.ps1 -StartTime "01/05/2024 10:00 AM" -EndTime "01/09/2024 11:00 AM" -OutputFormat "csv" -PageSize 50 -Filter1 @("Activity", "DLPRuleMatch") -Filter2 @("Workload","Endpoint") -Filter3 @("PolicyName","USB-Monitoring")

#.\ExportActivity_PowershellV1.ps1 -StartTime "01/13/2024 00:00:01" -EndTime "01/14/2024 23:59:59" -OutputFormat "csv" -PageSize 50 -Filter1 @("Activity", "DLPRuleMatch") -Filter2 @("Workload","Endpoint")


