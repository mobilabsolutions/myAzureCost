param (
    [parameter(Mandatory = $true,
        HelpMessage = "Enter a tag label e.g. 'ms-resource-usage'")]
    [String]$tagLabel,
    [parameter(Mandatory = $true,
        HelpMessage = "Enter a tag value e.g. 'azure-cloud-shell'")]
    [String]$tagValue
)

[void][Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 
[void][Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms.DataVisualization")

#Loginto Azure subscription - Get Execution Context.
$connectionName = "AzureRunAsConnection"
$AzureSubscriptionId = "myAzureCostAzureSubscriptionId"
$storageAccount = Get-AutomationVariable -Name "myAzureCostStorageAccountName"
$tableName = Get-AutomationVariable -Name "myAzureCostSATable"
$containerName = Get-AutomationVariable -Name "myAzureCostSAContainer"
$myAzureCostPriceSheetURI = Get-AutomationVariable -Name 'myAzureCostPriceSheetURI'

try {
    # Get the connection "AzureRunAsConnection "
    $servicePrincipalConnection = Get-AutomationConnection -Name $connectionName  
    $subscriptionID = Get-AutomationVariable -Name $AzureSubscriptionId  

    "Logging in to Azure..."
    $account = Login-AzAccount `
        -ServicePrincipal `
        -TenantId $servicePrincipalConnection.TenantId `
        -ApplicationId $servicePrincipalConnection.ApplicationId `
        -CertificateThumbprint $servicePrincipalConnection.CertificateThumbprint `
        -Environment AzureCloud

    "Setting subscription context to $subscriptionID..."
    Set-AzContext -SubscriptionId $subscriptionID

    "Login result:" 
    Write-Output $account

    $date = [dateTime]::Today.AddMonths(-1)
    $year = $date.Year
    $month = $date.Month

    $startOfMonth = Get-Date -Year $year -Month $month -Day 1 -Hour 0 -Minute 0 -Second 0 -Millisecond 0
    $endOfMonth = $startOfMonth.AddMonths(1).AddDays(-1).AddHours(23)

    $ConsumptionDate = $startOfMonth
    $EndDate = $endOfMonth

    Write-Output "Fill Table with costs between $($ConsumptionDate.ToString("dd'/'MM'/'yyyy")) and $($EndDate.ToString("dd'/'MM'/'yyyy"))"

    $currentDate = $ConsumptionDate

    $RGName = (Get-AzResource -Name $storageAccount -ResourceType 'Microsoft.Storage/storageAccounts').ResourceGroupName
    $sa = Get-AzStorageAccount -Name $storageAccount -ResourceGroupName $RGName
    $ctx = $sa.Context
    $cloudTable = (Get-AzStorageTable –Name $tableName –Context $ctx).CloudTable

    $priceListPath = "$Env:temp\$($(Split-Path $myAzureCostPriceSheetURI -Leaf) -replace "[\?]{1}.*",'')" 
    "downloading pricelist..."
    Invoke-WebRequest -Uri $myAzureCostPriceSheetURI -OutFile $priceListPath
    "found in $($priceListPath): $(Test-Path $priceListPath)" 
    if (!(Test-Path $priceListPath)) {
        "...no pricelist file found - exit!"
        $ErrorActionPreference = 'Stop'
        Get-Content $priceListPath
    }

    while ($currentDate -le $EndDate) {
        Write-Output "Insert cost of $($currentDate.ToString("dd'/'MM'/'yyyy"))"
 
        $UsageAggregations = @()
        $ErrorActionPreference = "SilentlyContinue"
        $UsageAggregates = $null
        do {
            if ($UsageAggregates.ContinuationToken) {
                "continue"
                $UsageAggregates = Get-UsageAggregates -ContinuationToken $($UsageAggregates.ContinuationToken) -ShowDetails $true -Verbose -ReportedStartTime $currentDate -ReportedEndTime $currentDate.addHours(25) -AggregationGranularity Hourly
            }
            else {
                "first data"
                $UsageAggregates = Get-UsageAggregates -ShowDetails $true -Verbose -ReportedStartTime $currentDate -ReportedEndTime $currentDate.addHours(25) -AggregationGranularity Hourly
            }

            foreach ($item in $UsageAggregates.UsageAggregations) {
                $UsageAggregations += $item
            }
        }
        while ($UsageAggregates.ContinuationToken)

        $UsageToExport = $UsageAggregations | % { $_.Properties | select-object UsageStartTime, UsageEndTime, MeterCategory, MeterSubCategory, MeterName, @{N = 'InstanceName'; E = { ($_.InstanceData | ConvertFrom-Json).'Microsoft.Resources'.resourceUri.Split('/') | select -Last 1 } }, @{N = 'RG'; E = { ($_.InstanceData | ConvertFrom-Json).'Microsoft.Resources'.resourceUri.Split('/')[4] } }, @{N = 'Location'; E = { ($_.InstanceData | ConvertFrom-Json).'Microsoft.Resources'.location } }, @{N = 'Quantity'; E = { $_.Quantity } }, Unit, MeterId, @{N = 'Tags'; E = { ($_.InstanceData | ConvertFrom-Json).'Microsoft.Resources'.tags } } } | where { ($(get-Date $_.UsageStartTime) -ge $(Get-date $currentDate.ToShortDateString()) -and ($(get-Date $_.UsageStartTime) -lt $(Get-date $currentDate.AddDays(1).ToShortDateString())) -and (($_.Tags -like "*$tagLabel=$tagValue*")) ) } 
        # sum up quantities of instances with same MeterID, date and rg 
        $data = $UsageToExport | Group-Object InstanceName, RG, MeterID
        $result = @()
        $result += foreach ($item in $data) {
            $item.Group | Select-Object -Unique @{N = 'UsageStartTime'; E = { $($currentDate.ToString("d")) } }, @{N = 'UsageEndTime'; E = { $($currentDate.AddDays(1).ToString("d")) } }, MeterCategory, MeterSubCategory, MeterName, InstanceName, RG, Location, @{Name = 'Quantity'; Expression = { (($item.Group) | Measure-Object -Property Quantity -sum).Sum } }, Unit, MeterId, Tags
        }

        $result | Export-Csv "$Env:temp/Usage.csv" -Encoding UTF8 -Delimiter ';' -NoTypeInformation

        $fileName = "$($currentDate.ToString("yyyyMMdd"))Consumption.csv"

        Set-AzStorageBlobContent -Container $containerName -Context $ctx -File "$Env:temp/Usage.csv" -Blob $fileName -Force

        $blob = Get-AzStorageBlob -Container $containerName -Context $ctx -Blob "$fileName"
        "found: $($blob.Name)"
        $token = New-AzStorageBlobSASToken -Context $ctx  -CloudBlob $($blob.ICloudBlob) -StartTime ([datetime]::Now).AddHours(-1) -ExpiryTime ([datetime]::Now).AddHours(1) -Permission 'r'
        $uri = "https://$storageAccount.blob.core.windows.net/$containerName/$fileName$token"
        $usagePath = "$Env:temp\$($(Split-Path $uri -Leaf) -replace "[\?]{1}.*",'')" 
        "downloading usage..."
        Invoke-WebRequest -Uri $uri -OutFile $usagePath
        "found in $($usagePath): $(Test-Path $usagePath)" 
        if (!(Test-Path $usagePath)) {
            "...no usage file found - exit!"
            $ErrorActionPreference = 'Stop'
            Get-Content $usagePath
        }

        #selectively fill pricelist object
        $uniqueMeterIDs = Import-Csv -Path $usagePath -Delimiter ';' -Encoding UTF8 | % { $_.MeterID } | Select-Object -Unique
        $priceList = Import-Csv $priceListPath -Delimiter ';'  -Encoding UTF8 | where { $uniqueMeterIDs -contains $_.MeterId }

        $usageEntries = Import-Csv -Path $usagePath -Delimiter ';' -Encoding UTF8

        #calculate usage
        $costEntries = @()
        foreach ($usageEntry in $usageEntries) {
            Write-Host "." -NoNewline
            $costEntries += $usageEntry | select-object UsageStartTime, UsageEndTime, MeterCategory, MeterSubCategory, MeterName, InstanceName, RG, Location, @{N = 'Quantity'; E = { [decimal]$_.Quantity } }, Unit, @{N = 'UnitPrice'; E = { $MeterID = $_.MeterId ; ($priceList | where { $_.MeterId -eq $MeterID }).MeterRates -match "([0-9]+\.[0-9]*)" | Out-Null ; $price = [decimal]0; $price = ([decimal]($Matches[1])) ; $price } }, @{N = 'Estimated Costs'; E = { $MeterID = $_.MeterID ; ($priceList | where { $_.MeterId -eq $MeterID }).MeterRates -match "([0-9]+\.[0-9]*)" | Out-Null ; $price = [decimal]0; $price = ([decimal]($Matches[1]) * [decimal]$_.Quantity) ; $price } } , MeterID, Tags
        }

        $totalCost = $($costEntries | Measure-Object 'Estimated Costs' -Sum).Sum

        #region save history data in Table
        try {
            $entry = Get-AzTableRow -Table $cloudTable -PartitionKey $currentDate.ToString('MMMM') -rowKey "$($currentDate.ToString('dd'))"
            $entry.TotalCost = "{0:N2}" -f $totalCost
            $entry.Year = $currentDate.Year
            $entry | Update-AzTableRow -table $cloudTable
        }
        catch {
            Add-AzTableRow -table $cloudTable -partitionKey $currentDate.ToString('MMMM') `
                -rowKey "$($currentDate.ToString('dd'))" -property @{"TotalCost" = $("{0:N2}" -f $totalCost); "Year" = $currentDate.Year }
        }

        $currentDate = $currentDate.AddDays(1)
    }
}
catch {
    if (!$servicePrincipalConnection) {
        $ErrorMessage = "Connection $connectionName not found."
        throw $ErrorMessage
    }
    else {
        Write-Error -Message $_.Exception
        throw $_.Exception
    }
} 
Write-Output $account