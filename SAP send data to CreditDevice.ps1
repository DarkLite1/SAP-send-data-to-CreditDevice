#Requires -Version 5.1
#Requires -Modules ImportExcel

<#
.SYNOPSIS
    Send data coming from SAP to CreditDevice.

.DESCRIPTION
    SAP generates 2 .TXT files that contains the debtors and the invoices.
    The data in these files uploaded to the CreditDevice server so it can 
    be used to send out reminder mails in case invoices haven't been paid.

    The file created on the day that the script executes is the one that is 
    converted to an Excel file and send to the supplier by mail.

    In case there is no .ASC file created on the day that the script runs, 
    nothing is done and no mail is sent out.

.PARAMETER DebtorFile
    File path containing the debtors

.PARAMETER InvoiceFile
    File path containing the invoices
#>

[CmdLetBinding()]
Param (
    [Parameter(Mandatory)]
    [String]$ScriptName,
    [Parameter(Mandatory)]
    [String]$ImportFile,
    [String]$Token = $env:CREDIT_DEVICE_API_TOKEN,
    [String]$LogFolder = "$env:POWERSHELL_LOG_FOLDER\Application specific\SAP\$ScriptName",
    [String]$ScriptAdmin = $env:POWERSHELL_SCRIPT_ADMIN
)

Begin {
    Function Get-StringHC {
        <# 
        .SYNOPSIS
            Get a part of a string. On error return a blank string.
        #>
        Param (
            [Parameter(Mandatory)]
            [String]$String,
            [Parameter(Mandatory)]
            [Int]$Start,
            [Parameter(Mandatory)]
            [Int]$Length
        )
    
        try {
            $String.SubString($Start, $Length).Trim()
        }
        catch {
            $Error.RemoveAt(0)
    
            $totalLength = $String.length
            
            if ($Start -gt $totalLength) {
                Return ''
            }
            
            $calculatedLength = $totalLength - $Start
            $String.SubString($Start, $calculatedLength).Trim()
        }
    }
    
    Function Send-DataToCreditDeviceHC {
        <# 
        .SYNOPSIS
            Send debtor and invoice date to the REST API of CreditDevice.

        .PARAMETER Token
            Valid token issued by CreditDevice to talk to their API.

        .PARAMETER Type
            The type of the definition known by CreditDevice that maps the 
            data sent to their internal data structure.

        .PARAMETER Data
            Contains the data that needs to be send to the CreditDevice API.
        
        .PARAMETER MaxUploadsAtOnce
            Used to split-up the data sent to the CreditDevice API so we don't 
            spam it with one big chunk of data all at once.
        
        .PARAMETER ThrottleLimit
            The maximum of uploads we execute at the same time.

        .PARAMETER Timeout
            The time to wait for the CreditDevice API to return a response of 
            success or failure after the upload.
        #>

        [CmdLetBinding()]
        Param (
            [Parameter(Mandatory)]
            [String]$Token,
            [Parameter(Mandatory)]
            [ValidateSet('Debtor', 'Invoice')]
            [String]$Type,
            [Parameter(Mandatory)]
            [PSCustomObject[]]$Data,
            [Int]$MaxUploadsAtOnce = 4000,
            [Int]$ThrottleLimit = 4,
            [TimeSpan]$Timeout = (New-TimeSpan -Minutes 45)
        )

        try {
            $definitionID = switch ($Type) {
                'Debtor' { 1 ; break }
                'Invoice' { 2 ; break }
                Default { throw "Definition type '$Type' not implemented" }
            }

            #region Test token valid
            try {
                $M = 'Test token valid'
                Write-Verbose $M

                $params = @{
                    Method  = 'GET'
                    Uri     = 'https://api.directdevice.info/dam/clientinfo'
                    Headers = @{Authorization = "Bearer $Token" }
                    Verbose = $false
                }
                if (-not ($clientInfo = Invoke-RestMethod @params)) {
                    throw 'No client information found for token'
                }
            }
            catch {
                throw "Failed authenticating to uri '$($params.Uri)' with token '$Token': $_"
            }
            #endregion

            #region Test import definition
            try {
                $M = "Test import type '$Type' with definition ID '$definitionID'"
                Write-Verbose $M

                $params = @{
                    Method  = 'GET'
                    Uri     = 'https://api.directdevice.info/imm/imports/definitions'
                    Headers = @{Authorization = "Bearer $Token" }
                    Verbose = $false
                }
                $importDefinitions = Invoke-RestMethod @params
    
                if (-not ($importDefinitions.data)) {
                    throw 'No import definitions found'
                }
                if ($importDefinitions.data.id -notContains $definitionID) {
                    throw "Definition number '$definitionID' not found in the list of known import definitions '$($importDefinitions.data.id)' by CreditDevice"
                }
            }
            catch {
                throw "Failed to get the definitions: $_"
            }
            #endregion

            $scriptBlock = {
                Param (
                    [Parameter(Mandatory)]
                    [String]$Token,
                    [Parameter(Mandatory)]
                    [TimeSpan]$Timeout,
                    [Parameter(Mandatory)]
                    [String]$Type,
                    [Parameter(Mandatory)]
                    [Int]$DefinitionID,
                    [Parameter(Mandatory)]
                    [PSCustomObject[]]$DataChunk
                )

                $importTransaction = @{
                    id     = 0
                    result = $null
                }

                # $VerbosePreference = 'Continue'

                #region Get import transaction id
                try {
                    $M = "Get import transaction id for definition '$DefinitionID'"
                    Write-Verbose $M

                    $params = @{
                        Method      = 'POST'
                        Uri         = 'https://api.directdevice.info/imm/imports'
                        ContentType = 'application/json'
                        Headers     = @{
                            Authorization = "Bearer $Token" 
                            Accept        = 'application/json'
                        }
                        # UseBasicParsing = $true
                        Body        = [System.Text.Encoding]::UTF8.GetBytes((
                                @{
                                    id = $DefinitionID
                                } | ConvertTo-Json
                            )
                        )
                        Verbose     = $false
                    }
                    $importTransaction.id = (Invoke-RestMethod @params).id
    
                    if (-not $importTransaction.id) {
                        throw 'No transaction id received'
                    }
                }
                catch {
                    throw "Failed to create an import transaction for type '$Type' with definition number '$definitionID': $_"
                }
                #endregion
        
                #region Start import
                try {
                    $M = "Transaction ID '$($importTransaction.id)' : Start import"
                    Write-Verbose $M

                    $params = @{
                        Method      = 'PUT'
                        Uri         = 'https://api.directdevice.info/imm/imports/{0}/contents' -f $importTransaction.id
                        ContentType = 'application/json'
                        Headers     = @{
                            Authorization = "Bearer $Token" 
                            Accept        = 'application/json'
                        }
                        Body        = [System.Text.Encoding]::UTF8.GetBytes((
                                $DataChunk | ConvertTo-Json
                            )
                        )
                        Verbose     = $false
                    }
                    $importTransaction.result = Invoke-RestMethod @params
                }
                catch {
                    throw "Failed to start the import for type '$Type' with  definition number '$definitionID' and transaction id '$($importTransaction.id)': $_"
                }
                #endregion

                #region Check import status
                try {
                    $params = @{
                        Method  = 'GET'
                        Uri     = 'https://api.directdevice.info/imm/imports/{0}' -f $importTransaction.id
                        Headers = @{
                            Authorization = "Bearer $Token" 
                            Accept        = 'application/json'
                        }
                        Verbose = $false
                    }
                
                    $timer = [Diagnostics.Stopwatch]::StartNew()

                    while (
                    ($importTransaction.result.status -notMatch '^finished$|^failed$') -and 
                    ($timer.elapsed -lt $Timeout)
                    ) {
                        $importTransaction.result = Invoke-RestMethod @params

                        Write-Verbose "Transaction ID '$($importTransaction.id)' : Import status '$($importTransaction.result.status)' progress '$($importTransaction.result.progress)'"
                        Start-Sleep -Seconds 8
                    }

                    if ($importTransaction.result.status -notMatch '^finished$|^failed$') {
                        throw "We stopped waiting for the API to return the result 'Finished' or 'Failed' after '$([Math]::Round($Timeout.TotalMinutes,2))' minutes. Current status is '$($importTransaction.result.status)'"
                    }
                }
                catch {
                    throw "Failed to check for the import status of type '$Type' with transaction id '$($importTransaction.id)' and import definition '$definitionID': $_"
                }
                #endregion
            }

            #region Upload data in batches
            $jobs = @()

            for ($i = 0; $i -lt $Data.Count; $i = $i + $MaxUploadsAtOnce ) {
                $end = if ($MaxUploadsAtOnce -eq 1) {
                    $i
                }
                else {
                ($i + $MaxUploadsAtOnce) - 1
                }
            
                $M = "Batch upload $($Data.Count) records: $i > $end"
                Write-Verbose $M

                $jobParams = @{
                    ScriptBlock  = $scriptBlock 
                    ArgumentList = $Token, $Timeout, $Type, $definitionID, $Data[$i..$end]
                }
                $jobs += Start-Job @jobParams
                Wait-MaxRunningJobsHC -Name $jobs -MaxThreads $ThrottleLimit
            } 
            #endregion

            $jobs | Receive-Job -Wait -AutoRemoveJob

            Write-Verbose 'Import done'
        }
        catch {
            throw "Failed sending data to CreditDevice: $_"
        }
    }

    try {
        Import-EventLogParamsHC -Source $ScriptName
        Write-EventLog @EventStartParams
        Get-ScriptRuntimeHC -Start

        #region Logging
        try {
            $LogParams = @{
                LogFolder    = New-Item -Path $LogFolder -ItemType 'Directory' -Force -ErrorAction 'Stop'
                Name         = $ScriptName
                Date         = 'ScriptStartTime'
                NoFormatting = $true
            }
            $LogFile = New-LogFileNameHC @LogParams
        }
        Catch {
            throw "Failed creating the log folder '$LogFolder': $_"
        }
        #endregion
        
        #region Import .json file
        $M = "Import .json file '$ImportFile'"
        Write-Verbose $M; Write-EventLog @EventOutParams -Message $M
        
        $file = Get-Content $ImportFile -Raw -EA Stop | ConvertFrom-Json
        #endregion
        
        #region Test .json file properties
        #region MailTo
        if (-not $file.MailTo) {
            throw "Input file '$ImportFile': Property 'MailTo' is missing"
        }
        #endregion

        #region DebtorFile
        if (-not $file.DebtorFile) {
            throw "Input file '$ImportFile': Property 'DebtorFile' is missing"
        }
        if (-not (Test-Path -LiteralPath $file.DebtorFile -PathType Leaf)) {
            throw "Input file '$ImportFile': Debtor file '$($file.DebtorFile)' not found"
        }
        #endregion

        #region InvoiceFile
        if (-not $file.InvoiceFile) {
            throw "Input file '$ImportFile': Property 'InvoiceFile' is missing"
        }
        if (-not (Test-Path -LiteralPath $file.InvoiceFile -PathType Leaf)) {
            throw "Input file '$ImportFile': Invoice file '$($file.InvoiceFile)' not found"
        }
        #endregion
        #endregion
    }
    catch {
        Write-Warning $_
        Send-MailHC -To $ScriptAdmin -Subject 'FAILURE' -Priority 'High' -Message $_ -Header $ScriptName
        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"
        Write-EventLog @EventEndParams; Exit 1
    }
}

Process {
    try {
        $excelParams = @{
            Path               = $LogFile + ' - Converted data.xlsx'
            AutoSize           = $true
            FreezeTopRow       = $true
            NoNumberConversion = '*'
        }

        #region Copy source file Debtor.txt to log folder
        $M = "Copy debtor source file '$($file.DebtorFile)' to log folder"
        Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M

        $copyParams = @{
            Path        = $file.DebtorFile
            Destination = $LogFile + ' - Debtor.txt'
            ErrorAction = 'Stop'
        }
        Copy-Item @copyParams
        #endregion

        #region Copy source file Invoice.txt to log folder
        $M = "Copy invoice source file '$($file.InvoiceFile)' to log folder"
        Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M

        $copyParams = @{
            Path        = $file.InvoiceFile
            Destination = $LogFile + ' - Invoice.txt'
            ErrorAction = 'Stop'
        }
        Copy-Item @copyParams
        #endregion

        #region Get debtor and invoice file content
        $fileContent = @{
            debtor  = @{
                raw       = Get-Content -LiteralPath $file.DebtorFile -Encoding UTF8
                converted = [System.Collections.ArrayList]@()
            }
            invoice = @{
                raw       = Get-Content -LiteralPath $file.InvoiceFile -Encoding UTF8
                converted = [System.Collections.ArrayList]@()
            }
        }
    
        $M = "Imported rows: debtor file '{0}' invoice file '{1}'" -f 
        ($fileContent.debtor.raw | Measure-Object).Count,
        ($fileContent.invoice.raw | Measure-Object).Count
        Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M
        #endregion

        #region Convert debtor file to objects
        $M = "Convert file debtor '$($file.DebtorFile)' to objects"
        Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M
            
        foreach (
            $line in 
            $fileContent.debtor.raw
        ) {
            try {
                Write-Verbose "line: $line"
                $params = @{
                    String = $line
                }
                Get-StringHC @params -Start 682 -Length 3
                If ($companyCode = Get-StringHC @params -Start 10 -Length 4) {
                    $creditExposure = if (
                        ($line.length -gt 654) -and
                        ($sapCreditExposure = $line.SubString(654, 15).Trim())
                    ) {
                        $tmp = $sapCreditExposure.Replace('.', ',')
                        if ($tmp[-1] -eq '-') {
                            '-' + $tmp.Substring(0, $tmp.Length - 1)
                        }
                        else { $tmp }
                    }
                    else { '' }
                    
                    $null = $fileContent.debtor.converted.Add(
                        [PSCustomObject]@{
                            DebtorNumber          = Get-StringHC @params -Start 0 -Length 10
                            CompanyCode           = $companyCode
                            Name                  = Get-StringHC @params -Start 14 -Length 35
                            NameExtra             = Get-StringHC @params -Start 129 -Length 35
                            Street                = Get-StringHC @params -Start 49 -Length 35
                            PostalCode            = Get-StringHC @params -Start 84 -Length 10
                            City                  = Get-StringHC @params -Start 94 -Length 35
                            CountryCode           = Get-StringHC @params -Start 230 -Length 3
                            CountryName           = Get-StringHC @params -Start 574 -Length 22
                            PoBox                 = Get-StringHC @params -Start 164 -Length 18
                            PoBoxPostalCode       = Get-StringHC @params -Start 182 -Length 10
                            PoBoxCity             = Get-StringHC @params -Start 192 -Length 35
                            PhoneNumber           = Get-StringHC @params -Start 233 -Length 16
                            MobilePhoneNumber     = Get-StringHC @params -Start 249 -Length 16
                            EmailAddress          = Get-StringHC @params -Start 265 -Length 50
                            Comment               = Get-StringHC @params -Start 533 -Length 36
                            CreditLimit           = Get-StringHC @params -Start 356 -Length 18
                            VatRegistrationNumber = Get-StringHC @params -Start 377 -Length 20
                            AccountGroup          = Get-StringHC @params -Start 442 -Length 4
                            CustomerLanguage      = Get-StringHC @params -Start 446 -Length 1
                            PaymentTerms          = Get-StringHC @params -Start 451 -Length 4
                            DunsNumber            = Get-StringHC @params -Start 596 -Length 11
                            Rating                = Get-StringHC @params -Start 682 -Length 3
                            DbCreditLimit         = Get-StringHC @params -Start 621 -Length 20
                            NextInReview          = Get-StringHC @params -Start 641 -Length 13
                            RiskCategory          = Get-StringHC @params -Start 669 -Length 3
                            CreditAccount         = Get-StringHC @params -Start 674 -Length 8
                            CreditExposure        = $creditExposure
                        }
                    )
                }
            }    
            catch {
                Write-Warning "Failed converting debtor data '$line': $_"
            }
        }
        #endregion

        #region Export debtor file to Excel
        $excelParams.WorksheetName = 'Debtor'
        $excelParams.TableName = 'Debtor'
        $fileContent.debtor.converted | Export-Excel @excelParams

        $M = "Exported '$($fileContent.debtor.converted.Count)' rows to Excel file '$($excelParams.Path)'"
        Write-Verbose $M; Write-EventLog @EventOutParams -Message $M
        #endregion

        #region Send debtor data to CreditDevice
        $sendParams = @{
            Token = $Token
            Type  = 'Debtor'
            Data  = $fileContent.debtor.converted
        }
        Send-DataToCreditDeviceHC @sendParams
        #endregion

        #region Convert invoice file to objects
        $M = "Convert invoice file '$($file.InvoiceFile)' to objects"
        Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M
            
        foreach (
            $line in 
            $fileContent.invoice.raw
        ) {
            try {
                $documentType = $line.SubString(144, 2).Trim()
                $sapDocumentNumber = $line.SubString(14, 10).Trim()
                # $sapDocumentNumber = $line.SubString(174, 12).Trim()
                $reference = if ($line.length -gt 189) { 
                    $line.SubString(189, $line.length - 189).Trim()
                }
                else { '' }
                $invoiceNumber = switch ($documentType) {
                    'RV' { $reference ; break }
                    'DB' { $reference ; break }
                    'DC' { $sapDocumentNumber ; break }
                    'DM' { $sapDocumentNumber ; break }
                    Default { '' }
                }

                $null = $fileContent.invoice.converted.Add(
                    [PSCustomObject]@{
                        SapDocumentNumber = $sapDocumentNumber
                        DebtorNumber      = $line.SubString(0, 10).Trim()
                        CompanyCode       = $line.SubString(10, 4).Trim()
                        BusinessArea      = $line.SubString(136, 4).Trim()
                        DocumentType      = $documentType
                        Reference         = $reference
                        InvoiceNumber     = $invoiceNumber
                        Description       = $line.SubString(75, 50).Trim()
                        DocumentDate      = $line.SubString(35, 8).Trim()
                        NetDueDate        = $line.SubString(125, 8).Trim()
                        Amount            = $line.SubString(44, 16).Trim()
                        Currency          = $line.SubString(133, 3).Trim()
                    }
                )
            }    
            catch {
                Write-Warning "Failed converting invoice data '$line': $_"
            }
        }
        #endregion

        #region Export invoice file to Excel
        $excelParams.WorksheetName = 'Invoice'
        $excelParams.TableName = 'Invoice'
        $fileContent.invoice.converted | Export-Excel @excelParams

        $M = "Exported '$($fileContent.invoice.converted.Count)' rows to Excel file '$($excelParams.Path)'"
        Write-Verbose $M; Write-EventLog @EventOutParams -Message $M
        #endregion

        #region Send invoice data to CreditDevice
        $sendParams = @{
            Token = $Token
            Type  = 'Invoice'
            Data  = $fileContent.invoice.converted
        }
        Send-DataToCreditDeviceHC @sendParams
        #endregion
    }
    catch {
        Write-Warning $_
        Send-MailHC -To $ScriptAdmin -Subject 'FAILURE' -Priority 'High' -Message $_ -Header $ScriptName
        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"
        Write-EventLog @EventEndParams; Exit 1
    }
}

End {
    try { 
        $mailParams = @{
            To          = $file.MailTo
            Bcc         = $ScriptAdmin
            Subject     = '{0} invoices, {1} debtors' -f 
            $fileContent.invoice.converted.Count,
            $fileContent.debtor.converted.Count
            Attachments = $excelParams.Path
            LogFolder   = $LogParams.LogFolder
            Header      = $ScriptName
            Save        = $LogFile + ' - Mail.html'
            Priority    = 'Normal'
        }

        #region Set mail subject and priority
        $systemErrors = $Error.Exception.Message | 
        Where-Object { $_ } | Get-Unique
    
        if ($systemErrors) {
            $mailParams.Subject = "{0} error{1}, {2}" -f 
            $systemErrors.Count, $(if ($systemErrors.Count -ge 2) { 's' }), 
            $mailParams.Subject 
            $mailParams.Priority = 'High'
        }
        #endregion
    
        #region Create system errors HTML list
        $htmlSystemErrorsList = $null
    
        if ($systemErrors) {
            $systemErrors | ForEach-Object {
                Write-EventLog @EventErrorParams -Message $_
            }
    
            $htmlSystemErrorsList = $systemErrors | 
            ConvertTo-HtmlListHC -Spacing Wide -Header 'System errors:'
        }
        #endregion
    
        #region Send mail to end user
        $mailParams.Message = "
        $htmlSystemErrorsList
        <p>First test on converting data</p>
            <p><i>* Check the attachment for details</i></p>
        "

        Get-ScriptRuntimeHC -Stop
        Send-MailHC @mailParams
        #endregion
    }
    catch {
        Write-Warning $_
        Send-MailHC -To $ScriptAdmin -Subject 'FAILURE' -Priority 'High' -Message $_ -Header $ScriptName
        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"
        Exit 1
    }
    Finally {
        Write-EventLog @EventEndParams
    }
}