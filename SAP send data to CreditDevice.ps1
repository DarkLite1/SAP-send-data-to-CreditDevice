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

        $mailParams = @{
            To          = $file.MailTo
            Bcc         = $ScriptAdmin
            Attachments = @($excelParams.Path)
            LogFolder   = $LogParams.LogFolder
            Header      = $ScriptName
            Save        = $LogFile + ' - Mail.html'
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

        $mailParams.Attachments += $copyParams.Destination
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

        $mailParams.Attachments += $copyParams.Destination
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

        #region Send mail to end user
        $mailParams += @{
            Subject = '{0} invoices, {1} debtors' -f 
            ($fileContent.invoice.converted | Measure-Object).Count,
            ($fileContent.debtor.converted | Measure-Object).Count
            Message =
            "<p>First test on converting data</p>
                <p><i>* Check the attachment for details</i></p>"
        }
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