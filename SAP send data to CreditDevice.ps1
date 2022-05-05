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
                raw       = Get-Content -LiteralPath $file.DebtorFile
                converted = @()
            }
            invoice = @{
                raw       = Get-Content -LiteralPath $file.InvoiceFile
                converted = @()
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
            
        $fileContent.debtor.converted += foreach (
            $line in 
            $fileContent.debtor.raw
        ) {
            try {
                [PSCustomObject]@{
                    DebtorNumber       = $line.SubString(0, 10).Trim()
                    PlantNumber        = $line.SubString(10, 4).Trim()
                    CompanyName        = $line.SubString(14, 35).Trim()
                    StreetAddress1     = $line.SubString(49, 35).Trim()
                    PostalCode         = $line.SubString(84, 10).Trim()
                    City               = $line.SubString(94, 35).Trim()
                    StreetAddress2     = $line.SubString(129, 35).Trim()
                    StreetAddress3     = $line.SubString(164, 18).Trim()
                    StreetAddress4     = $line.SubString(182, 10).Trim()
                    StreetAddress5     = $line.SubString(192, 35).Trim()
                    StreetAddress6     = $line.SubString(227, 3).Trim()
                    CountryCode        = $line.SubString(230, 3).Trim()             
                    PhoneNumber1       = $line.SubString(233, 16).Trim()
                    PhoneNumber2       = $line.SubString(249, 16).Trim()
                    EmailAddress       = $line.SubString(265, 50).Trim()
                    FaxNumber          = $line.SubString(315, 31).Trim()
                    SearchCode         = $line.SubString(346, 10).Trim()
                    CreditLimit        = $line.SubString(356, 18).Trim()
                    Currency           = $line.SubString(374, 3).Trim()
                    RegistrationNumber = $line.SubString(377, 20).Trim()
                    URL                = $line.SubString(397, 35).Trim()
                    OriginalCustomer   = $line.SubString(432, 10).Trim()
                    AccountGroup       = $line.SubString(442, 4).Trim()
                    CustomerLanguage   = $line.SubString(446, 1).Trim()
                    DeletionFlag       = $line.SubString(447, 1).Trim()
                    CustomerCurrency   = $line.SubString(448, 3).Trim()
                    PaymentTerms       = $line.SubString(451, 4).Trim()
                    AccountPosition    = $line.SubString(455, 2).Trim()
                    Collection         = $line.SubString(457, 4).Trim()
                    AccountNumber      = $line.SubString(461, 18).Trim()
                    IBAN               = $line.SubString(479, 34).Trim()
                    BIC                = $line.SubString(513, 11).Trim()
                    LegalEntity        = $line.SubString(524, 4).Trim()
                    BkGk               = $line.SubString(528, 2).Trim()
                    Comment1           = $line.SubString(530, 3).Trim()
                    Comment2           = $line.SubString(533, 3).Trim()
                    Comment3           = $line.SubString(536, 3).Trim()
                    Comment4           = $line.SubString(539, 30).Trim()
                    ParentCompany      = $line.SubString(569, 1).Trim()
                    DunningClerk       = $line.SubString(570, 2).Trim()
                    AccountClerk       = $line.SubString(572, 2).Trim()
                    CountryName        = $line.SubString(574, 22).Trim()
                    DunningNumber      = $line.SubString(596, 25).Trim()
                    DbCreditLimit      = $line.SubString(621, 20).Trim()
                    NextInReview       = $line.SubString(641, 13).Trim()
                    CreditExposure     = $line.SubString(654, 14).Trim()
                    RiskCategory       = $line.SubString(668, 3).Trim()
                    CreditAccount      = $line.SubString(671, 8).Trim()
                    Rating             = $line.SubString(679, 2).Trim()
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

        $M = "Exported '$(($fileContent.debtor.converted | Measure-Object).Count)' rows to Excel file '$($excelParams.Path)'"
        Write-Verbose $M; Write-EventLog @EventOutParams -Message $M
        #endregion

        #region Convert invoice file to objects
        $M = "Convert invoice file '$($file.InvoiceFile)' to objects"
        Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M
            
        $fileContent.invoice.converted += foreach (
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

        $M = "Exported '$(($fileContent.invoice.converted | Measure-Object).Count)' rows to Excel file '$($excelParams.Path)'"
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
        # Send-MailHC @mailParams
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