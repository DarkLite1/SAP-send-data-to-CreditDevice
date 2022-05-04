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
        $mailParams = @{
            To        = $file.MailTo
            Bcc       = $ScriptAdmin
            LogFolder = $LogParams.LogFolder
            Header    = $ScriptName
            Save      = $LogFile + ' - Mail.html'
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
            [PSCustomObject]@{
                DebtorNumber       = $line.SubString(0, 10).Trim()
                PlantNumber        = $line.SubString(10, 4).Trim()
                CompanyName        = $line.SubString(14, 36).Trim()
                StreetAddress1     = $line.SubString(50, 35).Trim()
                PostalCode         = $line.SubString(85, 10).Trim()
                City               = $line.SubString(95, 35).Trim()
                StreetAddress2     = $line.SubString(130, 35).Trim()
                StreetAddress3     = $line.SubString(165, 18).Trim()
                StreetAddress4     = $line.SubString(183, 10).Trim()
                StreetAddress5     = $line.SubString(193, 35).Trim()
                StreetAddress6     = $line.SubString(228, 3).Trim()
                CountryCode        = $line.SubString(231, 3).Trim()             
                PhoneNumber1       = $line.SubString(234, 16).Trim()
                PhoneNumber2       = $line.SubString(250, 16).Trim()
                EmailAddress       = $line.SubString(266, 50).Trim()
                FaxNumber          = $line.SubString(316, 31).Trim()
                SearchCode         = $line.SubString(347, 10).Trim()
                CreditLimit        = $line.SubString(357, 18).Trim()
                Currency           = $line.SubString(375, 3).Trim()
                RegistrationNumber = $line.SubString(378, 20).Trim()
                URL                = $line.SubString(398, 35).Trim()
                OriginalCustomer   = $line.SubString(433, 10).Trim()
                AccountGroup       = $line.SubString(443, 4).Trim()
                CustomerLanguage   = $line.SubString(447, 1).Trim()
                DeletionFlag       = $line.SubString(448, 1).Trim()
                CustomerCurrency   = $line.SubString(449, 3).Trim()
                PaymentTerms       = $line.SubString(452, 4).Trim()
                AccountPosition    = $line.SubString(456, 2).Trim()
                Collection         = $line.SubString(458, 4).Trim()
                AccountNumber      = $line.SubString(462, 18).Trim()
                IBAN               = $line.SubString(480, 34).Trim()
                BIC                = $line.SubString(514, 11).Trim()
                LegalEntity        = $line.SubString(525, 4).Trim()
                BkGk               = $line.SubString(529, 2).Trim()
                Comment1           = $line.SubString(531, 3).Trim()
                Comment2           = $line.SubString(534, 3).Trim()
                Comment3           = $line.SubString(537, 3).Trim()
                Comment4           = $line.SubString(540, 30).Trim()
                ParentCompany      = $line.SubString(570, 1).Trim()
                DunningClerk       = $line.SubString(571, 2).Trim()
                AccountClerk       = $line.SubString(573, 2).Trim()
                CountryName        = $line.SubString(575, 22).Trim()
                DunningNumber      = $line.SubString(597, 25).Trim()
                DbCreditLimit      = $line.SubString(622, 20).Trim()
                NextInReview       = $line.SubString(642, 13).Trim()
                CreditExposure     = $line.SubString(655, 15).Trim()
                RiskCategory       = $line.SubString(670, 3).Trim()
                CreditAccount      = $line.SubString(673, 10).Trim()
                Rating             = $line.SubString(683, 2).Trim()
            }
        }
        #endregion

        #region Export debtor file to Excel
        if ($fileContent.debtor.converted) {
            $excelParams = @{
                Path               = $LogFile + ' - Debtor.xlsx'
                AutoSize           = $true
                WorksheetName      = 'Debtor'
                TableName          = 'Debtor'
                FreezeTopRow       = $true
                NoNumberConversion = '*'
            }
            $fileContent.debtor.converted | Export-Excel @excelParams

            $M = "Exported '$(($fileContent.debtor.converted | Measure-Object).Count)' rows to Excel file '$($excelParams.Path)'"
            Write-Verbose $M; Write-EventLog @EventOutParams -Message $M
            
            [Array]$mailParams.Attachments += $excelParams.Path
        }
        #endregion

        #region Convert invoice file to objects
        $M = "Convert invoice file '$($file.InvoiceFile)' to objects"
        Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M
            
        $fileContent.invoice.converted += foreach (
            $line in 
            $fileContent.invoice.raw
        ) {
            [PSCustomObject]@{
                DebtorNumber      = $line.SubString(0, 14).Trim()
                InvoiceNumber     = $line.SubString(36, 8).Trim()
                InvoiceDate       = $line.SubString(44, 16).Trim()
                InvoiceDueDate    = $line.SubString(60, 16).Trim()
                InvoiceAmount     = $line.SubString(76, 50).Trim()
                OutstandingAmount = $line.SubString(126, 8).Trim()
                Description       = $line.SubString(134, 3).Trim()
                Currency          = $line.SubString(137, 4).Trim()
                BusinessArea      = $line.SubString(141, 4).Trim()
                CompanyCode       = $line.SubString(145, 2).Trim()
                DocumentType      = $line.SubString(147, 16).Trim()
                DunningBlock      = $line.SubString(163, 1).Trim()             
                BusinessLine      = $line.SubString(187, 3).Trim()
                Reference         = $line.SubString(190, 10).Trim()
            }
        }
        #endregion

        #region Export invoice file to Excel
        if ($fileContent.invoice.converted) {
            $excelParams = @{
                Path               = $LogFile + ' - Invoice.xlsx'
                AutoSize           = $true
                WorksheetName      = 'Invoice'
                TableName          = 'Invoice'
                FreezeTopRow       = $true
                NoNumberConversion = '*'
            }
            $fileContent.invoice.converted | Export-Excel @excelParams

            $M = "Exported '$(($fileContent.invoice.converted | Measure-Object).Count)' rows to Excel file '$($excelParams.Path)'"
            Write-Verbose $M; Write-EventLog @EventOutParams -Message $M

            [Array]$mailParams.Attachments += $excelParams.Path
        }
        #endregion

        #region Send mail to end user
        $mailParams += @{
            Subject = 'Upload'
            Message =
            "<p>Folder <a href=""$($InvoicesFolderItem.FullName)"">$($InvoicesFolderItem.Name)</a>:</p>
                $table
                $ErrorTable
                <p>Files are only removed when their file name (InvoiceReference) is no longer registered as 'unpaid' in the Onguard database and their creation time is older than ''.</p>
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