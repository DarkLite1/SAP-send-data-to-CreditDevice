#Requires -Modules Pester
#Requires -Version 5.1

BeforeAll {
    $testOutParams = @{
        FilePath = (New-Item "TestDrive:/params.json" -ItemType File).FullName
        Encoding = 'utf8'
    }

    $testImportFile = @{
        MailTo      = 'bob@contoso.com'
        DebtorFile  = (New-Item "TestDrive:/deb.txt" -ItemType File).FullName
        InvoiceFile = (New-Item "TestDrive:/inv.txt" -ItemType File).FullName
    }
    $testImportFile | ConvertTo-Json -Depth 5 | Out-File @testOutParams
    
    $testScript = $PSCommandPath.Replace('.Tests.ps1', '.ps1')
    $testParams = @{
        ScriptName = 'Test (Brecht)'
        ImportFile = $testOutParams.FilePath
        LogFolder  = New-Item 'TestDrive:/log' -ItemType Directory
    }

    Mock Send-MailHC
    Mock Write-EventLog
}
Describe 'the mandatory parameters are' {
    It '<_>' -ForEach @('ImportFile', 'ScriptName') {
        (Get-Command $testScript).Parameters[$_].Attributes.Mandatory | 
        Should -BeTrue
    }
}
Describe 'send an e-mail to the admin when' {
    BeforeAll {
        $MailAdminParams = {
            ($To -eq $ScriptAdmin) -and ($Priority -eq 'High') -and 
            ($Subject -eq 'FAILURE')
        }    
    }
    AfterAll {
        $testImportFile | ConvertTo-Json -Depth 5 | Out-File @testOutParams
    }
    It 'the log folder cannot be created' {
        $testNewParams = $testParams.clone()
        $testNewParams.LogFolder = 'xxx:://notExistingLocation'

        .$testScript @testNewParams

        Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
            (&$MailAdminParams) -and 
            ($Message -like '*Failed creating the log folder*')
        }
    }
    Context 'the ImportFile' {
        It 'is not found' {
            $testNewParams = $testParams.clone()
            $testNewParams.ImportFile = 'nonExisting.json'
    
            .$testScript @testNewParams
    
            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailAdminParams) -and ($Message -like "Cannot find path*nonExisting.json*")
            }
            Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                $EntryType -eq 'Error'
            }
        }
        Context 'property' {
            It 'MailTo is missing' {
                $testNewImportFile = $testImportFile.Clone()
                $testNewImportFile.MailTo = $null
                $testNewImportFile | ConvertTo-Json | Out-File @testOutParams
                
                .$testScript @testParams
                
                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailAdminParams) -and ($Message -like "*$ImportFile*Property 'MailTo' is missing*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            }
            It 'DebtorFile is missing' {
                $testNewImportFile = $testImportFile.Clone()
                $testNewImportFile.DebtorFile = $null
                $testNewImportFile | ConvertTo-Json | Out-File @testOutParams
                
                .$testScript @testParams
                
                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailAdminParams) -and ($Message -like "*$ImportFile*Property 'DebtorFile' is missing*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            }
            It 'DebtorFile path does not exist' {
                $testNewImportFile = $testImportFile.Clone()
                $testNewImportFile.DebtorFile = 'TestDrive:/NotExisting.txt'
                $testNewImportFile | ConvertTo-Json | Out-File @testOutParams
                
                .$testScript @testParams
                
                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailAdminParams) -and ($Message -like "*$ImportFile*Debtor file '$($testNewImportFile.DebtorFile)' not found*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            }
            It 'InvoiceFile is missing' {
                $testNewImportFile = $testImportFile.Clone()
                $testNewImportFile.InvoiceFile = $null
                $testNewImportFile | ConvertTo-Json | Out-File @testOutParams
                
                .$testScript @testParams
                
                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailAdminParams) -and ($Message -like "*$ImportFile*Property 'InvoiceFile' is missing*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            }
            It 'InvoiceFile path does not exist' {
                $testNewImportFile = $testImportFile.Clone()
                $testNewImportFile.InvoiceFile = 'TestDrive:/NotExisting.txt'
                $testNewImportFile | ConvertTo-Json | Out-File @testOutParams
                
                .$testScript @testParams
                
                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailAdminParams) -and ($Message -like "*$ImportFile*Invoice file '$($testNewImportFile.InvoiceFile)' not found*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            }
        }
    }
}
Describe 'when all tests pass' {
    BeforeAll {
        $testDate = @{
            Debtor  = @"
        0021999981BE99THE FEDERATION                     STAR FLEET STREET 9                1234      GALAXY                             JEAN-LUC PICARD                                                                                   01 GA 0345 12 12 12                   piacard@starfleet.com                                                            PICARDJL                   2EURBE0xxxx4x688                                           0021920781SF01F EURZBAR                                                                     BE14     EI                                       WORLD                                                          0,0020210101             3186.31C011111111181
        0029843423    CONTOSO                            STREET IN REDMOND 1                9999      REDMOND                                                                                                                              02 US 555 43 33 68    0477 11 11 11   info@contoso.com                                                                 CNORRIS                14000   GB5555554536                                           0021920823US01N USD                                                                         BE98        EI /also customer 21111114       M    US                    37-003-0021                           4250,0020210101             7306.50US99999999993F2
"@
            Invoice = @"
            0021111417US106001076236US10999999999999928          752,14          752,142175116578                                        20220430EURUS6YBE10RV2175116578       000000000  6001076236     2175116578
            0021439990GB104900050714GB10246546546610726        -1165,00        -1165,00DP2165769534  15/05/2021  (23/07/2021)            20210726EURGBE6BE10DZ                 000000000  4900050714  RMC
"@
        }
        $testExportedExcelRows = @{
            Debtor  = @(
                @{
                    DebtorNumber       = '0021999981'
                    PlantNumber        = 'BE99'
                    CompanyName        = 'THE FEDERATION'
                    StreetAddress1     = 'STAR FLEET STREET 9'
                    PostalCode         = '1234'
                    City               = 'GALAXY'
                    StreetAddress2     = $null
                    StreetAddress3     = $null
                    StreetAddress4     = $null
                    StreetAddress5     = $null
                    StreetAddress6     = '01'
                    CountryCode        = 'GA'             
                    PhoneNumber1       = '0345 12 12 12'
                    PhoneNumber2       = $null
                    EmailAddress       = 'piacard@starfleet.com'
                    FaxNumber          = $null
                    SearchCode         = 'PICARDJL'
                    URL                = $null
                    CreditLimit        = '2'
                    Currency           = 'EUR'
                    RegistrationNumber = 'BE0xxxx4x688'
                    OriginalCustomer   = '0021920781'
                    AccountGroup       = '0021920781'
                    DeletionFlag       = 'Test1'
                    CustomerLanguage   = 'Test1'
                    CustomerCurrency   = 'Test1'
                    PaymentTerms       = 'Test1'
                    AccountPosition    = 'Test1'
                    Collection         = 'Test1'
                    AccountNumber      = 'Test1'
                    IBAN               = 'Test1'
                    BIC                = 'Test1'
                    LegalEntity        = 'Test1'
                    BkGk               = 'Test1'
                    Comment1           = 'Test1'
                    Comment2           = 'Test1'
                    Comment3           = 'Test1'
                    Comment4           = 'Test1'
                    ParentCompany      = 'Test1'
                    DunningClerk       = 'Test1'
                    AccountClerk       = 'Test1'
                    CountryName        = 'Test1'
                    DunningNumber      = 'Test1'
                    DbCreditLimit      = 'Test1'
                    NextInReview       = 'Test1'
                    CreditExposure     = 'Test1'
                    RiskCategory       = 'Test1'
                    CreditAccount      = 'Test1'
                    Rating             = 'Test1'
                }
            )
            Invoice = @(
                @{
                    DebtorNumber      = '0021999981BE99'
                    InvoiceNumber     = 'KROMMENI'
                    InvoiceDate       = 2105880255
                    InvoiceDueDate    = 'Rosariumlaan 47'
                    InvoiceAmount     = 2104737363
                    OutstandingAmount = 22016630
                    Description       = 'Faber W Krommenie'
                    Currency          = 'Rosariumlaan 47'
                    BusinessArea      = 'Rosariumlaan 47'
                    CompanyCode       = 'Rosariumlaan 47'
                    DocumentType      = 'Rosariumlaan 47'
                    DunningBlock      = 103464
                    BusinessLine      = 'CEM I 42,5 N BULK'
                    Reference         = 29.700
                }
            )
        }

        $testDate.Debtor | Out-File -FilePath $testImportFile.DebtorFile
        $testDate.Invoice | Out-File -FilePath $testImportFile.InvoiceFile

        $testMail = @{
            From           = 'boss@contoso.com'
            To             = 'bob@contoso.com'
            Bcc            = @('jack@contoso.com', 'mike@contoso.com')
            SentItemsPath  = '\PowerShell\{0} SENT' -f $testParams.ScriptName
            EventLogSource = $testParams.ScriptName
            Subject        = 'Picard, 2 deliveries'
            Body           = "<p>Dear supplier</p><p>Since delivery date <b>15/03/2022</b> there have been <b>2 deliveries</b>.</p><p><i>* Check the attachment for details</i></p>*"
        }
        
        .$testScript @testParams
    }
    Context 'copy source data to log folder' {
        It 'Debtor.txt' {
            Get-ChildItem $testParams.LogFolder -File -Recurse -Filter '* - Debtor.txt' | Should -Not -BeNullOrEmpty
        }
        It 'Invoice.txt' {
            Get-ChildItem $testParams.LogFolder -File -Recurse -Filter '* - Invoice.txt' | Should -Not -BeNullOrEmpty
        }
    } -Tag test
    Context 'export an Excel file' {
        BeforeAll {
            $testExcelLogFile = Get-ChildItem $testParams.LogFolder -File -Recurse -Filter '* - Picard - Summary.xlsx'

            $actual = Import-Excel -Path $testExcelLogFile.FullName -WorksheetName 'Data'
        }
        It 'to the log folder' {
            $testExcelLogFile | Should -Not -BeNullOrEmpty
        }
        It 'with the correct total rows' {
            $actual | Should -HaveCount $testExportedExcelRows.Count
        }
        It 'with the correct data in the rows' {
            foreach ($testRow in $testExportedExcelRows) {
                $actualRow = $actual | Where-Object {
                    $_.ShipmentNumber -eq $testRow.ShipmentNumber
                }
                @(
                    'Plant', 'DeliveryNumber', 'ShipToNumber', 'ShipToName',
                    'Address', 'City', 'MaterialNumber', 'MaterialDescription',
                    'Tonnage', 'LoadingDate', 'TruckID', 'PickingStatus', 
                    'SiloBulkID', 'File'
                ) | ForEach-Object {
                    $actualRow.$_ | Should -Be $testRow.$_
                }
            }
        }
    }
    It 'create a sent items folder in the mailbox' {
        Should -Invoke New-MailboxFolderHC -Exactly 1 -Scope Describe 
    }
    It 'send a summary mail to the user' {
        Should -Invoke Send-MailAuthenticatedHC -Exactly 1 -Scope Describe -ParameterFilter {
            ($From -eq $testMail.From) -and
            ($To -eq $testMail.To) -and
            ($Bcc -contains $ScriptAdmin) -and
            ($Bcc -contains $testMail.Bcc[0]) -and
            ($Bcc -contains $testMail.Bcc[1]) -and
            ($SentItemsPath -eq $testMail.SentItemsPath) -and
            ($EventLogSource -eq $testMail.EventLogSource) -and
            ($Subject -eq $testMail.Subject) -and
            ($Attachments.Count -eq 1) -and
            ($Attachments[0] -like '* - Picard - Summary.xlsx') -and
            # ($Attachments[0] -Like '* - Picard - Test1.asc') -and
            ($Body -like $testMail.Body)
        }
    }
}