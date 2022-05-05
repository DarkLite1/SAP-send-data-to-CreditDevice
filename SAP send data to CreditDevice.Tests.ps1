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
0021419307BE106001077828BE10202200120220228          141,57          141,572165881081                                        20220429EURBEN3BE10RV2165881081       000000000  6001077828  RMC2165881081
0021402383BE106001081024BE10202200120220314        18783,42        18783,422190485876                                        20220430EURBE2DBE10RV2190485876       000000000  6001081024  CEM2190485876
0021657240BE105100000446BE10202100120210831          370,20          370,205100000446                                        20211031EURBE10BE10DMDM-202108-0004   000000000  5100000446  CEM
0021002568BE101100000058BE10202100120211115        -6795,00        -6795,00                                                  20211115EURBE16BE10DC                 000000000  1100000058  CEM
0021614403BE104900006280BE10202100220210129          -84,10          -84,10DB2165721342  31/12/2020  (17/01/2021)            20210129EURBEN4BE10DZ                 000000000  4900006280  RMC
0021626016BE105100000161BE10202200120220331        14508,49        14508,495100000161                                        20220430EURBE6YBE10DM                 000000000  5100000161
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
                    StreetAddress2     = 'JEAN-LUC PICARD'
                    StreetAddress3     = ''
                    StreetAddress4     = ''
                    StreetAddress5     = ''
                    StreetAddress6     = '01'
                    CountryCode        = 'GA'             
                    PhoneNumber1       = '0345 12 12 12'
                    PhoneNumber2       = ''
                    EmailAddress       = 'piacard@starfleet.com'
                    FaxNumber          = ''
                    SearchCode         = 'PICARDJL'
                    CreditLimit        = '2'
                    Currency           = 'EUR'
                    RegistrationNumber = 'BE0xxxx4x688'
                    URL                = ''
                    OriginalCustomer   = '0021920781'
                    AccountGroup       = 'SF01'
                    CustomerLanguage   = 'F'
                    DeletionFlag       = ''
                    CustomerCurrency   = 'EUR'
                    PaymentTerms       = 'ZBAR'
                    AccountPosition    = ''
                    Collection         = ''
                    AccountNumber      = ''
                    IBAN               = ''
                    BIC                = ''
                    LegalEntity        = 'BE14'
                    BkGk               = ''
                    Comment1           = ''
                    Comment2           = 'EI'
                    Comment3           = ''
                    Comment4           = ''
                    ParentCompany      = ''
                    DunningClerk       = ''
                    AccountClerk       = ''
                    CountryName        = 'WORLD'
                    DunningNumber      = ''
                    DbCreditLimit      = '0,00'
                    NextInReview       = '20210101'
                    CreditExposure     = '3186.3'
                    RiskCategory       = '1C0'
                    CreditAccount      = '11111111'
                    Rating             = '18'
                }
                @{
                    DebtorNumber       = '0029843423'
                    PlantNumber        = ''
                    CompanyName        = 'CONTOSO'
                    StreetAddress1     = 'STREET IN REDMOND 1'
                    PostalCode         = '9999'
                    City               = 'REDMOND'
                    StreetAddress2     = ''
                    StreetAddress3     = ''
                    StreetAddress4     = ''
                    StreetAddress5     = ''
                    StreetAddress6     = '02'
                    CountryCode        = 'US'             
                    PhoneNumber1       = '555 43 33 68'
                    PhoneNumber2       = '0477 11 11 11'
                    EmailAddress       = 'info@contoso.com'
                    FaxNumber          = ''
                    SearchCode         = 'CNORRIS'
                    URL                = ''
                    CreditLimit        = '14000'
                    Currency           = ''
                    RegistrationNumber = 'GB5555554536'
                    OriginalCustomer   = '0021920823'
                    AccountGroup       = 'US01'
                    DeletionFlag       = ''
                    CustomerLanguage   = 'N'
                    CustomerCurrency   = 'USD'
                    PaymentTerms       = ''
                    AccountPosition    = ''
                    Collection         = ''
                    AccountNumber      = ''
                    IBAN               = ''
                    BIC                = ''
                    LegalEntity        = 'BE98'
                    BkGk               = ''
                    Comment1           = ''
                    Comment2           = ''
                    Comment3           = 'EI'
                    Comment4           = '/also customer 21111114'
                    ParentCompany      = 'M'
                    DunningClerk       = ''
                    AccountClerk       = ''
                    CountryName        = 'US'
                    DunningNumber      = '37-003-0021'
                    DbCreditLimit      = '4250,00'
                    NextInReview       = '20210101'
                    CreditExposure     = '7306.5'
                    RiskCategory       = '0US'
                    CreditAccount      = '99999999'
                    Rating             = '99'
                }
            )
            Invoice = @(
                @{
                    # documentType RV | DB, InvoiceNumber = Reference
                    # documentType DM | DC, InvoiceNumber = SapDocumentNumber
                    
                    # documentType RV
                    SapDocumentNumber = '6001077828'
                    DebtorNumber      = '0021419307'
                    CompanyCode       = 'BE10'
                    BusinessArea      = 'BEN3'
                    DocumentType      = 'RV'
                    Reference         = '2165881081'
                    InvoiceNumber     = '2165881081'
                    Description       = ''
                    DocumentDate      = '20220228'
                    NetDueDate        = '20220429'
                    Amount            = '141,57'
                    Currency          = 'EUR'
                }
                @{
                    # documentType DB
                    SapDocumentNumber = '6001081024'
                    DebtorNumber      = '0021402383'
                    CompanyCode       = 'BE10'
                    BusinessArea      = 'BE2D'
                    DocumentType      = 'RV'
                    Reference         = '2190485876'
                    InvoiceNumber     = '2190485876'
                    Description       = ''
                    DocumentDate      = '20220314'
                    NetDueDate        = '20220430'
                    Amount            = '18783,42'
                    Currency          = 'EUR'
                }
                @{
                    # documentType DM
                    SapDocumentNumber = '5100000446'
                    DebtorNumber      = '0021657240'
                    CompanyCode       = 'BE10'
                    BusinessArea      = 'BE10'
                    DocumentType      = 'DM'
                    Reference         = ''
                    InvoiceNumber     = '5100000446'
                    Description       = ''
                    DocumentDate      = '20210831'
                    NetDueDate        = '20211031'
                    Amount            = '370,20'
                    Currency          = 'EUR'
                }
                @{
                    # documentType DC
                    SapDocumentNumber = '1100000058'
                    DebtorNumber      = '0021002568'
                    CompanyCode       = 'BE10'
                    BusinessArea      = 'BE16'
                    DocumentType      = 'DC'
                    Reference         = ''
                    InvoiceNumber     = '1100000058'
                    Description       = ''
                    DocumentDate      = '20211115'
                    NetDueDate        = '20211115'
                    Amount            = '-6795,00'
                    Currency          = 'EUR'
                }
                @{
                    # documentType DZ
                    SapDocumentNumber = '4900006280'
                    DebtorNumber      = '0021614403'
                    CompanyCode       = 'BE10'
                    BusinessArea      = 'BEN4'
                    DocumentType      = 'DZ'
                    Reference         = ''
                    InvoiceNumber     = ''
                    Description       = 'DB2165721342  31/12/2020  (17/01/2021}'
                    DocumentDate      = '20210129'
                    NetDueDate        = '20210129'
                    Amount            = '-84,10'
                    Currency          = 'EUR'
                }
                @{
                    # no business line
                    SapDocumentNumber = '5100000161' 
                    DebtorNumber      = '0021626016'
                    CompanyCode       = 'BE10'
                    BusinessArea      = 'BE6Y'
                    DocumentType      = 'DM'
                    Reference         = ''
                    InvoiceNumber     = '4900006280'
                    Description       = '5100000161'
                    DocumentDate      = '20220331'
                    NetDueDate        = '20220430'
                    Amount            = '14508,49'
                    Currency          = 'EUR'
                }
            )
        }

        $testDate.Debtor | Out-File -FilePath $testImportFile.DebtorFile
        $testDate.Invoice | Out-File -FilePath $testImportFile.InvoiceFile
        
        .$testScript @testParams
    }
    Context 'copy source files to log folder' {
        It 'Debtor.txt' {
            Get-ChildItem $testParams.LogFolder -File -Recurse -Filter '* - Debtor.txt' | Should -Not -BeNullOrEmpty
        }
        It 'Invoice.txt' {
            Get-ChildItem $testParams.LogFolder -File -Recurse -Filter '* - Invoice.txt' | Should -Not -BeNullOrEmpty
        }
    }
    Context 'export an Excel file' {
        BeforeAll {
            $testExcelLogFile = Get-ChildItem $testParams.LogFolder -File -Recurse -Filter '* - Converted data.xlsx'
        }
        It 'to the log folder' {
            $testExcelLogFile | Should -Not -BeNullOrEmpty
        }
        Context "with worksheet 'Debtor'" {
            BeforeAll {
                $actual = Import-Excel -Path $testExcelLogFile.FullName -WorksheetName 'Debtor'
            }
            It 'with the correct total rows' {
                $actual | Should -HaveCount $testExportedExcelRows.Debtor.Count
            }
            It 'with the correct data in the rows' {
                foreach ($testRow in $testExportedExcelRows.Debtor) {
                    $actualRow = $actual | Where-Object {
                        $_.SapDocumentNumber -eq $testRow.SapDocumentNumber
                    }
                    @(
                        'DebtorNumber', 'PlantNumber', 'CompanyName', 
                        'StreetAddress1', 'PostalCode', 'City', 
                        'StreetAddress2', 'StreetAddress3',
                        'StreetAddress4', 'StreetAddress5',
                        'StreetAddress6', 'CountryCode',
                        'PhoneNumber1', 'PhoneNumber2',
                        'EmailAddress', 'FaxNumber',
                        'SearchCode', 'CreditLimit',
                        'Currency', 'RegistrationNumber',
                        'URL', 'OriginalCustomer',
                        'AccountGroup', 'CustomerLanguage',
                        'DeletionFlag', 'CustomerCurrency',
                        'PaymentTerms', 'AccountPosition',
                        'Collection', 'AccountNumber',
                        'IBAN', 'BIC',
                        'LegalEntity', 'BkGk',
                        'Comment1', 'Comment2',
                        'Comment3', 'Comment4',
                        'ParentCompany', 'DunningClerk',
                        'AccountClerk', 'CountryName',
                        'DunningNumber', 'DbCreditLimit',
                        'NextInReview', 'CreditExposure',
                        'RiskCategory', 'CreditAccount',
                        'Rating'
                    ) | ForEach-Object {
                        $actualRow.$_ | Should -Be $testRow.$_
                    }
                }
            }
        }
        Context "with worksheet 'Invoice'" {
            BeforeAll {
                $actual = Import-Excel -Path $testExcelLogFile.FullName -WorksheetName 'Invoice'
            }
            It 'with the correct total rows' {
                $actual | Should -HaveCount $testExportedExcelRows.Invoice.Count
            }
            It 'with the correct data in the rows' {
                foreach ($testRow in $testExportedExcelRows.Invoice) {
                    $actualRow = $actual | Where-Object {
                        $_.SapDocumentNumber -eq $testRow.SapDocumentNumber
                    }
                    @(
                        'DebtorNumber',     
                        'CompanyCode',     
                        'InvoiceNumber',    
                        'DocumentDate',     
                        'NetDueDate',
                        'Amount', 
                        'Description',     
                        'Currency',     
                        'BusinessArea',     
                        'DocumentType',     
                        'Reference'
                    ) | ForEach-Object {
                        $actualRow.$_ | Should -Be $testRow.$_
                    }
                }
            }
        }
    }
    It 'send a summary mail to the user' {
        Should -Invoke Send-MailHC -Exactly 1 -Scope Describe -ParameterFilter {
            ($To -eq $testImportFile.MailTo) -and
            ($Bcc -eq $ScriptAdmin) -and
            ($Subject -eq '2 invoices, 2 debtors') -and
            ($Attachments.Count -eq 3) -and
            ($Attachments[0] -like '* - Converted data.xlsx') -and
            ($Attachments[1] -like '* - Debtor.txt') -and
            ($Attachments[2] -like '* - Invoice.txt')
            #  -and
            # ($Body -like "<p>Dear supplier</p><p>Since delivery date <b>15/03/2022</b> there have been <b>2 deliveries</b>.</p><p><i>* Check the attachment for details</i></p>*")
        }
    } -Tag test
}