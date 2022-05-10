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
        $testData = @{
            Debtor  = @"
0021920631NL30Switch Poeren Realisatie NW-2 VOF  Westkanaaldijk 2                   3542DA    UTRECHT                            Project 3487 Wintrack              POSTBUS 1025      3600 BA   MAARSSEN                           10 NL 302486911                       crediteuren.scno@strukton.com                                                    SWITCH                     0   NL861210736B01                                         0021920631BE01N EUR                                                                         BE14     EMA    Geen E-invoicing/ KEY        M    NEDERLAND             49-341-0464                           5000,0020221104            43386.04NL10021911144O2
0021510989NL30Den Ouden Aannemingsbedrijf B.V.   Hermalen 7                         5481 XX   Schijndel                                                             POSTBUS 12        5480 AA                                      07 NL 735431000                       j.bongers@denoudengroep.com                       735498360                      OUDEN                 105875   NL801764063B01                                         0021510989BE01N EURN060  00010653742800        NL20INGB0653742800                INGBNL2A   BE01  KEY23.RH                                6969NEDERLAND             41-589-4518    510989               200000,0020221104           385574.20NL100215109892A2
0021403350BE10DE BOEVER PETER BVBA               RIJKSWEG 66  A                     9870      MACHELEN - ZULTE                                                                                                                     08 BE 09/3801923      0475/648839     info@deboeverbvba.be                              09/3801923                     DEBOEVERPE              5000   BE0439876885                                           0021403350BE01N EURN030      390-0458705-47    BE22390045870547                  BBRUBEBB   BE14     EI    facturen per post vanaf 2022 -     BELGIE                50-561-0055    0001011245            12250,0020200101            1344.17-BE20021403350D2
0021920631    Switch Poeren Realisatie NW-2 VOF  Westkanaaldijk 2                   3542DA    UTRECHT                            Project 3487 Wintrack              POSTBUS 1025      3600 BA   MAARSSEN                           10 NL 302486911                       crediteuren.scno@strukton.com                                                    SWITCH                     0   NL861210736B01                                         0021920631BE01N EUR                                                                         BE14     EMA    Geen E-invoicing/ KEY        M    NEDERLAND             49-341-0464                           5000,0020221104            43386.04NL10021911144O2
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
                # skip debtor line without CompanyCode
                @{
                    DebtorNumber          = '0021920631'
                    CompanyCode           = 'NL30'
                    Name                  = 'Switch Poeren Realisatie NW-2 VOF'
                    NameExtra             = 'Project 3487 Wintrack'
                    Street                = 'Westkanaaldijk 2'
                    PostalCode            = '3542DA'
                    City                  = 'UTRECHT'
                    CountryCode           = 'NL'             
                    CountryName           = 'NEDERLAND'
                    PoBox                 = 'POSTBUS 1025'
                    PoBoxPostalCode       = '3600 BA'
                    PoBoxCity             = 'MAARSSEN'
                    PhoneNumber           = '302486911'
                    MobilePhoneNumber     = ''
                    EmailAddress          = 'crediteuren.scno@strukton.com'
                    Comment               = 'EMA    Geen E-invoicing/ KEY'
                    CreditLimit           = '0'
                    VatRegistrationNumber = 'NL861210736B01'
                    AccountGroup          = 'BE01'
                    CustomerLanguage      = 'N'
                    PaymentTerms          = ''
                    DunsNumber            = '49-341-0464'
                    DbCreditLimit         = '5000,00'
                    NextInReview          = '20221104'
                    CreditExposure        = '43386,04'
                    RiskCategory          = 'NL1'
                    CreditAccount         = '21911144'
                    Rating                = 'O2'
                }
                @{
                    DebtorNumber          = '0021510989'
                    CompanyCode           = 'NL30'
                    Name                  = 'Den Ouden Aannemingsbedrijf B.V.'
                    NameExtra             = ''
                    Street                = 'Hermalen 7'
                    PostalCode            = '5481 XX'
                    City                  = 'Schijndel'
                    CountryCode           = 'NL'             
                    CountryName           = 'NEDERLAND'
                    PoBox                 = 'POSTBUS 12'
                    PoBoxPostalCode       = '5480 AA'
                    PoBoxCity             = ''
                    PhoneNumber           = '735431000'
                    MobilePhoneNumber     = ''
                    EmailAddress          = 'j.bongers@denoudengroep.com'
                    Comment               = '23.RH'
                    CreditLimit           = '105875'
                    VatRegistrationNumber = 'NL801764063B01'
                    AccountGroup          = 'BE01'
                    CustomerLanguage      = 'N'
                    PaymentTerms          = 'N060'
                    DunsNumber            = '41-589-4518'
                    Rating                = '2A2'
                    DbCreditLimit         = '200000,00'
                    NextInReview          = '20221104'
                    CreditExposure        = '385574,20' 
                    RiskCategory          = 'NL1'
                    CreditAccount         = '21510989'
                }
                @{
                    DebtorNumber          = '0021403350'
                    CompanyCode           = 'BE10'
                    Name                  = 'DE BOEVER PETER BVBA'
                    NameExtra             = ''
                    Street                = 'RIJKSWEG 66  A'
                    PostalCode            = '9870'
                    City                  = 'MACHELEN - ZULTE'
                    CountryCode           = 'BE'             
                    CountryName           = 'BELGIE'
                    PoBox                 = ''
                    PoBoxPostalCode       = ''
                    PoBoxCity             = ''
                    PhoneNumber           = '09/3801923'
                    MobilePhoneNumber     = '0475/648839'
                    EmailAddress          = 'info@deboeverbvba.be'
                    Comment               = 'EI    facturen per post vanaf 2022 -'
                    CreditLimit           = '5000'
                    VatRegistrationNumber = 'BE0439876885'
                    AccountGroup          = 'BE01'
                    CustomerLanguage      = 'N'
                    PaymentTerms          = 'N030'
                    DunsNumber            = '50-561-0055'
                    Rating                = 'D2'
                    DbCreditLimit         = '12250,00'
                    NextInReview          = '20200101'
                    CreditExposure        = '-1344,17' # convert '1344.17-'
                    RiskCategory          = 'BE2'
                    CreditAccount         = '21403350'
                }
            )
            Invoice = @(
                # documentType RV | DB, InvoiceNumber = Reference
                # documentType DM | DC, InvoiceNumber = SapDocumentNumber
                @{
                    # documentType RV
                    SapDocumentNumber = '6001077828'
                    DebtorNumber      = '0021419307'
                    CompanyCode       = 'BE10'
                    BusinessArea      = 'BEN3'
                    DocumentType      = 'RV'
                    Reference         = '2165881081'
                    InvoiceNumber     = '2165881081'
                    Description       = '2165881081'
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
                    Description       = '2190485876'
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
                    Description       = '5100000446'
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
                    Description       = 'DB2165721342  31/12/2020  (17/01/2021)'
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
                    InvoiceNumber     = '5100000161'
                    Description       = '5100000161'
                    DocumentDate      = '20220331'
                    NetDueDate        = '20220430'
                    Amount            = '14508,49'
                    Currency          = 'EUR'
                }
            )
        }

        $testData.Debtor | Out-File -FilePath $testImportFile.DebtorFile
        $testData.Invoice | Out-File -FilePath $testImportFile.InvoiceFile
        
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
                        $_.DebtorNumber -eq $testRow.DebtorNumber
                    }
                    @(
                        'CompanyCode',
                        'Name', 'NameExtra', 
                        'Street', 'PostalCode',
                        'City', 'CountryCode', 'CountryName',
                        'PoBox', 'PoBoxPostalCode','PoBoxCity', 
                        'MobilePhoneNumber','PhoneNumber', 
                        'EmailAddress', 'Comment',
                        'CreditLimit',
                        'VatRegistrationNumber',
                        'AccountGroup', 'CustomerLanguage',
                        'PaymentTerms', 
                        'DunsNumber', 'DbCreditLimit',
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
            ($Subject -eq '6 invoices, 3 debtors') -and
            ($Attachments.Count -eq 3) -and
            ($Attachments[0] -like '* - Converted data.xlsx') -and
            ($Attachments[1] -like '* - Debtor.txt') -and
            ($Attachments[2] -like '* - Invoice.txt')
            #  -and
            # ($Body -like "<p>Dear supplier</p><p>Since delivery date <b>15/03/2022</b> there have been <b>2 deliveries</b>.</p><p><i>* Check the attachment for details</i></p>*")
        }
    } -tag test
}