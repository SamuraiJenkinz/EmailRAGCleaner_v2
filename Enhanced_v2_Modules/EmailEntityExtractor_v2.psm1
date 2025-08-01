# EmailEntityExtractor_v2.psm1 - Advanced Entity Extraction for Email Content
# Enhanced entity recognition optimized for business email content and RAG systems

Export-ModuleMember -Function @(
    'Extract-EmailEntities',
    'Extract-BusinessEntities',
    'Extract-PersonalInformation',
    'Extract-TechnicalEntities',
    'Get-EntityConfidenceScore',
    'Merge-EntityResults'
)

function Extract-EmailEntities {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$Content,
        
        [Parameter(Mandatory=$false)]
        [hashtable]$EmailContext = @{},
        
        [Parameter(Mandatory=$false)]
        [switch]$IncludeBusinessEntities = $true,
        
        [Parameter(Mandatory=$false)]
        [switch]$IncludePersonalInfo = $true,
        
        [Parameter(Mandatory=$false)]
        [switch]$IncludeTechnicalEntities = $true,
        
        [Parameter(Mandatory=$false)]
        [double]$MinConfidenceScore = 0.6
    )
    
    try {
        Write-Verbose "Starting advanced entity extraction for email content..."
        
        if (-not $Content -or $Content.Trim() -eq "") {
            return Get-EmptyEntityResult
        }
        
        $entityResults = @{
            ExtractedAt = Get-Date -Format "yyyy-MM-ddTHH:mm:ss.fffZ"
            ExtractorVersion = "2.0"
            ContentLength = $Content.Length
            ProcessingContext = $EmailContext
            Entities = @{}
        }
        
        # Core entities (always extracted)
        Write-Verbose "Extracting core entities..."
        $coreEntities = Extract-CoreEntities -Content $Content
        $entityResults.Entities.Core = $coreEntities
        
        # Business entities
        if ($IncludeBusinessEntities) {
            Write-Verbose "Extracting business entities..."
            $businessEntities = Extract-BusinessEntities -Content $Content -EmailContext $EmailContext
            $entityResults.Entities.Business = $businessEntities
        }
        
        # Personal information
        if ($IncludePersonalInfo) {
            Write-Verbose "Extracting personal information..."
            $personalInfo = Extract-PersonalInformation -Content $Content
            $entityResults.Entities.Personal = $personalInfo
        }
        
        # Technical entities
        if ($IncludeTechnicalEntities) {
            Write-Verbose "Extracting technical entities..."
            $technicalEntities = Extract-TechnicalEntities -Content $Content
            $entityResults.Entities.Technical = $technicalEntities
        }
        
        # Context-aware entities
        Write-Verbose "Extracting context-aware entities..."
        $contextEntities = Extract-ContextAwareEntities -Content $Content -EmailContext $EmailContext
        $entityResults.Entities.Contextual = $contextEntities
        
        # Filter by confidence score
        if ($MinConfidenceScore -gt 0) {
            $entityResults = Filter-EntitiesByConfidence -EntityResults $entityResults -MinScore $MinConfidenceScore
        }
        
        # Calculate summary statistics
        $entityResults.Summary = Calculate-EntitySummary -EntityResults $entityResults
        
        Write-Verbose "Entity extraction completed. Found $($entityResults.Summary.TotalEntities) entities with $($entityResults.Summary.HighConfidenceEntities) high-confidence matches."
        
        return $entityResults
        
    } catch {
        Write-Error "Entity extraction failed: $($_.Exception.Message)"
        
        return @{
            ExtractedAt = Get-Date -Format "yyyy-MM-ddTHH:mm:ss.fffZ"
            ExtractorVersion = "2.0"
            Error = $_.Exception.Message
            Entities = @{}
            Summary = @{ TotalEntities = 0; ErrorOccurred = $true }
        }
    }
}

function Extract-CoreEntities {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$Content
    )
    
    try {
        $coreEntities = @{
            EmailAddresses = @()
            PhoneNumbers = @()
            URLs = @()
            Dates = @()
            Times = @()
            IPAddresses = @()
            MonetaryAmounts = @()
            Percentages = @()
        }
        
        # Enhanced email pattern with validation
        $emailPattern = '\b[A-Za-z0-9]([A-Za-z0-9._%-]*[A-Za-z0-9])?@[A-Za-z0-9]([A-Za-z0-9.-]*[A-Za-z0-9])?\.[A-Za-z]{2,}\b'
        $emailMatches = [regex]::Matches($Content, $emailPattern)
        foreach ($match in $emailMatches) {
            $email = $match.Value.ToLower()
            $domain = $email.Split('@')[1]
            $coreEntities.EmailAddresses += @{
                Value = $email
                Domain = $domain
                IsBusinessDomain = Test-BusinessDomain -Domain $domain
                Position = $match.Index
                ConfidenceScore = 0.95
            }
        }
        
        # Enhanced phone number patterns
        $phonePatterns = @(
            '\+?1[-.\s]?\(?[0-9]{3}\)?[-.\s]?[0-9]{3}[-.\s]?[0-9]{4}\b',  # US format
            '\b\(?[0-9]{3}\)?[-.\s]?[0-9]{3}[-.\s]?[0-9]{4}\b',           # Standard format
            '\+[0-9]{1,4}[-.\s]?[0-9]{1,4}[-.\s]?[0-9]{1,4}[-.\s]?[0-9]{1,9}', # International
            '\b[0-9]{3}[-.\s][0-9]{3}[-.\s][0-9]{4}\b'                    # Simple format
        )
        
        foreach ($pattern in $phonePatterns) {
            $phoneMatches = [regex]::Matches($Content, $pattern)
            foreach ($match in $phoneMatches) {
                $phone = $match.Value
                $normalizedPhone = $phone -replace '[^\d+]', ''
                if ($normalizedPhone.Length -ge 10) {
                    $coreEntities.PhoneNumbers += @{
                        Value = $phone
                        Normalized = $normalizedPhone
                        Format = Get-PhoneFormat -Phone $phone
                        Position = $match.Index
                        ConfidenceScore = 0.90
                    }
                }
            }
        }
        
        # Enhanced URL patterns
        $urlPatterns = @(
            'https?://[^\s<>"{}|\\^`\[\]]+',
            'www\.[a-zA-Z0-9][a-zA-Z0-9-]*[a-zA-Z0-9]*\.[a-zA-Z]{2,}[^\s<>"{}|\\^`\[\]]*',
            'ftp://[^\s<>"{}|\\^`\[\]]+'
        )
        
        foreach ($pattern in $urlPatterns) {
            $urlMatches = [regex]::Matches($Content, $pattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
            foreach ($match in $urlMatches) {
                $url = $match.Value
                $coreEntities.URLs += @{
                    Value = $url
                    Domain = Extract-DomainFromURL -URL $url
                    Protocol = Extract-ProtocolFromURL -URL $url
                    IsSecure = $url.StartsWith('https://')
                    Position = $match.Index
                    ConfidenceScore = 0.95
                }
            }
        }
        
        # Advanced date patterns
        $datePatterns = @(
            '\b(?:January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{1,2},?\s+\d{4}\b',
            '\b(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\.?\s+\d{1,2},?\s+\d{4}\b',
            '\b\d{1,2}[/-]\d{1,2}[/-]\d{2,4}\b',
            '\b\d{4}[/-]\d{1,2}[/-]\d{1,2}\b',
            '\b\d{1,2}\s+(?:January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{4}\b'
        )
        
        foreach ($pattern in $datePatterns) {
            $dateMatches = [regex]::Matches($Content, $pattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
            foreach ($match in $dateMatches) {
                $date = $match.Value
                $parsedDate = Parse-DateString -DateString $date
                $coreEntities.Dates += @{
                    Value = $date
                    ParsedDate = $parsedDate.ParsedDate
                    Format = $parsedDate.Format
                    IsValid = $parsedDate.IsValid
                    Position = $match.Index
                    ConfidenceScore = $(if ($parsedDate.IsValid) { 0.90 } else { 0.60 })
                }
            }
        }
        
        # Time patterns
        $timePatterns = @(
            '\b\d{1,2}:\d{2}(?::\d{2})?\s*(?:AM|PM)\b',
            '\b\d{1,2}:\d{2}(?::\d{2})?\b',
            '\b(?:1[0-2]|0?[1-9]):\d{2}\s*(?:AM|PM)\b'
        )
        
        foreach ($pattern in $timePatterns) {
            $timeMatches = [regex]::Matches($Content, $pattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
            foreach ($match in $timeMatches) {
                $time = $match.Value
                $coreEntities.Times += @{
                    Value = $time
                    Is24Hour = -not ($time -match '(?i)(AM|PM)')
                    Position = $match.Index
                    ConfidenceScore = 0.85
                }
            }
        }
        
        # IP Address patterns
        $ipPattern = '\b(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\b'
        $ipMatches = [regex]::Matches($Content, $ipPattern)
        foreach ($match in $ipMatches) {
            $ip = $match.Value
            $coreEntities.IPAddresses += @{
                Value = $ip
                IsPrivate = Test-PrivateIP -IP $ip
                Position = $match.Index
                ConfidenceScore = 0.95
            }
        }
        
        # Monetary amounts
        $moneyPatterns = @(
            '\$\d{1,3}(?:,\d{3})*(?:\.\d{2})?',
            '\b\d{1,3}(?:,\d{3})*(?:\.\d{2})?\s*(?:USD|EUR|GBP|CAD|AUD)\b'
        )
        
        foreach ($pattern in $moneyPatterns) {
            $moneyMatches = [regex]::Matches($Content, $pattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
            foreach ($match in $moneyMatches) {
                $amount = $match.Value
                $coreEntities.MonetaryAmounts += @{
                    Value = $amount
                    Currency = Extract-Currency -AmountString $amount
                    NumericValue = Extract-NumericValue -AmountString $amount
                    Position = $match.Index
                    ConfidenceScore = 0.90
                }
            }
        }
        
        # Percentages
        $percentagePattern = '\b\d+(?:\.\d+)?%\b'
        $percentageMatches = [regex]::Matches($Content, $percentagePattern)
        foreach ($match in $percentageMatches) {
            $percentage = $match.Value
            $coreEntities.Percentages += @{
                Value = $percentage
                NumericValue = [double]($percentage -replace '%', '')
                Position = $match.Index
                ConfidenceScore = 0.95
            }
        }
        
        return $coreEntities
        
    } catch {
        Write-Verbose "Warning: Core entity extraction failed: $($_.Exception.Message)"
        return @{
            EmailAddresses = @()
            PhoneNumbers = @()
            URLs = @()
            Dates = @()
            Times = @()
            IPAddresses = @()
            MonetaryAmounts = @()
            Percentages = @()
        }
    }
}

function Extract-BusinessEntities {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$Content,
        
        [Parameter(Mandatory=$false)]
        [hashtable]$EmailContext = @{}
    )
    
    try {
        $businessEntities = @{
            CompanyNames = @()
            Projects = @()
            Departments = @()
            Positions = @()
            MeetingIndicators = @()
            ActionItems = @()
            DeadlineReferences = @()
            DocumentReferences = @()
        }
        
        # Company name patterns (basic heuristics)
        $companyPatterns = @(
            '\b[A-Z][a-zA-Z\s&]{2,}\s+(?:Inc|LLC|Corp|Corporation|Company|Co|Ltd|Limited|Group|Associates|Solutions|Technologies|Systems|Services|Consulting|Partners)\b',
            '\b(?:Microsoft|Google|Apple|Amazon|Facebook|Oracle|IBM|Intel|Adobe|Salesforce|Netflix|Tesla|Spotify|Twitter|LinkedIn|Uber|Airbnb|Zoom|Slack)\b'
        )
        
        foreach ($pattern in $companyPatterns) {
            $companyMatches = [regex]::Matches($Content, $pattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
            foreach ($match in $companyMatches) {
                $company = $match.Value.Trim()
                $businessEntities.CompanyNames += @{
                    Value = $company
                    Position = $match.Index
                    ConfidenceScore = 0.75
                    Source = "Pattern"
                }
            }
        }
        
        # Project references
        $projectPatterns = @(
            '(?i)\b(?:project|initiative|program|campaign|effort)\s+["\']?([A-Z][a-zA-Z\s\d-]{2,20})["\']?',
            '(?i)\b([A-Z][a-zA-Z\d-]{3,15})\s+(?:project|initiative|program)',
            '(?i)(?:working on|part of|assigned to)\s+(?:the\s+)?([A-Z][a-zA-Z\s\d-]{3,25})'
        )
        
        foreach ($pattern in $projectPatterns) {
            $projectMatches = [regex]::Matches($Content, $pattern)
            foreach ($match in $projectMatches) {
                $project = $(if ($match.Groups.Count -gt 1) { $match.Groups[1].Value.Trim() } else { $match.Value.Trim() })
                $businessEntities.Projects += @{
                    Value = $project
                    Position = $match.Index
                    ConfidenceScore = 0.70
                    Context = Get-SurroundingContext -Content $Content -Position $match.Index -Length 50
                }
            }
        }
        
        # Department names
        $departmentPatterns = @(
            '\b(?:IT|HR|Finance|Marketing|Sales|Engineering|Operations|Legal|Compliance|Security|Research|Development|Quality|Support|Customer\s+Service|Human\s+Resources|Information\s+Technology)\s+(?:Department|Team|Division|Group)?\b',
            '\b(?:R&D|QA|DevOps|InfoSec|FinOps|MarCom)\b'
        )
        
        foreach ($pattern in $departmentPatterns) {
            $deptMatches = [regex]::Matches($Content, $pattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
            foreach ($match in $deptMatches) {
                $department = $match.Value.Trim()
                $businessEntities.Departments += @{
                    Value = $department
                    Position = $match.Index
                    ConfidenceScore = 0.80
                }
            }
        }
        
        # Position/Role titles
        $positionPatterns = @(
            '\b(?:CEO|CTO|CFO|COO|CIO|CMO|VP|Vice\s+President|Director|Manager|Lead|Senior|Principal|Architect|Engineer|Developer|Analyst|Specialist|Coordinator|Administrator)\b',
            '\b(?:Project\s+Manager|Product\s+Manager|Business\s+Analyst|System\s+Administrator|Software\s+Engineer|Data\s+Scientist|UX\s+Designer)\b'
        )
        
        foreach ($pattern in $positionPatterns) {
            $positionMatches = [regex]::Matches($Content, $pattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
            foreach ($match in $positionMatches) {
                $position = $match.Value.Trim()
                $businessEntities.Positions += @{
                    Value = $position
                    Position = $match.Index
                    ConfidenceScore = 0.85
                    Level = Get-PositionLevel -Position $position
                }
            }
        }
        
        # Meeting indicators
        $meetingPatterns = @(
            '(?i)\b(?:meeting|call|conference|webinar|presentation|demo|review|standup|sync|catchup|one-on-one|1:1)\b',
            '(?i)\b(?:zoom|teams|skype|webex|meet|hangout)\s+(?:meeting|call|link)',
            '(?i)(?:scheduled for|meeting at|call at|conference room|dial in)'
        )
        
        foreach ($pattern in $meetingPatterns) {
            $meetingMatches = [regex]::Matches($Content, $pattern)
            foreach ($match in $meetingMatches) {
                $meeting = $match.Value.Trim()
                $businessEntities.MeetingIndicators += @{
                    Value = $meeting
                    Position = $match.Index
                    ConfidenceScore = 0.75
                    Context = Get-SurroundingContext -Content $Content -Position $match.Index -Length 100
                }
            }
        }
        
        # Action items
        $actionPatterns = @(
            '(?i)(?:please|could\s+you|can\s+you|need\s+to|should|must|action\s+item|todo|to\s+do)',
            '(?i)(?:deadline|due\s+date|by\s+\w+day|urgently|asap|priority)',
            '(?i)(?:follow\s+up|next\s+steps|action\s+required|pending)'
        )
        
        foreach ($pattern in $actionPatterns) {
            $actionMatches = [regex]::Matches($Content, $pattern)
            foreach ($match in $actionMatches) {
                $action = $match.Value.Trim()
                $businessEntities.ActionItems += @{
                    Value = $action
                    Position = $match.Index
                    ConfidenceScore = 0.65
                    Priority = Get-ActionPriority -ActionText $action
                    Context = Get-SurroundingContext -Content $Content -Position $match.Index -Length 80
                }
            }
        }
        
        # Deadline references
        $deadlinePatterns = @(
            '(?i)(?:deadline|due|expires?|ends?)\s+(?:on|by|at)?\s*([A-Za-z]+\s+\d{1,2}(?:,?\s+\d{4})?)',
            '(?i)(?:by|before|until)\s+(\w+day)',
            '(?i)(?:eod|end\s+of\s+day|cop|close\s+of\s+play)'
        )
        
        foreach ($pattern in $deadlinePatterns) {
            $deadlineMatches = [regex]::Matches($Content, $pattern)
            foreach ($match in $deadlineMatches) {
                $deadline = $match.Value.Trim()
                $businessEntities.DeadlineReferences += @{
                    Value = $deadline
                    Position = $match.Index
                    ConfidenceScore = 0.80
                    Urgency = Get-DeadlineUrgency -DeadlineText $deadline
                }
            }
        }
        
        # Document references
        $documentPatterns = @(
            '\b[A-Za-z0-9_-]+\.(?:pdf|doc|docx|xls|xlsx|ppt|pptx|txt|csv)\b',
            '(?i)(?:document|file|report|spreadsheet|presentation|proposal|contract|agreement)\s+(?:named|called|titled)?\s*["\']?([A-Za-z0-9\s_-]{3,30})["\']?',
            '(?i)(?:attached|attachment|please\s+find|see\s+attached)'
        )
        
        foreach ($pattern in $documentPatterns) {
            $docMatches = [regex]::Matches($Content, $pattern)
            foreach ($match in $docMatches) {
                $document = $match.Value.Trim()
                $businessEntities.DocumentReferences += @{
                    Value = $document
                    Position = $match.Index
                    ConfidenceScore = 0.75
                    Type = Get-DocumentType -DocumentText $document
                }
            }
        }
        
        return $businessEntities
        
    } catch {
        Write-Verbose "Warning: Business entity extraction failed: $($_.Exception.Message)"
        return @{
            CompanyNames = @()
            Projects = @()
            Departments = @()
            Positions = @()
            MeetingIndicators = @()
            ActionItems = @()
            DeadlineReferences = @()
            DocumentReferences = @()
        }
    }
}

function Extract-PersonalInformation {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$Content
    )
    
    try {
        $personalInfo = @{
            PersonNames = @()
            Locations = @()
            PersonalEvents = @()
            ContactInfo = @()
        }
        
        # Person name patterns (basic)
        $namePatterns = @(
            '\b[A-Z][a-z]+\s+[A-Z][a-z]+(?:\s+[A-Z][a-z]+)?\b',  # First Last (Middle)
            '(?i)(?:mr|mrs|ms|dr|prof|professor)\s+[A-Z][a-z]+(?:\s+[A-Z][a-z]+)?'  # Title + Name
        )
        
        foreach ($pattern in $namePatterns) {
            $nameMatches = [regex]::Matches($Content, $pattern)
            foreach ($match in $nameMatches) {
                $name = $match.Value.Trim()
                # Filter out common false positives
                if (-not (Test-CommonWord -Word $name) -and $name.Length -lt 40) {
                    $personalInfo.PersonNames += @{
                        Value = $name
                        Position = $match.Index
                        ConfidenceScore = 0.60
                        HasTitle = $name -match '(?i)^(?:mr|mrs|ms|dr|prof)'
                    }
                }
            }
        }
        
        # Location patterns
        $locationPatterns = @(
            '\b[A-Z][a-zA-Z\s]{2,20},\s*[A-Z]{2}\b',  # City, State
            '\b[A-Z][a-zA-Z\s]{2,20},\s*[A-Z][a-zA-Z\s]{4,20}\b',  # City, Country
            '\b\d{5}(?:-\d{4})?\b',  # ZIP codes
            '(?i)\b(?:New\s+York|Los\s+Angeles|Chicago|Houston|Phoenix|Philadelphia|San\s+Antonio|San\s+Diego|Dallas|San\s+Jose|Austin|Jacksonville|Fort\s+Worth|Columbus|Charlotte|San\s+Francisco|Indianapolis|Seattle|Denver|Washington|Boston|Nashville|Detroit|Oklahoma\s+City|Portland|Las\s+Vegas|Memphis|Louisville|Baltimore|Milwaukee|Albuquerque|Tucson|Fresno|Sacramento|Mesa|Kansas\s+City|Atlanta|Omaha|Colorado\s+Springs|Raleigh|Miami|Long\s+Beach|Virginia\s+Beach|Oakland|Minneapolis|Tampa)\b'
        )
        
        foreach ($pattern in $locationPatterns) {
            $locationMatches = [regex]::Matches($Content, $pattern)
            foreach ($match in $locationMatches) {
                $location = $match.Value.Trim()
                $personalInfo.Locations += @{
                    Value = $location
                    Position = $match.Index
                    ConfidenceScore = 0.70
                    Type = Get-LocationType -Location $location
                }
            }
        }
        
        # Personal events
        $eventPatterns = @(
            '(?i)\b(?:birthday|anniversary|wedding|graduation|vacation|holiday|travel|trip)\b',
            '(?i)(?:going\s+to|visiting|traveling\s+to|vacation\s+in)'
        )
        
        foreach ($pattern in $eventPatterns) {
            $eventMatches = [regex]::Matches($Content, $pattern)
            foreach ($match in $eventMatches) {
                $event = $match.Value.Trim()
                $personalInfo.PersonalEvents += @{
                    Value = $event
                    Position = $match.Index
                    ConfidenceScore = 0.65
                    Context = Get-SurroundingContext -Content $Content -Position $match.Index -Length 60
                }
            }
        }
        
        return $personalInfo
        
    } catch {
        Write-Verbose "Warning: Personal information extraction failed: $($_.Exception.Message)"
        return @{
            PersonNames = @()
            Locations = @()
            PersonalEvents = @()
            ContactInfo = @()
        }
    }
}

function Extract-TechnicalEntities {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$Content
    )
    
    try {
        $technicalEntities = @{
            Technologies = @()
            ProgrammingLanguages = @()
            Protocols = @()
            FileFormats = @()
            TechnicalTerms = @()
            SystemNames = @()
        }
        
        # Technology and platform names
        $technologyPatterns = @(
            '\b(?:Azure|AWS|Google\s+Cloud|GCP|Docker|Kubernetes|Jenkins|Git|GitHub|GitLab|Bitbucket|Jira|Confluence|Slack|Teams|Zoom|Office\s+365|SharePoint|OneDrive|Salesforce|ServiceNow|Tableau|Power\s+BI|Splunk|Elasticsearch|MongoDB|PostgreSQL|MySQL|Oracle|SQL\s+Server|Redis|Kafka|RabbitMQ|Nginx|Apache|IIS|Tomcat|Node\.js|React|Angular|Vue\.js|Express|Spring|Django|Flask|Laravel|Rails|\.NET|Java|Python|JavaScript|TypeScript|C#|C\+\+|Go|Rust|PHP|Ruby|Swift|Kotlin|Scala|Perl|PowerShell|Bash|Linux|Windows|macOS|Ubuntu|CentOS|RHEL|Debian|SUSE)\b'
        )
        
        $techMatches = [regex]::Matches($Content, $technologyPatterns, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
        foreach ($match in $techMatches) {
            $tech = $match.Value
            $technicalEntities.Technologies += @{
                Value = $tech
                Position = $match.Index
                ConfidenceScore = 0.85
                Category = Get-TechnologyCategory -Technology $tech
            }
        }
        
        # Programming languages (more specific)
        $langPatterns = @(
            '\b(?:JavaScript|TypeScript|Python|Java|C#|C\+\+|Go|Rust|PHP|Ruby|Swift|Kotlin|Scala|Perl|PowerShell|Bash|SQL|HTML|CSS|XML|JSON|YAML|Markdown)\b'
        )
        
        $langMatches = [regex]::Matches($Content, $langPatterns, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
        foreach ($match in $langMatches) {
            $language = $match.Value
            $technicalEntities.ProgrammingLanguages += @{
                Value = $language
                Position = $match.Index
                ConfidenceScore = 0.90
                Type = Get-LanguageType -Language $language
            }
        }
        
        # Network protocols and technical terms
        $protocolPatterns = @(
            '\b(?:HTTP|HTTPS|FTP|SFTP|SSH|SSL|TLS|TCP|UDP|SMTP|POP3|IMAP|DNS|DHCP|VPN|API|REST|SOAP|GraphQL|OAuth|SAML|JWT|LDAP|AD|Active\s+Directory)\b'
        )
        
        $protocolMatches = [regex]::Matches($Content, $protocolPatterns, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
        foreach ($match in $protocolMatches) {
            $protocol = $match.Value
            $technicalEntities.Protocols += @{
                Value = $protocol
                Position = $match.Index
                ConfidenceScore = 0.90
                Category = Get-ProtocolCategory -Protocol $protocol
            }
        }
        
        # File formats and extensions
        $fileFormatPattern = '\b[a-zA-Z0-9_-]+\.(?:exe|dll|msi|zip|rar|tar|gz|pdf|doc|docx|xls|xlsx|ppt|pptx|txt|csv|xml|json|yaml|yml|html|css|js|ts|py|java|cs|cpp|h|php|rb|go|rs|sql|md|log|config|ini|properties|env)\b'
        $fileFormatMatches = [regex]::Matches($Content, $fileFormatPattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
        foreach ($match in $fileFormatMatches) {
            $file = $match.Value
            $extension = [System.IO.Path]::GetExtension($file)
            $technicalEntities.FileFormats += @{
                Value = $file
                Extension = $extension
                Position = $match.Index
                ConfidenceScore = 0.95
                Category = Get-FileCategory -Extension $extension
            }
        }
        
        # System and server names
        $systemPatterns = @(
            '\b(?:server|database|db|prod|production|staging|dev|development|test|qa|uat|localhost|webapp|website|portal|dashboard|api|service|microservice|container|pod|cluster|node|instance|vm|virtual\s+machine)\b[-\w]*',
            '\b[A-Z][A-Z0-9]{2,}-[A-Z0-9]{3,}\b'  # Server naming conventions
        )
        
        foreach ($pattern in $systemPatterns) {
            $systemMatches = [regex]::Matches($Content, $pattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
            foreach ($match in $systemMatches) {
                $system = $match.Value
                $technicalEntities.SystemNames += @{
                    Value = $system
                    Position = $match.Index
                    ConfidenceScore = 0.75
                    Environment = Get-SystemEnvironment -SystemName $system
                }
            }
        }
        
        return $technicalEntities
        
    } catch {
        Write-Verbose "Warning: Technical entity extraction failed: $($_.Exception.Message)"
        return @{
            Technologies = @()
            ProgrammingLanguages = @()
            Protocols = @()
            FileFormats = @()
            TechnicalTerms = @()
            SystemNames = @()
        }
    }
}

function Extract-ContextAwareEntities {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$Content,
        
        [Parameter(Mandatory=$false)]
        [hashtable]$EmailContext = @{}
    )
    
    try {
        $contextEntities = @{
            Topics = @()
            Sentiment = @()
            Intent = @()
            Priority = @()
        }
        
        # Topic modeling (basic keyword clustering)
        $topicKeywords = @{
            'Project Management' = @('project', 'milestone', 'deliverable', 'timeline', 'resource', 'budget', 'scope', 'requirement')
            'Technical Support' = @('issue', 'problem', 'error', 'bug', 'fix', 'troubleshoot', 'resolve', 'support')
            'Sales' = @('proposal', 'quote', 'deal', 'contract', 'client', 'customer', 'revenue', 'target')
            'HR' = @('employee', 'hire', 'interview', 'performance', 'review', 'training', 'policy', 'benefit')
            'Finance' = @('budget', 'cost', 'expense', 'invoice', 'payment', 'revenue', 'profit', 'financial')
            'Marketing' = @('campaign', 'brand', 'promotion', 'advertisement', 'lead', 'conversion', 'engagement')
        }
        
        foreach ($topic in $topicKeywords.Keys) {
            $keywordCount = 0
            foreach ($keyword in $topicKeywords[$topic]) {
                if ($Content -match "(?i)\b$keyword\b") {
                    $keywordCount++
                }
            }
            
            if ($keywordCount -gt 0) {
                $contextEntities.Topics += @{
                    Value = $topic
                    MatchCount = $keywordCount
                    ConfidenceScore = [Math]::Min(0.95, $keywordCount / $topicKeywords[$topic].Count)
                    Keywords = $topicKeywords[$topic] | Where-Object { $Content -match "(?i)\b$_\b" }
                }
            }
        }
        
        # Basic sentiment analysis
        $positiveWords = @('great', 'excellent', 'good', 'pleased', 'happy', 'satisfied', 'success', 'achievement', 'congratulations', 'thank')
        $negativeWords = @('problem', 'issue', 'concern', 'error', 'failed', 'disappointed', 'urgent', 'critical', 'blocker', 'delay')
        
        $positiveCount = 0
        $negativeCount = 0
        
        foreach ($word in $positiveWords) {
            $positiveCount += ([regex]::Matches($Content, "(?i)\b$word\b")).Count
        }
        
        foreach ($word in $negativeWords) {
            $negativeCount += ([regex]::Matches($Content, "(?i)\b$word\b")).Count
        }
        
        $sentimentScore = if (($positiveCount + $negativeCount) -gt 0) {
            ($positiveCount - $negativeCount) / ($positiveCount + $negativeCount)
        } else { 0 }
        
        $contextEntities.Sentiment += @{
            Score = [Math]::Round($sentimentScore, 2)
            Classification = $(if ($sentimentScore -gt 0.2) { "Positive" } elseif ($sentimentScore -lt -0.2) { "Negative" } else { "Neutral" })
            PositiveWords = $positiveCount
            NegativeWords = $negativeCount
            ConfidenceScore = [Math]::Min(0.80, [Math]::Abs($sentimentScore))
        }
        
        # Intent detection
        $intentPatterns = @{
            'Request' = @('please', 'could you', 'can you', 'need', 'require', 'request')
            'Information' = @('what', 'when', 'where', 'how', 'why', 'which', 'status', 'update')
            'Action' = @('do', 'complete', 'finish', 'implement', 'create', 'develop', 'build')
            'Meeting' = @('meet', 'call', 'discuss', 'review', 'presentation', 'demo')
            'Follow-up' = @('follow up', 'check in', 'reminder', 'pending', 'waiting')
        }
        
        foreach ($intent in $intentPatterns.Keys) {
            $matchCount = 0
            foreach ($pattern in $intentPatterns[$intent]) {
                $matchCount += ([regex]::Matches($Content, "(?i)\b$pattern\b")).Count
            }
            
            if ($matchCount -gt 0) {
                $contextEntities.Intent += @{
                    Value = $intent
                    MatchCount = $matchCount
                    ConfidenceScore = [Math]::Min(0.90, $matchCount / 5.0)
                }
            }
        }
        
        # Priority detection
        $priorityPatterns = @{
            'High' = @('urgent', 'asap', 'critical', 'emergency', 'immediately', 'high priority')
            'Medium' = @('important', 'soon', 'priority', 'needed', 'required')
            'Low' = @('when possible', 'low priority', 'no rush', 'eventually')
        }
        
        foreach ($priority in $priorityPatterns.Keys) {
            foreach ($pattern in $priorityPatterns[$priority]) {
                if ($Content -match "(?i)\b$pattern\b") {
                    $contextEntities.Priority += @{
                        Value = $priority
                        MatchedPattern = $pattern
                        ConfidenceScore = 0.80
                    }
                    break
                }
            }
        }
        
        return $contextEntities
        
    } catch {
        Write-Verbose "Warning: Context-aware entity extraction failed: $($_.Exception.Message)"
        return @{
            Topics = @()
            Sentiment = @()
            Intent = @()
            Priority = @()
        }
    }
}

# Helper functions (continued in next part due to length)
function Get-EmptyEntityResult {
    return @{
        ExtractedAt = Get-Date -Format "yyyy-MM-ddTHH:mm:ss.fffZ"
        ExtractorVersion = "2.0"
        ContentLength = 0
        Entities = @{}
        Summary = @{ TotalEntities = 0; EmptyContent = $true }
    }
}

function Test-BusinessDomain {
    [CmdletBinding()]
    param([string]$Domain)
    
    $personalDomains = @('gmail.com', 'yahoo.com', 'hotmail.com', 'outlook.com', 'aol.com', 'icloud.com')
    return $Domain -notin $personalDomains
}

function Get-PhoneFormat {
    [CmdletBinding()]
    param([string]$Phone)
    
    if ($Phone -match '\+') { return "International" }
    elseif ($Phone -match '\(\d{3}\)') { return "US Format" }
    else { return "Standard" }
}

function Parse-DateString {
    [CmdletBinding()]
    param([string]$DateString)
    
    try {
        $parsedDate = [DateTime]::Parse($DateString)
        return @{
            ParsedDate = $parsedDate.ToString("yyyy-MM-dd")
            Format = "Parsed"
            IsValid = $true
        }
    } catch {
        return @{
            ParsedDate = $null
            Format = "Unknown"
            IsValid = $false
        }
    }
}

function Calculate-EntitySummary {
    [CmdletBinding()]
    param([hashtable]$EntityResults)
    
    $totalEntities = 0
    $highConfidenceEntities = 0
    
    foreach ($category in $EntityResults.Entities.Keys) {
        foreach ($entityType in $EntityResults.Entities[$category].Keys) {
            $entities = $EntityResults.Entities[$category][$entityType]
            if ($entities -is [array]) {
                $totalEntities += $entities.Count
                $highConfidenceEntities += ($entities | Where-Object { $_.ConfidenceScore -ge 0.8 }).Count
            }
        }
    }
    
    return @{
        TotalEntities = $totalEntities
        HighConfidenceEntities = $highConfidenceEntities
        Categories = $EntityResults.Entities.Keys.Count
        ConfidenceRatio = $(if ($totalEntities -gt 0) { [Math]::Round($highConfidenceEntities / $totalEntities, 2) } else { 0 })
    }
}

# Additional helper functions with basic implementations
function Test-PrivateIP { param($IP) return $IP.StartsWith('10.') -or $IP.StartsWith('192.168.') -or $IP.StartsWith('172.') }
function Extract-DomainFromURL { param($URL) try { return ([System.Uri]$URL).Host } catch { return "" } }
function Extract-ProtocolFromURL { param($URL) try { return ([System.Uri]$URL).Scheme } catch { return "" } }
function Extract-Currency { 
    param($AmountString) 
    if ($AmountString -match '\$') { 
        return "USD" 
    } else { 
        return "Unknown" 
    } 
}
function Extract-NumericValue { 
    param($AmountString) 
    return [regex]::Replace($AmountString, '[^\d.]', '') 
}
function Get-SurroundingContext { 
    param($Content, $Position, $Length) 
    $start = [Math]::Max(0, $Position - $Length/2) 
    return $Content.Substring($start, [Math]::Min($Length, $Content.Length - $start)) 
}
function Get-PositionLevel { 
    param($Position) 
    if ($Position -match '(?i)(CEO|CTO|VP|Director)') { 
        return "Executive" 
    } elseif ($Position -match '(?i)(Manager|Lead)') { 
        return "Management" 
    } else { 
        return "Individual Contributor" 
    } 
}
function Get-ActionPriority { 
    param($ActionText) 
    if ($ActionText -match '(?i)(urgent|asap|critical)') { 
        return "High" 
    } else { 
        return "Medium" 
    } 
}
function Get-DeadlineUrgency { 
    param($DeadlineText) 
    if ($DeadlineText -match '(?i)(today|tomorrow|asap|urgent)') { 
        return "High" 
    } else { 
        return "Medium" 
    } 
}
function Get-DocumentType { 
    param($DocumentText) 
    if ($DocumentText -match '\.pdf$') { 
        return "PDF" 
    } elseif ($DocumentText -match '\.(doc|docx)$') { 
        return "Word" 
    } else { 
        return "Other" 
    } 
}
function Test-CommonWord { 
    param($Word) 
    $commonWords = @('The New', 'For You', 'To Do', 'On The') 
    return $Word -in $commonWords 
}
function Get-LocationType { 
    param($Location) 
    if ($Location -match '\d{5}') { 
        return "ZIP Code" 
    } elseif ($Location -match ',') { 
        return "City, State/Country" 
    } else { 
        return "General" 
    } 
}
function Get-TechnologyCategory { param($Technology) if ($Technology -match '(?i)(Azure|AWS|GCP)') { return "Cloud" } else { return "General" } }
function Get-LanguageType { param($Language) if ($Language -match '(?i)(HTML|CSS|XML)') { return "Markup" } else { return "Programming" } }
function Get-ProtocolCategory { param($Protocol) if ($Protocol -match '(?i)(HTTP|FTP|SSH)') { return "Network" } else { return "General" } }
function Get-FileCategory { 
    param($Extension) 
    if ($Extension -match '(?i)\.(exe|dll|msi)') { 
        return "Executable" 
    } else { 
        return "Document" 
    } 
}
function Get-SystemEnvironment { 
    param($SystemName) 
    if ($SystemName -match '(?i)(prod|production)') { 
        return "Production" 
    } elseif ($SystemName -match '(?i)(dev|test)') { 
        return "Development" 
    } else { 
        return "Unknown" 
    } 
}

function Filter-EntitiesByConfidence {
    [CmdletBinding()]
    param(
        [hashtable]$EntityResults,
        [double]$MinScore
    )
    
    foreach ($category in $EntityResults.Entities.Keys) {
        foreach ($entityType in $EntityResults.Entities[$category].Keys) {
            $entities = $EntityResults.Entities[$category][$entityType]
            if ($entities -is [array]) {
                $EntityResults.Entities[$category][$entityType] = $entities | Where-Object { $_.ConfidenceScore -ge $MinScore }
            }
        }
    }
    
    return $EntityResults
}

Write-Verbose "EmailEntityExtractor_v2 module loaded successfully"