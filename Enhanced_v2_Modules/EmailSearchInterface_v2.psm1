# EmailSearchInterface_v2.psm1 - Hybrid Search Interface for Email RAG System
# User-friendly interface for searching emails with vector and keyword search capabilities

Export-ModuleMember -Function @(
    'Find-EmailContent',
    'Search-EmailsBySender',
    'Search-EmailsByDateRange',
    'Search-EmailsAdvanced',
    'Get-RelatedEmails',
    'Export-SearchResults',
    'New-SearchReport'
)

# Import required modules
Import-Module (Join-Path $PSScriptRoot "AzureAISearchIntegration_v2.psm1") -Force

function Find-EmailContent {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$SearchConfig,
        
        [Parameter(Mandatory=$true)]
        [string]$Query,
        
        [Parameter(Mandatory=$false)]
        [string]$IndexName = "email-rag-index",
        
        [Parameter(Mandatory=$false)]
        [int]$MaxResults = 20,
        
        [Parameter(Mandatory=$false)]
        [ValidateSet("Keyword", "Semantic", "Hybrid", "Vector")]
        [string]$SearchType = "Hybrid",
        
        [Parameter(Mandatory=$false)]
        [switch]$IncludeContext = $true,
        
        [Parameter(Mandatory=$false)]
        [switch]$GroupByEmail = $true,
        
        [Parameter(Mandatory=$false)]
        [double]$MinimumScore = 0.0
    )
    
    try {
        Write-Verbose "Searching for: '$Query' using $SearchType search"
        
        $searchResults = switch ($SearchType) {
            "Keyword" {
                Search-EmailDocuments -SearchConfig $SearchConfig -Query $Query -IndexName $IndexName -Top $MaxResults
            }
            "Semantic" {
                Search-EmailDocuments -SearchConfig $SearchConfig -Query $Query -IndexName $IndexName -Top $MaxResults
            }
            "Hybrid" {
                Search-EmailHybrid -SearchConfig $SearchConfig -Query $Query -IndexName $IndexName -Top $MaxResults -UseSemanticSearch:$true -GenerateQueryEmbedding:$true
            }
            "Vector" {
                Search-EmailHybrid -SearchConfig $SearchConfig -Query $Query -IndexName $IndexName -Top $MaxResults -UseSemanticSearch:$false -GenerateQueryEmbedding:$true
            }
        }
        
        if (-not $searchResults.Results -or $searchResults.Results.Count -eq 0) {
            Write-Host "No results found for query: '$Query'" -ForegroundColor Yellow
            return @{
                Query = $Query
                SearchType = $SearchType
                TotalResults = 0
                Results = @()
                SearchedAt = Get-Date -Format "yyyy-MM-ddTHH:mm:ss.fffZ"
            }
        }
        
        # Filter by minimum score if specified
        $filteredResults = if ($MinimumScore -gt 0) {
            $searchResults.Results | Where-Object { 
                $score = if ($_.HybridScore) { $_.HybridScore } elseif ($_.'@search.score') { $_.'@search.score' } else { 1.0 }
                $score -ge $MinimumScore
            }
        } else {
            $searchResults.Results
        }
        
        # Enhance results with additional context
        $enhancedResults = @()
        foreach ($result in $filteredResults) {
            $enhanced = Enhance-SearchResult -Result $result -IncludeContext:$IncludeContext
            $enhancedResults += $enhanced
        }
        
        # Group by email if requested
        $finalResults = if ($GroupByEmail -and $enhancedResults.Count -gt 1) {
            Group-ResultsByEmail -Results $enhancedResults
        } else {
            $enhancedResults
        }
        
        # Display results summary
        Write-Host "`nSearch Results for: '$Query'" -ForegroundColor Green
        Write-Host "Search Type: $SearchType" -ForegroundColor Cyan
        Write-Host "Total Results: $($finalResults.Count)" -ForegroundColor Cyan
        
        if ($finalResults.Count -gt 0) {
            Write-Host "`nTop Results:" -ForegroundColor Yellow
            for ($i = 0; $i -lt [Math]::Min(5, $finalResults.Count); $i++) {
                $result = $finalResults[$i]
                $score = if ($result.SearchScore) { " (Score: $($result.SearchScore))" } else { "" }
                Write-Host "  $($i + 1). $($result.Subject)$score" -ForegroundColor White
                Write-Host "     From: $($result.SenderName) | Date: $($result.SentDate)" -ForegroundColor Gray
                if ($result.ContentPreview) {
                    $preview = if ($result.ContentPreview.Length -gt 100) { $result.ContentPreview.Substring(0, 100) + "..." } else { $result.ContentPreview }
                    Write-Host "     Preview: $preview" -ForegroundColor Gray
                }
                Write-Host ""
            }
        }
        
        return @{
            Query = $Query
            SearchType = $SearchType
            TotalResults = $finalResults.Count
            Results = $finalResults
            SearchedAt = Get-Date -Format "yyyy-MM-ddTHH:mm:ss.fffZ"
            SearchStats = $searchResults
        }
        
    } catch {
        Write-Error "Search failed: $($_.Exception.Message)"
        throw
    }
}

function Search-EmailsBySender {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$SearchConfig,
        
        [Parameter(Mandatory=$true)]
        [string]$SenderName,
        
        [Parameter(Mandatory=$false)]
        [string]$IndexName = "email-rag-index",
        
        [Parameter(Mandatory=$false)]
        [int]$MaxResults = 50,
        
        [Parameter(Mandatory=$false)]
        [string]$AdditionalQuery,
        
        [Parameter(Mandatory=$false)]
        [switch]$ExactMatch = $false
    )
    
    try {
        $filter = if ($ExactMatch) {
            "sender_name eq '$SenderName'"
        } else {
            "search.ismatch('$SenderName', 'sender_name')"
        }
        
        $searchQuery = if ($AdditionalQuery) {
            $AdditionalQuery
        } else {
            "*"
        }
        
        Write-Verbose "Searching emails from sender: $SenderName"
        
        $searchResults = Search-EmailDocuments -SearchConfig $SearchConfig -Query $searchQuery -IndexName $IndexName -Top $MaxResults -Filter $filter -OrderBy "sent_date desc"
        
        $enhancedResults = @()
        foreach ($result in $searchResults.Results) {
            $enhanced = @{
                Id = $result.id
                Subject = $result.title
                SenderName = $result.sender_name
                SenderEmail = $result.sender_email
                SentDate = $result.sent_date
                ContentPreview = if ($result.content) { $result.content.Substring(0, [Math]::Min(200, $result.content.Length)) } else { "" }
                ChunkInfo = if ($result.chunk_number) { "Chunk $($result.chunk_number) of $($result.total_chunks)" } else { "Full Email" }
                SearchScore = if ($result.'@search.score') { [Math]::Round($result.'@search.score', 2) } else { 0 }
            }
            $enhancedResults += $enhanced
        }
        
        Write-Host "`nEmails from '$SenderName': $($enhancedResults.Count) results" -ForegroundColor Green
        
        return @{
            SenderName = $SenderName
            TotalResults = $enhancedResults.Count
            Results = $enhancedResults
            SearchedAt = Get-Date -Format "yyyy-MM-ddTHH:mm:ss.fffZ"
        }
        
    } catch {
        Write-Error "Sender search failed: $($_.Exception.Message)"
        throw
    }
}

function Search-EmailsByDateRange {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$SearchConfig,
        
        [Parameter(Mandatory=$true)]
        [DateTime]$StartDate,
        
        [Parameter(Mandatory=$true)]
        [DateTime]$EndDate,
        
        [Parameter(Mandatory=$false)]
        [string]$IndexName = "email-rag-index",
        
        [Parameter(Mandatory=$false)]
        [string]$Query = "*",
        
        [Parameter(Mandatory=$false)]
        [int]$MaxResults = 100
    )
    
    try {
        $startDateStr = $StartDate.ToString("yyyy-MM-ddTHH:mm:ss.fffZ")
        $endDateStr = $EndDate.ToString("yyyy-MM-ddTHH:mm:ss.fffZ")
        
        $filter = "sent_date ge $startDateStr and sent_date le $endDateStr"
        
        Write-Verbose "Searching emails between $($StartDate.ToString("yyyy-MM-dd")) and $($EndDate.ToString("yyyy-MM-dd"))"
        
        $searchResults = Search-EmailDocuments -SearchConfig $SearchConfig -Query $Query -IndexName $IndexName -Top $MaxResults -Filter $filter -OrderBy "sent_date desc"
        
        # Group results by date
        $resultsByDate = @{}
        foreach ($result in $searchResults.Results) {
            $dateKey = if ($result.sent_date) {
                ([DateTime]$result.sent_date).ToString("yyyy-MM-dd")
            } else {
                "Unknown Date"
            }
            
            if (-not $resultsByDate.ContainsKey($dateKey)) {
                $resultsByDate[$dateKey] = @()
            }
            
            $resultsByDate[$dateKey] += @{
                Id = $result.id
                Subject = $result.title
                SenderName = $result.sender_name
                SentDate = $result.sent_date
                ContentPreview = if ($result.content) { $result.content.Substring(0, [Math]::Min(150, $result.content.Length)) } else { "" }
            }
        }
        
        Write-Host "`nEmails from $($StartDate.ToString("yyyy-MM-dd")) to $($EndDate.ToString("yyyy-MM-dd")): $($searchResults.Results.Count) results" -ForegroundColor Green
        
        foreach ($date in ($resultsByDate.Keys | Sort-Object -Descending)) {
            Write-Host "`n$date ($($resultsByDate[$date].Count) emails):" -ForegroundColor Yellow
            foreach ($email in $resultsByDate[$date] | Select-Object -First 3) {
                Write-Host "  • $($email.Subject) - $($email.SenderName)" -ForegroundColor White
            }
        }
        
        return @{
            StartDate = $StartDate
            EndDate = $EndDate
            TotalResults = $searchResults.Results.Count
            ResultsByDate = $resultsByDate
            SearchedAt = Get-Date -Format "yyyy-MM-ddTHH:mm:ss.fffZ"
        }
        
    } catch {
        Write-Error "Date range search failed: $($_.Exception.Message)"
        throw
    }
}

function Search-EmailsAdvanced {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$SearchConfig,
        
        [Parameter(Mandatory=$false)]
        [string]$Query = "*",
        
        [Parameter(Mandatory=$false)]
        [string]$IndexName = "email-rag-index",
        
        [Parameter(Mandatory=$false)]
        [string]$SenderName,
        
        [Parameter(Mandatory=$false)]
        [string]$Subject,
        
        [Parameter(Mandatory=$false)]
        [DateTime]$StartDate,
        
        [Parameter(Mandatory=$false)]
        [DateTime]$EndDate,
        
        [Parameter(Mandatory=$false)]
        [switch]$HasAttachments,
        
        [Parameter(Mandatory=$false)]
        [ValidateSet("Low", "Normal", "High")]
        [string]$Importance,
        
        [Parameter(Mandatory=$false)]
        [string[]]$Keywords,
        
        [Parameter(Mandatory=$false)]
        [int]$MaxResults = 50,
        
        [Parameter(Mandatory=$false)]
        [ValidateSet("Relevance", "Date", "Sender")]
        [string]$SortBy = "Relevance"
    )
    
    try {
        # Build advanced filter
        $filterParts = @()
        
        if ($SenderName) {
            $filterParts += "search.ismatch('$SenderName', 'sender_name')"
        }
        
        if ($Subject) {
            $filterParts += "search.ismatch('$Subject', 'title')"
        }
        
        if ($StartDate) {
            $startDateStr = $StartDate.ToString("yyyy-MM-ddTHH:mm:ss.fffZ")
            $filterParts += "sent_date ge $startDateStr"
        }
        
        if ($EndDate) {
            $endDateStr = $EndDate.ToString("yyyy-MM-ddTHH:mm:ss.fffZ")
            $filterParts += "sent_date le $endDateStr"
        }
        
        if ($HasAttachments) {
            $filterParts += "has_attachments eq true"
        }
        
        if ($Importance) {
            $filterParts += "importance eq '$Importance'"
        }
        
        $filter = if ($filterParts.Count -gt 0) { $filterParts -join " and " } else { $null }
        
        # Build search query with keywords
        $searchQuery = if ($Keywords -and $Keywords.Count -gt 0) {
            if ($Query -eq "*") {
                $Keywords -join " OR "
            } else {
                "$Query AND ($($Keywords -join " OR "))"
            }
        } else {
            $Query
        }
        
        # Determine sort order
        $orderBy = switch ($SortBy) {
            "Date" { "sent_date desc" }
            "Sender" { "sender_name asc" }
            default { $null }
        }
        
        Write-Verbose "Executing advanced search with query: '$searchQuery'"
        if ($filter) { Write-Verbose "Filter: $filter" }
        
        $searchResults = Search-EmailDocuments -SearchConfig $SearchConfig -Query $searchQuery -IndexName $IndexName -Top $MaxResults -Filter $filter -OrderBy $orderBy
        
        # Create summary
        $summary = @{
            TotalResults = $searchResults.TotalResults
            QueryUsed = $searchQuery
            FilterUsed = $filter
            SortBy = $SortBy
            Criteria = @{
                SenderName = $SenderName
                Subject = $Subject
                DateRange = if ($StartDate -or $EndDate) { "$($StartDate) to $($EndDate)" } else { $null }
                HasAttachments = $HasAttachments.IsPresent
                Importance = $Importance
                Keywords = $Keywords
            }
        }
        
        Write-Host "`nAdvanced Search Results: $($searchResults.Results.Count) emails found" -ForegroundColor Green
        
        return @{
            Summary = $summary
            Results = $searchResults.Results | ForEach-Object {
                @{
                    Id = $_.id
                    Subject = $_.title
                    SenderName = $_.sender_name
                    SenderEmail = $_.sender_email
                    SentDate = $_.sent_date
                    HasAttachments = $_.has_attachments
                    Importance = $_.importance
                    ContentPreview = if ($_.content) { $_.content.Substring(0, [Math]::Min(200, $_.content.Length)) } else { "" }
                    SearchScore = if ($_.'@search.score') { [Math]::Round($_.'@search.score', 2) } else { 0 }
                }
            }
            SearchedAt = Get-Date -Format "yyyy-MM-ddTHH:mm:ss.fffZ"
        }
        
    } catch {
        Write-Error "Advanced search failed: $($_.Exception.Message)"
        throw
    }
}

function Get-RelatedEmails {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$SearchConfig,
        
        [Parameter(Mandatory=$true)]
        [string]$EmailId,
        
        [Parameter(Mandatory=$false)]
        [string]$IndexName = "email-rag-index",
        
        [Parameter(Mandatory=$false)]
        [int]$MaxResults = 10,
        
        [Parameter(Mandatory=$false)]
        [ValidateSet("Sender", "Subject", "Content", "All")]
        [string]$RelationshipType = "All"
    )
    
    try {
        # First, get the reference email
        Write-Verbose "Getting reference email: $EmailId"
        $referenceEmail = Search-EmailDocuments -SearchConfig $SearchConfig -Query "id:$EmailId" -IndexName $IndexName -Top 1
        
        if (-not $referenceEmail.Results -or $referenceEmail.Results.Count -eq 0) {
            throw "Reference email not found: $EmailId"
        }
        
        $refEmail = $referenceEmail.Results[0]
        $relatedEmails = @()
        
        # Find related emails based on relationship type
        switch ($RelationshipType) {
            "Sender" {
                if ($refEmail.sender_name) {
                    $senderResults = Search-EmailsBySender -SearchConfig $SearchConfig -SenderName $refEmail.sender_name -IndexName $IndexName -MaxResults $MaxResults
                    $relatedEmails += $senderResults.Results | Where-Object { $_.Id -ne $EmailId }
                }
            }
            "Subject" {
                if ($refEmail.title) {
                    # Extract key words from subject for similarity search
                    $subjectWords = ($refEmail.title -split '\s+' | Where-Object { $_.Length -gt 3 }) -join " OR "
                    $subjectResults = Search-EmailDocuments -SearchConfig $SearchConfig -Query $subjectWords -SearchFields @("title") -IndexName $IndexName -Top $MaxResults
                    $relatedEmails += $subjectResults.Results | Where-Object { $_.id -ne $EmailId }
                }
            }
            "Content" {
                if ($refEmail.content) {
                    # Use semantic search for content similarity
                    $contentPreview = $refEmail.content.Substring(0, [Math]::Min(500, $refEmail.content.Length))
                    $contentResults = Search-EmailHybrid -SearchConfig $SearchConfig -Query $contentPreview -IndexName $IndexName -Top $MaxResults -UseSemanticSearch:$true
                    $relatedEmails += $contentResults.Results | Where-Object { $_.id -ne $EmailId }
                }
            }
            "All" {
                # Combination approach
                if ($refEmail.sender_name) {
                    $filter = "sender_name eq '$($refEmail.sender_name)'"
                    $senderResults = Search-EmailDocuments -SearchConfig $SearchConfig -Query "*" -Filter $filter -IndexName $IndexName -Top 5
                    $relatedEmails += $senderResults.Results | Where-Object { $_.id -ne $EmailId }
                }
                
                if ($refEmail.title) {
                    $subjectWords = ($refEmail.title -split '\s+' | Where-Object { $_.Length -gt 3 } | Select-Object -First 5) -join " OR "
                    $subjectResults = Search-EmailDocuments -SearchConfig $SearchConfig -Query $subjectWords -SearchFields @("title") -IndexName $IndexName -Top 5
                    $relatedEmails += $subjectResults.Results | Where-Object { $_.id -ne $EmailId }
                }
            }
        }
        
        # Remove duplicates and score by relevance
        $uniqueRelated = @{}
        foreach ($email in $relatedEmails) {
            if (-not $uniqueRelated.ContainsKey($email.id)) {
                $email | Add-Member -NotePropertyName "RelationshipScore" -NotePropertyValue (Calculate-RelationshipScore -RefEmail $refEmail -CompareEmail $email)
                $uniqueRelated[$email.id] = $email
            }
        }
        
        $sortedRelated = $uniqueRelated.Values | Sort-Object RelationshipScore -Descending | Select-Object -First $MaxResults
        
        Write-Host "`nFound $($sortedRelated.Count) related emails for: $($refEmail.title)" -ForegroundColor Green
        
        foreach ($related in $sortedRelated | Select-Object -First 5) {
            Write-Host "  • $($related.title) (Score: $($related.RelationshipScore))" -ForegroundColor White
            Write-Host "    From: $($related.sender_name) | Date: $($related.sent_date)" -ForegroundColor Gray
        }
        
        return @{
            ReferenceEmail = @{
                Id = $refEmail.id
                Subject = $refEmail.title
                SenderName = $refEmail.sender_name
                SentDate = $refEmail.sent_date
            }
            RelatedEmails = $sortedRelated
            RelationshipType = $RelationshipType
            SearchedAt = Get-Date -Format "yyyy-MM-ddTHH:mm:ss.fffZ"
        }
        
    } catch {
        Write-Error "Related emails search failed: $($_.Exception.Message)"
        throw
    }
}

function Export-SearchResults {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$SearchResults,
        
        [Parameter(Mandatory=$true)]
        [string]$OutputPath,
        
        [Parameter(Mandatory=$false)]
        [ValidateSet("CSV", "JSON", "HTML")]
        [string]$Format = "CSV",
        
        [Parameter(Mandatory=$false)]
        [switch]$IncludeContent = $false
    )
    
    try {
        $exportData = @()
        
        foreach ($result in $SearchResults.Results) {
            $exportItem = @{
                Id = $result.Id
                Subject = $result.Subject
                SenderName = $result.SenderName
                SenderEmail = $result.SenderEmail
                SentDate = $result.SentDate
                SearchScore = $result.SearchScore
                ChunkInfo = $result.ChunkInfo
            }
            
            if ($IncludeContent) {
                $exportItem.ContentPreview = $result.ContentPreview
            }
            
            $exportData += New-Object PSObject -Property $exportItem
        }
        
        switch ($Format) {
            "CSV" {
                $exportData | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
            }
            "JSON" {
                $exportData | ConvertTo-Json -Depth 3 | Out-File -FilePath $OutputPath -Encoding UTF8
            }
            "HTML" {
                $html = Generate-SearchResultsHTML -Results $exportData -SearchInfo $SearchResults
                $html | Out-File -FilePath $OutputPath -Encoding UTF8
            }
        }
        
        Write-Host "Search results exported to: $OutputPath" -ForegroundColor Green
        Write-Host "Format: $Format | Records: $($exportData.Count)" -ForegroundColor Cyan
        
        return @{
            OutputPath = $OutputPath
            Format = $Format
            RecordCount = $exportData.Count
            ExportedAt = Get-Date -Format "yyyy-MM-ddTHH:mm:ss.fffZ"
        }
        
    } catch {
        Write-Error "Export failed: $($_.Exception.Message)"
        throw
    }
}

function New-SearchReport {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$SearchConfig,
        
        [Parameter(Mandatory=$false)]
        [string]$IndexName = "email-rag-index",
        
        [Parameter(Mandatory=$true)]
        [string]$ReportPath
    )
    
    try {
        Write-Host "Generating search analytics report..." -ForegroundColor Yellow
        
        # Get index statistics
        $indexStats = Get-SearchStatistics -SearchConfig $SearchConfig -IndexName $IndexName
        
        # Sample searches for analytics
        $sampleQueries = @("project", "meeting", "urgent", "deadline", "report")
        $queryAnalytics = @()
        
        foreach ($query in $sampleQueries) {
            try {
                $result = Search-EmailDocuments -SearchConfig $SearchConfig -Query $query -IndexName $IndexName -Top 10
                $queryAnalytics += @{
                    Query = $query
                    ResultCount = $result.TotalResults
                    AvgScore = if ($result.Results.Count -gt 0) { 
                        [Math]::Round(($result.Results | ForEach-Object { $_.'@search.score' } | Measure-Object -Average).Average, 2) 
                    } else { 0 }
                }
            } catch {
                $queryAnalytics += @{
                    Query = $query
                    ResultCount = 0
                    AvgScore = 0
                    Error = $_.Exception.Message
                }
            }
        }
        
        # Generate report
        $report = @{
            ReportTitle = "Email Search Analytics Report"
            GeneratedAt = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            IndexName = $IndexName
            IndexStatistics = $indexStats
            QueryAnalytics = $queryAnalytics
            SearchCapabilities = @{
                KeywordSearch = $true
                SemanticSearch = $true
                VectorSearch = ($SearchConfig.OpenAI -and $SearchConfig.OpenAI.ApiKey)
                HybridSearch = $true
            }
        }
        
        $reportJson = $report | ConvertTo-Json -Depth 5
        $reportJson | Out-File -FilePath $ReportPath -Encoding UTF8
        
        Write-Host "`nSearch Analytics Report Generated!" -ForegroundColor Green
        Write-Host "Report saved to: $ReportPath" -ForegroundColor Cyan
        Write-Host "Index contains: $($indexStats.DocumentCount) documents" -ForegroundColor Cyan
        
        return $report
        
    } catch {
        Write-Error "Report generation failed: $($_.Exception.Message)"
        throw
    }
}

# Helper functions
function Enhance-SearchResult {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        $Result,
        
        [Parameter(Mandatory=$false)]
        [switch]$IncludeContext = $true
    )
    
    return @{
        Id = $Result.id
        Subject = $Result.title
        SenderName = $Result.sender_name
        SenderEmail = $Result.sender_email
        SentDate = $Result.sent_date
        ContentPreview = if ($Result.content) { $Result.content.Substring(0, [Math]::Min(300, $Result.content.Length)) } else { "" }
        ChunkInfo = if ($Result.chunk_number) { "Chunk $($Result.chunk_number) of $($Result.total_chunks)" } else { "Full Email" }
        SearchScore = if ($Result.HybridScore) { [Math]::Round($Result.HybridScore, 2) } elseif ($Result.'@search.score') { [Math]::Round($Result.'@search.score', 2) } else { 0 }
        HasAttachments = $Result.has_attachments
        Importance = $Result.importance
        Keywords = $Result.keywords
    }
}

function Group-ResultsByEmail {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [array]$Results
    )
    
    $grouped = @{}
    
    foreach ($result in $Results) {
        $emailKey = if ($result.Subject) { $result.Subject } else { $result.Id }
        
        if (-not $grouped.ContainsKey($emailKey)) {
            $grouped[$emailKey] = @{
                Subject = $result.Subject
                SenderName = $result.SenderName
                SentDate = $result.SentDate
                Chunks = @()
                BestScore = $result.SearchScore
            }
        }
        
        $grouped[$emailKey].Chunks += $result
        if ($result.SearchScore -gt $grouped[$emailKey].BestScore) {
            $grouped[$emailKey].BestScore = $result.SearchScore
        }
    }
    
    return $grouped.Values | Sort-Object BestScore -Descending
}

function Calculate-RelationshipScore {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        $RefEmail,
        
        [Parameter(Mandatory=$true)]
        $CompareEmail
    )
    
    $score = 0
    
    # Same sender (high weight)
    if ($RefEmail.sender_name -eq $CompareEmail.sender_name) {
        $score += 50
    }
    
    # Similar subject (medium weight)
    if ($RefEmail.title -and $CompareEmail.title) {
        $refWords = $RefEmail.title -split '\s+' | ForEach-Object { $_.ToLower() }
        $compWords = $CompareEmail.title -split '\s+' | ForEach-Object { $_.ToLower() }
        $commonWords = $refWords | Where-Object { $_ -in $compWords }
        $score += ($commonWords.Count / [Math]::Max($refWords.Count, 1)) * 30
    }
    
    # Date proximity (low weight)
    if ($RefEmail.sent_date -and $CompareEmail.sent_date) {
        try {
            $refDate = [DateTime]$RefEmail.sent_date
            $compDate = [DateTime]$CompareEmail.sent_date
            $daysDiff = [Math]::Abs(($refDate - $compDate).TotalDays)
            if ($daysDiff -le 7) {
                $score += 20 - ($daysDiff * 2)
            }
        } catch {
            # Ignore date parsing errors
        }
    }
    
    return [Math]::Round($score, 1)
}

function Generate-SearchResultsHTML {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [array]$Results,
        
        [Parameter(Mandatory=$true)]
        [hashtable]$SearchInfo
    )
    
    $html = @"
<!DOCTYPE html>
<html>
<head>
    <title>Email Search Results</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        .header { background-color: #f0f0f0; padding: 10px; border-radius: 5px; }
        .result { border: 1px solid #ddd; margin: 10px 0; padding: 10px; border-radius: 5px; }
        .subject { font-weight: bold; color: #0066cc; }
        .meta { color: #666; font-size: 0.9em; }
        .score { background-color: #e6f3ff; padding: 2px 6px; border-radius: 3px; }
    </style>
</head>
<body>
    <div class="header">
        <h1>Email Search Results</h1>
        <p>Query: <strong>$($SearchInfo.Query)</strong></p>
        <p>Results: <strong>$($Results.Count)</strong> | Generated: <strong>$($SearchInfo.SearchedAt)</strong></p>
    </div>
"@

    foreach ($result in $Results) {
        $html += @"
    <div class="result">
        <div class="subject">$($result.Subject)</div>
        <div class="meta">
            From: $($result.SenderName) | Date: $($result.SentDate) | Score: <span class="score">$($result.SearchScore)</span>
        </div>
        <div class="content">$($result.ContentPreview)</div>
    </div>
"@
    }

    $html += @"
</body>
</html>
"@

    return $html
}

Write-Verbose "EmailSearchInterface_v2 module loaded successfully"