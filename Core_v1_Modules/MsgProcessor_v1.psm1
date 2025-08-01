# MsgProcessor.psm1 - MSG File Processing Module (Clean Version)
# Professional Outlook MSG file processor with COM interface integration

Export-ModuleMember -Function @(
    'Read-MsgFile',
    'Get-MsgMetadata',
    'Extract-MsgAttachments',
    'Test-MsgFile'
)

function Read-MsgFile {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$FilePath,
        
        [Parameter(Mandatory=$false)]
        [switch]$IncludeAttachments = $true,
        
        [Parameter(Mandatory=$false)]
        [switch]$ExtractText = $true
    )
    
    if (-not (Test-Path $FilePath)) {
        throw "MSG file not found: $FilePath"
    }
    
    if (-not (Test-MsgFile -FilePath $FilePath)) {
        throw "Invalid MSG file format: $FilePath"
    }
    
    $outlook = $null
    $msg = $null
    
    try {
        Write-Verbose "Initializing Outlook COM interface..."
        $outlook = New-Object -ComObject Outlook.Application
        $namespace = $outlook.GetNamespace("MAPI")
        
        Write-Verbose "Opening MSG file: $FilePath"
        $msg = $namespace.OpenSharedItem($FilePath)
        
        if (-not $msg) {
            throw "Failed to open MSG file with Outlook COM interface"
        }
        
        # Extract basic email data
        $emailData = @{
            Subject = Get-SafeProperty $msg "Subject"
            Body = Get-SafeProperty $msg "Body"
            HTMLBody = Get-SafeProperty $msg "HTMLBody"
            
            Sender = @{
                Name = Get-SafeProperty $msg "SenderName"
                Email = Get-SafeProperty $msg "SenderEmailAddress"
                Type = Get-SafeProperty $msg "SenderEmailType"
            }
            
            Recipients = @{
                To = @()
                CC = @()
                BCC = @()
            }
            
            Sent = Get-SafeProperty $msg "SentOn"
            Received = Get-SafeProperty $msg "ReceivedTime"
            CreationTime = Get-SafeProperty $msg "CreationTime"
            LastModificationTime = Get-SafeProperty $msg "LastModificationTime"
            
            Size = Get-SafeProperty $msg "Size"
            Importance = Get-SafeProperty $msg "Importance"
            Priority = Get-SafeProperty $msg "Priority"
            MessageClass = Get-SafeProperty $msg "MessageClass"
            InternetMessageId = Get-SafeProperty $msg "InternetMessageId"
            ConversationTopic = Get-SafeProperty $msg "ConversationTopic"
            
            Unread = Get-SafeProperty $msg "UnRead"
            Sensitivity = Get-SafeProperty $msg "Sensitivity"
            ReadReceiptRequested = Get-SafeProperty $msg "ReadReceiptRequested"
            DeliveryReceiptRequested = Get-SafeProperty $msg "DeliveryReceiptRequested"
            
            Attachments = @()
            
            Metadata = @{
                FileName = Split-Path -Leaf $FilePath
                FilePath = $FilePath
                FileSize = (Get-Item $FilePath).Length
                ProcessedAt = Get-Date -Format "yyyy-MM-ddTHH:mm:ss.fffZ"
                ProcessorVersion = "1.0"
            }
        }
        
        # Extract recipients
        Write-Verbose "Extracting recipients..."
        if ($msg.Recipients) {
            foreach ($recipient in $msg.Recipients) {
                $recipientData = @{
                    Name = Get-SafeProperty $recipient "Name"
                    Email = Get-SafeProperty $recipient "Address"
                    Type = Get-SafeProperty $recipient "Type"
                }
                
                $recipientType = Get-SafeProperty $recipient "Type"
                switch ($recipientType) {
                    1 { $emailData.Recipients.To += $recipientData }
                    2 { $emailData.Recipients.CC += $recipientData }
                    3 { $emailData.Recipients.BCC += $recipientData }
                    default { $emailData.Recipients.To += $recipientData }
                }
            }
        }
        
        # Extract attachments if requested
        if ($IncludeAttachments -and $msg.Attachments) {
            Write-Verbose "Extracting attachment information..."
            foreach ($attachment in $msg.Attachments) {
                $attachmentData = @{
                    FileName = Get-SafeProperty $attachment "FileName"
                    DisplayName = Get-SafeProperty $attachment "DisplayName"
                    Size = Get-SafeProperty $attachment "Size"
                    Type = Get-SafeProperty $attachment "Type"
                    Position = Get-SafeProperty $attachment "Position"
                    Index = Get-SafeProperty $attachment "Index"
                }
                
                $emailData.Attachments += $attachmentData
            }
        }
        
        # Process content for better text extraction if requested
        if ($ExtractText) {
            $emailData = Add-ProcessedContent -EmailData $emailData
        }
        
        # Add processing statistics
        $emailData.Metadata.AttachmentCount = $emailData.Attachments.Count
        $emailData.Metadata.RecipientCount = ($emailData.Recipients.To.Count + $emailData.Recipients.CC.Count + $emailData.Recipients.BCC.Count)
        $emailData.Metadata.BodyLength = if ($emailData.Body) { $emailData.Body.Length } else { 0 }
        $emailData.Metadata.HTMLBodyLength = if ($emailData.HTMLBody) { $emailData.HTMLBody.Length } else { 0 }
        
        Write-Verbose "Successfully processed MSG file: $($emailData.Metadata.FileName)"
        return $emailData
        
    } catch {
        $errorMessage = "Failed to process MSG file '$FilePath': $($_.Exception.Message)"
        Write-Error $errorMessage
        throw $errorMessage
        
    } finally {
        if ($msg) {
            try {
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($msg) | Out-Null
            } catch {
                Write-Verbose "Warning: Failed to release MSG COM object"
            }
            $msg = $null
        }
        
        if ($outlook) {
            try {
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($outlook) | Out-Null
            } catch {
                Write-Verbose "Warning: Failed to release Outlook COM object"
            }
            $outlook = $null
        }
        
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}

function Get-SafeProperty {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        $ComObject,
        
        [Parameter(Mandatory=$true)]
        [string]$PropertyName
    )
    
    try {
        if ($ComObject -and $ComObject.PSObject.Properties[$PropertyName]) {
            $value = $ComObject.$PropertyName
            
            if ($value -is [System.__ComObject]) {
                return $value.ToString()
            }
            
            if ($value -is [DateTime]) {
                return $value.ToString("yyyy-MM-ddTHH:mm:ss.fffZ")
            }
            
            if ($value -eq [System.DBNull]::Value -or $value -eq $null) {
                return $null
            }
            
            return $value
        }
        
        return $null
        
    } catch {
        Write-Verbose "Warning: Failed to get property '$PropertyName': $($_.Exception.Message)"
        return $null
    }
}

function Add-ProcessedContent {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$EmailData
    )
    
    try {
        $EmailData.ProcessedContent = @{
            PlainText = ""
            CleanHTML = ""
            ExtractedText = ""
            WordCount = 0
            Language = $null
            HasImages = $false
            HasLinks = $false
            ContentType = "Unknown"
        }
        
        $primaryContent = ""
        if ($EmailData.HTMLBody -and $EmailData.HTMLBody.Trim() -ne "") {
            $primaryContent = $EmailData.HTMLBody
            $EmailData.ProcessedContent.ContentType = "HTML"
        } elseif ($EmailData.Body -and $EmailData.Body.Trim() -ne "") {
            $primaryContent = $EmailData.Body
            $EmailData.ProcessedContent.ContentType = "PlainText"
        }
        
        if ($primaryContent) {
            if ($EmailData.ProcessedContent.ContentType -eq "HTML") {
                $EmailData.ProcessedContent.PlainText = ConvertFrom-Html -HtmlContent $primaryContent
                $EmailData.ProcessedContent.CleanHTML = Clean-HtmlContent -HtmlContent $primaryContent
                
                $EmailData.ProcessedContent.HasImages = $primaryContent -match '<img\s+[^>]*>'
                $EmailData.ProcessedContent.HasLinks = $primaryContent -match '<a\s+[^>]*href'
                
            } else {
                $EmailData.ProcessedContent.PlainText = $primaryContent
            }
            
            $EmailData.ProcessedContent.ExtractedText = if ($EmailData.ProcessedContent.PlainText) {
                $EmailData.ProcessedContent.PlainText
            } else {
                $EmailData.Body
            }
            
            if ($EmailData.ProcessedContent.ExtractedText) {
                $words = ($EmailData.ProcessedContent.ExtractedText -split '\s+' | Where-Object { $_ -ne "" })
                $EmailData.ProcessedContent.WordCount = $words.Count
            }
            
            $EmailData.ProcessedContent.Language = Get-ContentLanguage -Content $EmailData.ProcessedContent.ExtractedText
        }
        
        return $EmailData
        
    } catch {
        Write-Verbose "Warning: Failed to add processed content: $($_.Exception.Message)"
        return $EmailData
    }
}

function ConvertFrom-Html {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$HtmlContent
    )
    
    try {
        if (-not $HtmlContent) {
            return ""
        }
        
        $text = $HtmlContent -replace '<!--.*?-->', ''
        $text = $text -replace '<script\b[^<]*(?:(?!<\/script>)<[^<]*)*<\/script>', ''
        $text = $text -replace '<style\b[^<]*(?:(?!<\/style>)<[^<]*)*<\/style>', ''
        
        $text = $text -replace '<br\s*/?>', "`n"
        $text = $text -replace '</?p[^>]*>', "`n`n"
        $text = $text -replace '</?div[^>]*>', "`n"
        
        $text = $text -replace '<li[^>]*>', "`n• "
        $text = $text -replace '</li>', ""
        
        $text = $text -replace '<[^>]*>', ''
        
        try {
            $text = [System.Web.HttpUtility]::HtmlDecode($text)
        } catch {
            $text = $text -replace '&amp;', '&'
            $text = $text -replace '&lt;', '<'
            $text = $text -replace '&gt;', '>'
            $text = $text -replace '&quot;', '"'
            $text = $text -replace '&#39;', "'"
            $text = $text -replace '&nbsp;', ' '
        }
        
        $text = $text -replace '\s+', ' '
        $text = $text -replace '\n\s*\n', "`n`n"
        $text = $text.Trim()
        
        return $text
        
    } catch {
        Write-Verbose "Warning: Failed to convert HTML to text: $($_.Exception.Message)"
        return $HtmlContent
    }
}

function Clean-HtmlContent {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$HtmlContent
    )
    
    try {
        if (-not $HtmlContent) {
            return ""
        }
        
        $cleanHtml = $HtmlContent -replace '<script\b[^<]*(?:(?!<\/script>)<[^<]*)*<\/script>', ''
        $cleanHtml = $cleanHtml -replace '<object\b[^<]*(?:(?!<\/object>)<[^<]*)*<\/object>', ''
        $cleanHtml = $cleanHtml -replace '<embed[^>]*>', ''
        $cleanHtml = $cleanHtml -replace '<applet\b[^<]*(?:(?!<\/applet>)<[^<]*)*<\/applet>', ''
        
        # Fixed regex pattern for tracking pixels - using backticks to escape quotes
        $cleanHtml = $cleanHtml -replace '<img[^>]*width\s*=\s*["`'']*1["`'']*[^>]*height\s*=\s*["`'']*1["`'']*[^>]*>', ''
        
        $cleanHtml = $cleanHtml -replace 'src\s*=\s*["`''][^"`'']*["`'']', 'src="#"'
        
        # Fixed regex pattern for event handlers
        $cleanHtml = $cleanHtml -replace '\s+on\w+\s*=\s*["`''][^"`'']*["`'']', ''
        
        return $cleanHtml
        
    } catch {
        Write-Verbose "Warning: Failed to clean HTML content: $($_.Exception.Message)"
        return $HtmlContent
    }
}

function Get-ContentLanguage {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [AllowEmptyString()]
        [string]$Content
    )
    
    try {
        if (-not $Content -or $Content.Length -lt 10) {
            return "Unknown"
        }
        
        $content = $Content.ToLower()
        
        $englishWords = @("the", "and", "that", "have", "for", "not", "with", "you", "this", "but", "his", "from", "they")
        $englishCount = 0
        foreach ($word in $englishWords) {
            if ($content -match "\b$word\b") {
                $englishCount++
            }
        }
        
        $spanishWords = @("que", "de", "no", "la", "el", "en", "es", "se", "lo", "le", "da", "su", "por", "son")
        $spanishCount = 0
        foreach ($word in $spanishWords) {
            if ($content -match "\b$word\b") {
                $spanishCount++
            }
        }
        
        $frenchWords = @("le", "de", "et", "à", "un", "il", "être", "et", "en", "avoir", "que", "pour", "dans", "ce")
        $frenchCount = 0
        foreach ($word in $frenchWords) {
            if ($content -match "\b$word\b") {
                $frenchCount++
            }
        }
        
        $maxCount = [Math]::Max($englishCount, [Math]::Max($spanishCount, $frenchCount))
        
        if ($maxCount -ge 2) {
            if ($englishCount -eq $maxCount) { return "English" }
            if ($spanishCount -eq $maxCount) { return "Spanish" }
            if ($frenchCount -eq $maxCount) { return "French" }
        }
        
        return "Unknown"
        
    } catch {
        Write-Verbose "Warning: Failed to detect language: $($_.Exception.Message)"
        return "Unknown"
    }
}

function Get-MsgMetadata {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$FilePath
    )
    
    if (-not (Test-Path $FilePath)) {
        throw "MSG file not found: $FilePath"
    }
    
    $outlook = $null
    $msg = $null
    
    try {
        $outlook = New-Object -ComObject Outlook.Application
        $namespace = $outlook.GetNamespace("MAPI")
        $msg = $namespace.OpenSharedItem($FilePath)
        
        if (-not $msg) {
            throw "Failed to open MSG file"
        }
        
        $metadata = @{
            FileName = Split-Path -Leaf $FilePath
            FilePath = $FilePath
            FileSize = (Get-Item $FilePath).Length
            Subject = Get-SafeProperty $msg "Subject"
            SenderName = Get-SafeProperty $msg "SenderName"
            SenderEmail = Get-SafeProperty $msg "SenderEmailAddress"
            ReceivedTime = Get-SafeProperty $msg "ReceivedTime"
            SentOn = Get-SafeProperty $msg "SentOn"
            Size = Get-SafeProperty $msg "Size"
            AttachmentCount = if ($msg.Attachments) { $msg.Attachments.Count } else { 0 }
            MessageClass = Get-SafeProperty $msg "MessageClass"
            Importance = Get-SafeProperty $msg "Importance"
            HasAttachments = if ($msg.Attachments) { $msg.Attachments.Count -gt 0 } else { $false }
            IsRead = -not (Get-SafeProperty $msg "UnRead")
            ExtractedAt = Get-Date -Format "yyyy-MM-ddTHH:mm:ss.fffZ"
        }
        
        return $metadata
        
    } catch {
        throw "Failed to extract MSG metadata: $($_.Exception.Message)"
        
    } finally {
        if ($msg) {
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($msg) | Out-Null
        }
        if ($outlook) {
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($outlook) | Out-Null
        }
        [System.GC]::Collect()
    }
}

function Extract-MsgAttachments {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$FilePath,
        
        [Parameter(Mandatory=$false)]
        [string]$OutputPath
    )
    
    if (-not (Test-Path $FilePath)) {
        throw "MSG file not found: $FilePath"
    }
    
    $outlook = $null
    $msg = $null
    
    try {
        $outlook = New-Object -ComObject Outlook.Application
        $namespace = $outlook.GetNamespace("MAPI")
        $msg = $namespace.OpenSharedItem($FilePath)
        
        if (-not $msg) {
            throw "Failed to open MSG file"
        }
        
        $attachments = @()
        
        if ($msg.Attachments -and $msg.Attachments.Count -gt 0) {
            for ($i = 1; $i -le $msg.Attachments.Count; $i++) {
                $attachment = $msg.Attachments.Item($i)
                
                $attachmentInfo = @{
                    Index = $i
                    FileName = Get-SafeProperty $attachment "FileName"
                    DisplayName = Get-SafeProperty $attachment "DisplayName"
                    Size = Get-SafeProperty $attachment "Size"
                    Type = Get-SafeProperty $attachment "Type"
                    Position = Get-SafeProperty $attachment "Position"
                    PathName = Get-SafeProperty $attachment "PathName"
                }
                
                if ($OutputPath -and (Test-Path $OutputPath)) {
                    try {
                        $safeName = $attachmentInfo.FileName -replace '[<>:"/\\|?*]', '_'
                        $attachmentPath = Join-Path $OutputPath $safeName
                        $attachment.SaveAsFile($attachmentPath)
                        $attachmentInfo.ExtractedPath = $attachmentPath
                        $attachmentInfo.Extracted = $true
                    } catch {
                        $attachmentInfo.Extracted = $false
                        $attachmentInfo.ExtractionError = $_.Exception.Message
                    }
                }
                
                $attachments += $attachmentInfo
            }
        }
        
        return $attachments
        
    } catch {
        throw "Failed to extract attachments: $($_.Exception.Message)"
        
    } finally {
        if ($msg) {
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($msg) | Out-Null
        }
        if ($outlook) {
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($outlook) | Out-Null
        }
        [System.GC]::Collect()
    }
}

function Test-MsgFile {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$FilePath
    )
    
    try {
        if (-not (Test-Path $FilePath)) {
            Write-Verbose "File not found: $FilePath"
            return $false
        }
        
        $extension = [System.IO.Path]::GetExtension($FilePath).ToLower()
        if ($extension -ne ".msg") {
            Write-Verbose "File does not have .msg extension: $extension"
            return $false
        }
        
        $fileInfo = Get-Item $FilePath
        if ($fileInfo.Length -lt 1024) {
            Write-Verbose "File too small to be valid MSG: $($fileInfo.Length) bytes"
            return $false
        }
        
        try {
            $bytes = [System.IO.File]::ReadAllBytes($FilePath)
            if ($bytes.Length -lt 8) {
                return $false
            }
            
            return $true
            
        } catch {
            Write-Verbose "Failed to read file header: $($_.Exception.Message)"
            return $false
        }
        
    } catch {
        Write-Verbose "Error testing MSG file: $($_.Exception.Message)"
        return $false
    }
}

Write-Verbose "MsgProcessor module loaded successfully"

try {
    Add-Type -AssemblyName System.Web
} catch {
    Write-Verbose "Warning: System.Web assembly not available - HTML decoding may be limited"
}