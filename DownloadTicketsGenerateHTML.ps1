# Define Zendesk credentials and API token
$ticketsFolder = "f:\zendesk 2.1"
$subdomain = "subdomain"
$email = "example@example.com"
$api_token = "apikeyhere"
$generatepdfs = $true
$wkhtmltopdfPath = "C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe"
$base64AuthInfo = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes("$email/token:$api_token"))

# Function to download a file
function Download-File {
    param (
        [string]$url,
        [string]$output
    )
    $progresspreference = 'SilentlyContinue'
    Invoke-WebRequest -Uri $url -OutFile $output -UseBasicParsing
}

# Function to sanitize filenames
function Sanitize-Filename {
    param (
        [string]$filename
    )
    $invalidChars = [System.IO.Path]::GetInvalidFileNameChars() + [char[]]":?/"
    foreach ($char in $invalidChars) {
        $filename = $filename -replace [regex]::Escape($char), ""
    }
    return $filename
}

# Function to fetch user details
function Get-User {
    param (
        [int64]$userId
    )
    $userUrl = "https://$subdomain.zendesk.com/api/v2/users/$userId.json"
    try {
        $userResponse = Invoke-RestMethod -Uri $userUrl -Headers @{Authorization=("Basic {0}" -f $base64AuthInfo)}
        return $userResponse.user
    } catch {
        Write-Host "Failed to retrieve user with ID: $userId"
        return $null
    }
}

# Create a folder for tickets
if (-Not (Test-Path -Path $ticketsFolder)) {
    New-Item -Path $ticketsFolder -ItemType Directory
}

# Function to process tickets
function Process-Tickets {
    param (
        [array]$tickets
    )

    foreach ($ticket in $tickets) {
        $ticketFolder = Join-Path -Path $ticketsFolder -ChildPath "Ticket_$($ticket.id)"
        if (Test-Path -Path $ticketFolder) {
            Write-Host "Skipping Ticket ID: $($ticket.id) as it already exists"
            continue
        }

        Write-Host "Processing Ticket ID: $($ticket.id)"

        # Create a folder for each ticket
        $ticketFolder = Join-Path -Path $ticketsFolder -ChildPath "Ticket_$($ticket.id)"
        if (-Not (Test-Path -Path $ticketFolder)) {
            New-Item -Path $ticketFolder -ItemType Directory
        }

        # Save ticket information
        $ticketFile = Join-Path -Path $ticketFolder -ChildPath "ticket_$($ticket.id).json"
        $ticket | ConvertTo-Json -depth 10 | Out-File -FilePath $ticketFile

        # Fetch comments for the ticket
        $commentsUrl = "https://$subdomain.zendesk.com/api/v2/tickets/$($ticket.id)/comments.json"
        
        # Retry mechanism for API rate limits
        $retryCount = 0
        $maxRetries = 5
        $waitTimeInSeconds = 30

        do {
            try {
                $commentsResponse = Invoke-RestMethod -Uri $commentsUrl -Headers @{Authorization=("Basic {0}" -f $base64AuthInfo)}
                $comments = $commentsResponse.comments
                $success = $true
            } catch {
                if ($_.Exception.Response.StatusCode.Value__ -eq 429) {
                    Write-Host "API rate limit exceeded. Waiting for $waitTimeInSeconds seconds before retrying..."
                    Start-Sleep -Seconds $waitTimeInSeconds
                    $retryCount++
                    $success = $false
                } else {
                    throw $_
                }
            }
        } while (-not $success -and $retryCount -lt $maxRetries)

        if (-not $success) {
            Write-Host "Failed to retrieve comments for Ticket ID: $($ticket.id) after $maxRetries retries."
            continue
        }
        $htmlContent = "<html><head><meta charset='UTF-8'></head><body style='font-family: Courier, monospace;'>"
        $htmlContent += "<h1>Ticket ID: $($ticket.id)</h1>"
        $htmlContent += "<p><strong>Subject:</strong> $($ticket.subject)</p>"
        $htmlContent += "<p><strong>Description:</strong> $($ticket.description)</p>"

        foreach ($comment in $comments) {
            # Save comment information
            $commentFile = Join-Path -Path $ticketFolder -ChildPath "comment_$($comment.id).json"
            $comment | ConvertTo-Json -depth 10 | Out-File -FilePath $commentFile

            # Fetch author information
            $author = Get-User -userId $comment.author_id
            $authorName = if ($author) { $author.name } else { "Unknown" }            

            # Add comment to HTML content
            $htmlContent += "<hr><p><strong>Comment ID:</strong> $($comment.id)</p>"
            $htmlContent += "<p><strong>Author:</strong> $($authorName)</p>"
            $htmlContent += "<p><strong>Created At:</strong> $($comment.created_at)</p>"
            $htmlContent += "<p>$($comment.body)</p>"

            # Download attachments
            foreach ($attachment in $comment.attachments) {
                $attachmentUrl = $attachment.content_url
                $sanitizedFilename = Sanitize-Filename $attachment.file_name
                $attachmentPath = Join-Path -Path $ticketFolder -ChildPath $sanitizedFilename
                Download-File -url $attachmentUrl -output $attachmentPath
                $htmlContent += "<p><a href='$sanitizedFilename'>Download Attachment: $sanitizedFilename</a></p>"
            }

            # Download inline images
            $inlineImagePattern = '!\[.*?\]\((.*?)\)'
            if ($comment.body -match $inlineImagePattern) {
                $matches = $matches[1] | ForEach-Object { $_ -replace '\?.*$', '' }
                foreach ($inlineImageUrl in $matches) {
                    $imageFileName = $inlineImageUrl.Split("/")[-1]
                    $sanitizedImageFileName = Sanitize-Filename $imageFileName
                    $imageFilePath = Join-Path -Path $ticketFolder -ChildPath $sanitizedImageFileName
                    Download-File -url $inlineImageUrl -output $imageFilePath
                    $comment.body = $comment.body -replace [regex]::Escape($inlineImageUrl), $sanitizedImageFileName
                }
            }
        }

        $htmlContent += "</body></html>"

        if ($generatepdfs -eq $true) {
            # Save HTML content to a file
            $htmlFile = Join-Path -Path $ticketFolder -ChildPath "ticket_$($ticket.id).html"
            $htmlContent | Out-File -FilePath $htmlFile

            # Convert HTML to PDF using wkhtmltopdf
            $pdfFile = Join-Path -Path $ticketFolder -ChildPath "ticket_$($ticket.id).pdf"

            Start-Process -FilePath $wkhtmltopdfPath -ArgumentList "`"$htmlFile`" `"$pdfFile`"" -NoNewWindow -Wait
        }

        Write-Host "Finished Processing Ticket ID: $($ticket.id)"
    }
}
# Incremental API endpoint - tickets since epoch time

$zendeskApiUrl = "https://$subdomain.zendesk.com/api/v2/incremental/tickets.json?start_time=0"

do {
    # Fetch tickets
    $response = Invoke-RestMethod -Uri $zendeskApiUrl -Headers @{Authorization=("Basic {0}" -f $base64AuthInfo)}
    $tickets = $response.tickets

    # Process the fetched tickets
    Process-Tickets -tickets $tickets

    # Update the API URL to the next page
    $zendeskApiUrl = $response.next_page

} while ($zendeskApiUrl -ne $null)

Write-Host "Tickets and attachments have been successfully downloaded and organized."
