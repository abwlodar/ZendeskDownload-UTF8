# Define Zendesk credentials and API token
$ticketsFolder = ".\" # By default will download to the current folder.
$subdomain = "subdomain"
$email = "test@test.com"
$api_token = "api_token_goes_here"
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
        $htmlContent = "<html><body>"
        $htmlContent += "<h1>Ticket ID: $($ticket.id)</h1>"
        $htmlContent += "<p><strong>Subject:</strong> $($ticket.subject)</p>"
        $htmlContent += "<p><strong>Description:</strong> $($ticket.description)</p>"

        foreach ($comment in $comments) {
            # Save comment information
            $commentFile = Join-Path -Path $ticketFolder -ChildPath "comment_$($comment.id).json"
            $comment | ConvertTo-Json -depth 10 | Out-File -FilePath $commentFile

            # Add comment to HTML content
            $htmlContent += "<hr><p><strong>Comment ID:</strong> $($comment.id)</p>"
            $htmlContent += "<p><strong>Author:</strong> $($comment.author_id)</p>"
            $htmlContent += "<p><strong>Created At:</strong> $($comment.created_at)</p>"
            $htmlContent += "<p>$($comment.body)</p>"

            # Download attachments
            foreach ($attachment in $comment.attachments) {
                $attachmentUrl = $attachment.content_url
                $attachmentPath = Join-Path -Path $ticketFolder -ChildPath $attachment.file_name
                Download-File -url $attachmentUrl -output $attachmentPath
                $htmlContent += "<p><a href='$attachment.file_name'>Download Attachment: $($attachment.file_name)</a></p>"
            }
        }

        $htmlContent += "</body></html>"
if ($generatepdfs -eq $true)
{
        # Save HTML content to a file
        $htmlFile = Join-Path -Path $ticketFolder -ChildPath "ticket_$($ticket.id).html"
        $htmlContent | Out-File -FilePath $htmlFile

        # Convert HTML to PDF using wkhtmltopdf
        $pdfFile = Join-Path -Path $ticketFolder -ChildPath "ticket_$($ticket.id).pdf"


            Start-Process -FilePath $wkhtmltopdfPath -ArgumentList "$htmlFile $pdfFile" -NoNewWindow -Wait

    }
        Write-Host "Finished Processing Ticket ID: $($ticket.id)"
    }

}
# Initial API endpoint
$zendeskApiUrl = "https://$subdomain.zendesk.com/api/v2/tickets.json"

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



# HTML template for the index file
$htmlHeader = @"
<html>
<head>
    <title>Zendesk Tickets</title>
    <style>
        body { font-family: Arial, sans-serif; }
        .ticket { margin-bottom: 10px; border: 1px solid #ccc; padding: 10px; }
        .ticket h2 { margin: 0; cursor: pointer; }
        .ticket-details { display: none; padding: 10px; border-top: 1px solid #ccc; }
        .attachment { margin: 5px 0; }
        #search { margin-bottom: 20px; }
        .paperclip { color: #000; margin-left: 5px; }
        html { overflow-y: scroll; }
    </style>
    <script>
        function loadTicketDetails(ticketId) {
            var detailsDiv = document.getElementById('details-' + ticketId);
            if (detailsDiv.innerHTML === '') {
                fetch('Ticket_' + ticketId + '/ticket_' + ticketId + '.html')
                .then(response => response.text())
                .then(data => detailsDiv.innerHTML = data)
                .catch(error => console.error('Error loading ticket details:', error));
            }
            toggleDetails(ticketId);
        }

        function toggleDetails(ticketId) {
            var details = document.getElementById('details-' + ticketId);
            if (details.style.display === 'none' || details.style.display === '') {
                details.style.display = 'block';
            } else {
                details.style.display = 'none';
            }
        }

        function searchTickets() {
            var input, filter, tickets, h2, i, txtValue;
            input = document.getElementById('search');
            filter = input.value.toLowerCase();
            tickets = document.getElementsByClassName('ticket');
            var filterAttachment = document.getElementById('filterAttachment').checked;

            for (i = 0; i < tickets.length; i++) {
                h2 = tickets[i].getElementsByTagName('h2')[0];
                txtValue = h2.textContent || h2.innerText;

                // Check if the ticket contains the filter text
                var matchesSearch = txtValue.toLowerCase().indexOf(filter) > -1;

                // Check if the ticket contains the 📎 character
                var hasAttachment = tickets[i].querySelector('.paperclip') !== null;

                // Apply filter logic
                if (matchesSearch && (!filterAttachment || (filterAttachment && hasAttachment))) {
                    tickets[i].style.display = '';
                } else {
                    tickets[i].style.display = 'none';
        }
    }
}

    </script>
</head>
<body>
    <h1>Zendesk Tickets</h1>
    <input type="text" id="search" onkeyup="searchTickets()" placeholder="Search for tickets...">
    <input type="checkbox" id="filterAttachment" onchange="searchTickets()" />
    <label for="filterAttachment">Filter by Attachment 📎</label>

"@

$htmlFooter = @"
</body>
</html>
"@

# Initialize the HTML content with the header
$htmlContent = $htmlHeader

# Function to read JSON files and convert to objects
function Read-JsonFile {
    param (
        [string]$filePath
    )
    Get-Content -Path $filePath | Out-String | ConvertFrom-Json
}

# Enumerate tickets and sort numerically by ticket ID
$ticketFolders = Get-ChildItem -Path $ticketsFolder -Directory | Sort-Object { [int]($_.Name -replace 'Ticket_', '') }

# Total number of tickets to process
$totalTickets = $ticketFolders.Count
$currentTicket = 0

# Generate HTML content with progress bar
foreach ($ticketFolder in $ticketFolders) {
    $ticketId = $ticketFolder.Name -replace 'Ticket_', ''
    $ticketFile = Join-Path -Path $ticketFolder.FullName -ChildPath "ticket_$ticketId.json"
    $ticket = Read-JsonFile -filePath $ticketFile

    # Check for attachments
    $hasAttachments = $false
    foreach ($commentFile in Get-ChildItem -Path $ticketFolder.FullName -Filter "comment_*.json") {
        $comment = Read-JsonFile -filePath $commentFile.FullName
        if ($comment.attachments.Count -gt 0) {
            $hasAttachments = $true
            break
        }
    }

    # Add paperclip symbol if there are attachments
    $paperclip = ""
    if ($hasAttachments) {
        $paperclip = "<span class='paperclip'>&#128206;</span>"  # Unicode for paperclip symbol
    }

    # Add ticket link
    $htmlContent += "<div class='ticket'>"
    $htmlContent += "<h2 onclick='loadTicketDetails($($ticket.id))'>Ticket ID: $($ticket.id) - $($ticket.subject) $paperclip</h2>"
    $htmlContent += "<div class='ticket-details' id='details-$($ticket.id)'></div>"
    $htmlContent += "</div>"

    # Generate individual ticket HTML file
    $ticketHtmlContent = "<div>"
    $ticketHtmlContent += "<p><strong>Description:</strong> $($ticket.description)</p>"

    foreach ($commentFile in Get-ChildItem -Path $ticketFolder.FullName -Filter "comment_*.json") {
        $comment = Read-JsonFile -filePath $commentFile.FullName
        $ticketHtmlContent += "<hr><p><strong>Comment ID:</strong> $($comment.id)</p>"
        $ticketHtmlContent += "<p><strong>Author:</strong> $($comment.author_id)</p>"
        $ticketHtmlContent += "<p><strong>Created At:</strong> $($comment.created_at)</p>"
        $ticketHtmlContent += "<p>$($comment.body)</p>"

        # Add attachments links
        foreach ($attachment in $comment.attachments) {
            $attachmentPath = Join-Path -Path Ticket_$($ticket.id)/ -ChildPath $attachment.file_name
            $ticketHtmlContent += "<p class='attachment'><a href='$attachmentPath'>Download Attachment: $($attachment.file_name)</a></p>"
        }
    }

    # Add PDF link to individual ticket HTML content
    if ($generatepdfs -eq $true)
    {
    $pdfLink = "<p><a href='Ticket_$($ticket.id)/ticket_$($ticket.id).pdf'>Click here to view printable PDF</a></p>"
    $ticketHtmlContent += $pdfLink
    $ticketHtmlContent += "</div>"
    }

    # Save individual ticket HTML file with UTF-8 encoding
    $ticketHtmlFilePath = Join-Path -Path $ticketFolder.FullName -ChildPath "ticket_$ticketId.html"
    [System.IO.File]::WriteAllText($ticketHtmlFilePath, $ticketHtmlContent, [System.Text.Encoding]::UTF8)

    # Update the progress bar
    $currentTicket++
    $percentComplete = [math]::Round(($currentTicket / $totalTickets) * 100)
    Write-Progress -Activity "Generating HTML" -Status "Processing ticket $currentTicket of $totalTickets" -PercentComplete $percentComplete
}

# Close the HTML content with the footer
$htmlContent += $htmlFooter

# Define the output HTML file path
$htmlFilePath = Join-Path -Path $ticketsFolder -ChildPath "index.html"

# Save the HTML content to the file with UTF-8 encoding
[System.IO.File]::WriteAllText($htmlFilePath, $htmlContent, [System.Text.Encoding]::UTF8)

Write-Host "HTML index has been successfully generated at $htmlFilePath"
