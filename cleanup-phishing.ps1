# ================================
# CONFIGURATION
# ================================

# Inline spam sender list
$senders = @(
    "somespam@gmail.com",
    "anotherspam@gmail.com",
    "yetanotherspam@gmail.com"
)

# Generate timestamped log file name
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$LogPath = "C:\spam_cleanup_$timestamp.csv"

# ================================
# BUILD SEARCH QUERY
# ================================

$SearchQuery = ($senders | ForEach-Object { "From:`"$_`"" }) -join " OR "

Write-Host "SearchQuery:"
Write-Host $SearchQuery
Write-Host ""

# ================================
# CREATE CSV HEADER
# ================================

"Date,Mailbox,DeletedItems" | Out-File $LogPath -Encoding UTF8

# ================================
# GET USER MAILBOXES ONLY
# ================================

$mailboxes = Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox

# ================================
# MAIN LOOP (DN prevents ambiguity)
# ================================

foreach ($mb in $mailboxes) {

    $MailboxDN = $mb.DistinguishedName
    $MailboxAddress = $mb.PrimarySmtpAddress.ToString()

    Write-Host "Processing: $MailboxAddress" -ForegroundColor Cyan

    try {
        $result = Search-Mailbox `
            -Identity $MailboxDN `
            -SearchQuery $SearchQuery `
            -DeleteContent `
            -Force `
            -ErrorAction Stop

        $deleted = $result.ResultItemsCount

        if ($deleted -gt 0) {
            $logLine = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$MailboxAddress,$deleted"
            Add-Content -Path $LogPath -Value $logLine
            Write-Host "Deleted: $deleted" -ForegroundColor Green
        }
        else {
            Write-Host "Nothing found"
        }
    }
    catch {
        Write-Host "Error processing $MailboxAddress" -ForegroundColor Red
    }
}

Write-Host ""
Write-Host "Completed."
Write-Host "Log file: $LogPath"
