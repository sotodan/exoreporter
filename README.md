# Exchange Online Reporter
  This script connects to Exchange Online, retrieves mailboxes with a specified domain, and gathers detailed information about each mailbox, including size, item count, and hold statuses. It then generates an HTML report and logs the processing user.

## .SYNOPSIS
  Generates a detailed report of mailboxes in an Exchange Online environment.
## .DESCRIPTION
  This script connects to Exchange Online, retrieves mailboxes with a specified domain, and gathers detailed information about each mailbox, including size, item count, and hold statuses. It then generates an HTML report and logs the processing user.
## .PARAMETER domain
    The domain to filter mailboxes (e.g., "@M365DS219944.onmicrosoft.com").
## .INPUTS
  None
## .OUTPUTS
  HTML report stored in C:\temp\ExchangeOnlineMailboxReport.html
  Log file stored in C:\temp\ProcessingUserLog.txt
## .NOTES
  Version:        1.0
  Author:         Daniel Soto
  GitHub:         Sotodan
  Creation Date:  07/14/2024
  Purpose/Change: Initial script development
