# Integrations

## `src/Schuman.Automation/Integrations/Excel.ps1`
- Reads tickets from Excel (`Read-TicketsFromExcel`).
- Writes results (`Write-TicketResultsToExcel`).
- Dashboard search (`Search-DashboardRows`).
- Dashboard row update (`Update-DashboardRow`).
- Strict COM cleanup (release + GC).

## `src/Schuman.Automation/Integrations/WebView2ServiceNow.ps1`
- Initializes WebView2 from Teams Add-in (no extra installation required).
- SSO session management (`New-ServiceNowSession`, `Close-ServiceNowSession`).
- JSONv2 queries (`Invoke-ServiceNowJsonv2Query`).
- Ticket/task retrieval plus label/cache resolution.
- SCTASK state update from UI (`Set-ServiceNowTaskState`).
