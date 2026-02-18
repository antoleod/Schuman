Set-StrictMode -Version Latest

$moduleRoot = Split-Path -Parent $MyInvocation.MyCommand.Path

. (Join-Path $moduleRoot 'Core\Bootstrap.ps1')
. (Join-Path $moduleRoot 'Core\Logging.ps1')
. (Join-Path $moduleRoot 'Core\Paths.ps1')
. (Join-Path $moduleRoot 'Core\Text.ps1')
. (Join-Path $moduleRoot 'Integrations\Excel.ps1')
. (Join-Path $moduleRoot 'Integrations\WebView2ServiceNow.ps1')
. (Join-Path $moduleRoot 'Workflows\TicketExport.ps1')
. (Join-Path $moduleRoot 'Workflows\Dashboard.ps1')
. (Join-Path $moduleRoot 'Workflows\Documents.ps1')
