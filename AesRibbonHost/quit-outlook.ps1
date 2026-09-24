try {
    $ol = [Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application')
    $ol.Quit()
    Write-Host 'Outlook.Quit sent'
} catch {
    Write-Host ('No running Outlook COM: ' + $_.Exception.Message)
}
