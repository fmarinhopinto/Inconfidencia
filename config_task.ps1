$task = Get-ScheduledTask -TaskName 'Atualizar_Dashboard_Inconfidencia'
Write-Host "--- ACTION ---"
$task.Actions | Format-List
Write-Host "--- TRIGGERS ---"
$task.Triggers | Format-List
Write-Host "--- SETTINGS ---"
$task.Settings | Format-List
Write-Host "--- PRINCIPAL ---"
$task.Principal | Format-List
