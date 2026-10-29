# Reconfigura a tarefa agendada com parametros robustos
$taskName = "Atualizar_Dashboard_Inconfidencia"
$workDir = "C:\Users\filipe.marinho\Desktop\PROJETO DASH INCONFIDENCIA"
$batPath = "C:\Users\filipe.marinho\Desktop\PROJETO DASH INCONFIDENCIA\automacao_diaria.bat"

# 1. Action com WorkingDirectory explicito
$action = New-ScheduledTaskAction -Execute $batPath -WorkingDirectory $workDir

# 2. Triggers: Diariamente as 07:30 E ao fazer logon
$triggerDaily = New-ScheduledTaskTrigger -Daily -At "07:30:00"
$triggerLogon = New-ScheduledTaskTrigger -AtLogOn -User "$env:USERDOMAIN\$env:USERNAME"

# 3. Settings avancadas
$settings = New-ScheduledTaskSettingsSet `
    -StartWhenAvailable `
    -DontStopIfGoingOnBatteries `
    -WakeToRun `
    -ExecutionTimeLimit (New-TimeSpan -Hours 2) `
    -Priority 5

# 4. Principal
$principal = New-ScheduledTaskPrincipal -UserId "$env:USERDOMAIN\$env:USERNAME" -LogonType Interactive

# 5. Registra a tarefa (substituindo a antiga)
Register-ScheduledTask `
    -TaskName $taskName `
    -Action $action `
    -Trigger @($triggerDaily, $triggerLogon) `
    -Settings $settings `
    -Principal $principal `
    -Force

Write-Host "Tarefa $taskName reconfigurada com sucesso total!"
