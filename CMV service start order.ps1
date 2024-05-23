# Stop services
Stop-Service -Name "MSSQLSERVER" -Force
Stop-Service -Name "MaconomyDaemon" -Force
Stop-Service -Name "CouplingService.w_21_0.PROD.8080" -Force

# Start services
Start-Service -Name "MSSQLSERVER"
Start-Service -Name "MaconomyDaemon"
Start-Service -Name "CouplingService.w_21_0.PROD.8080"