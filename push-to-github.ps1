# push-to-github.ps1 - pusht das lokale Repo zu GitHub.
# Verwendung:
#   .\push-to-github.ps1 -User "deinGitHubName" -Repo "gantt-addin" -Token "ghp_xxx"
#
# Den Token (Personal Access Token) erstellst du unter:
#   GitHub -> Settings -> Developer settings -> Personal access tokens
#   -> Tokens (classic) -> Generate new token -> Scope: "repo"
#
# Das Skript legt das Remote-Repo an (falls noch nicht vorhanden) und pusht main.
param(
  [Parameter(Mandatory = $true)][string]$User,
  [Parameter(Mandatory = $true)][string]$Repo,
  [Parameter(Mandatory = $true)][string]$Token
)

$ErrorActionPreference = "Stop"
$git = "$env:LOCALAPPDATA\MinGit\cmd"
if (Test-Path $git) { $env:Path = "$git;$env:Path" }
Set-Location $PSScriptRoot

# Proxy-Anmeldedaten fuer den Netzwerkzugriff
[System.Net.WebRequest]::DefaultWebProxy.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials

# 1) Repo via GitHub API anlegen (ignoriert Fehler, falls es schon existiert)
try {
  $body = @{ name = $Repo; private = $false; auto_init = $false } | ConvertTo-Json
  Invoke-RestMethod -Method Post -Uri "https://api.github.com/user/repos" `
    -Headers @{ Authorization = "token $Token"; "User-Agent" = $User; Accept = "application/vnd.github+json" } `
    -Body $body -ErrorAction Stop | Out-Null
  Write-Host "Repository '$Repo' angelegt." -ForegroundColor Green
} catch {
  Write-Host "Hinweis: Repo evtl. schon vorhanden - fahre fort." -ForegroundColor Yellow
}

# 2) Remote setzen (Token in der URL fuer die Authentifizierung)
$remoteUrl = "https://$User`:$Token@github.com/$User/$Repo.git"
& git remote remove origin 2>$null
& git remote add origin $remoteUrl

# 3) Push
& git push -u origin main

# 4) Remote-URL ohne Token hinterlegen (Token nicht dauerhaft speichern)
& git remote set-url origin "https://github.com/$User/$Repo.git"

Write-Host ""
Write-Host "Push abgeschlossen." -ForegroundColor Green
Write-Host "Aktiviere jetzt GitHub Pages:" -ForegroundColor Yellow
Write-Host "  Repo -> Settings -> Pages -> Branch 'main', Ordner '/docs' -> Save"
Write-Host "Deine App-URL wird dann:" -ForegroundColor Yellow
Write-Host "  https://$User.github.io/$Repo/taskpane.html"
Write-Host ""
Write-Host "Danach Produktions-Dateien bauen mit:" -ForegroundColor Yellow
Write-Host "  .\make-prod.ps1 -HostUrl `"https://$User.github.io/$Repo`""
