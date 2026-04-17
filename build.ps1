# ============================================================
#  build.ps1 - Genere IconScraper.ppam
#  PowerPoint Add-in : recherche d'icones via Iconify
# ============================================================

param(
    [string]$OutputPath = "$PSScriptRoot\IconScraper.ppam"
)

$ErrorActionPreference = "Stop"

Write-Host ""
Write-Host "================================================" -ForegroundColor Cyan
Write-Host "  IconScraper.ppam - Build" -ForegroundColor Cyan
Write-Host "================================================" -ForegroundColor Cyan
Write-Host ""

# Lecture du code VBA depuis les fichiers sources
$moduleCode = Get-Content -Path "$PSScriptRoot\src\module.vba" -Raw -Encoding UTF8

# Activation de l'acces au projet VBA via le registre
$regPath   = "HKCU:\Software\Microsoft\Office\16.0\PowerPoint\Security"
$origValue = $null
try {
    $origValue = (Get-ItemProperty -Path $regPath -Name "AccessVBOM" -ErrorAction Stop).AccessVBOM
} catch {}
Set-ItemProperty -Path $regPath -Name "AccessVBOM" -Value 1 -Type DWord -Force
Write-Host "[1/5] Acces VBA active dans le registre" -ForegroundColor Green

# Fermeture de PowerPoint s'il est deja ouvert (pour prendre en compte le registre)
$running = Get-Process -Name "POWERPNT" -ErrorAction SilentlyContinue
if ($running) {
    Write-Host "      PowerPoint deja ouvert - fermeture en cours..." -ForegroundColor Yellow
    $running | Stop-Process -Force
    Start-Sleep -Seconds 2
}

$pptApp = $null
$pres   = $null

try {
    Write-Host "[2/5] Demarrage de PowerPoint..." -ForegroundColor Yellow
    $pptApp         = New-Object -ComObject PowerPoint.Application
    $pptApp.Visible = [Microsoft.Office.Core.MsoTriState]::msoTrue

    Write-Host "[3/5] Creation de la presentation..." -ForegroundColor Yellow
    $pres   = $pptApp.Presentations.Add([Microsoft.Office.Core.MsoTriState]::msoTrue)
    $vbProj = $pres.VBProject

    if ($null -eq $vbProj) {
        throw "Acces au projet VBA refuse. Dans PowerPoint : Fichier -> Options -> Centre de gestion de la confidentialite -> Parametres des macros -> cocher 'Faire confiance au projet VBA'"
    }

    # Module VBA principal
    $mod      = $vbProj.VBComponents.Add(1)
    $mod.Name = "IconScraper"
    $mod.CodeModule.AddFromString($moduleCode)

    # Sauvegarde en .ppam
    Write-Host "[5/5] Sauvegarde en .ppam..." -ForegroundColor Yellow
    $pres.SaveAs($OutputPath, 30)

    Write-Host ""
    Write-Host "  SUCCES !" -ForegroundColor Green
    Write-Host "  Fichier : $OutputPath" -ForegroundColor White
    Write-Host ""
    Write-Host "  Installation :" -ForegroundColor Cyan
    Write-Host "  Fichier -> Options -> Complements" -ForegroundColor White
    Write-Host "  Gerer : Complements PowerPoint -> OK -> Ajouter" -ForegroundColor White
    Write-Host "  Selectionner : IconScraper.ppam" -ForegroundColor White
    Write-Host ""
    Write-Host "  Utilisation :" -ForegroundColor Cyan
    Write-Host "  Affichage -> Macros -> ShowIconScraper -> Executer" -ForegroundColor White
    Write-Host ""

} catch {
    Write-Host ""
    Write-Host "  ERREUR : $($_.Exception.Message)" -ForegroundColor Red
    Write-Host ""
} finally {
    if ($null -ne $pres)   { try { $pres.Close()  } catch {} }
    if ($null -ne $pptApp) { try { $pptApp.Quit() } catch {} }

    # Restauration du parametre securite VBA
    if ($null -eq $origValue) {
        Remove-ItemProperty -Path $regPath -Name "AccessVBOM" -ErrorAction SilentlyContinue
    } else {
        Set-ItemProperty -Path $regPath -Name "AccessVBOM" -Value $origValue -Type DWord -Force
    }
    Write-Host "  Securite VBA restauree." -ForegroundColor Green
    Write-Host ""
}
