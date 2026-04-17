# ============================================================
#  build.ps1  —  Génère IconScraper.ppam
#  PowerPoint Add-in : recherche d'icônes via Iconify (gratuit)
# ============================================================

param(
    [string]$OutputPath = "$PSScriptRoot\IconScraper.ppam"
)

$ErrorActionPreference = "Stop"

# ── Code du module VBA principal ─────────────────────────────────────────────
$moduleCode = @'
Option Explicit

Private Const ICONIFY_API As String = "https://api.iconify.design"
Private Const MAX_RESULTS As Integer = 60

' Point d'entree : ouvre le formulaire de recherche
Public Sub ShowIconScraper()
    frmIconScraper.Show
End Sub

' Recherche des icones via l'API Iconify
Public Function SearchIcons(query As String, Optional prefix As String = "") As String()
    Dim http As Object
    Dim url As String
    Dim icons() As String

    Set http = CreateObject("MSXML2.XMLHTTP60")
    url = ICONIFY_API & "/search?query=" & EncodeURL(query) & "&limit=" & MAX_RESULTS
    If prefix <> "" Then url = url & "&prefix=" & prefix

    On Error GoTo ErrHandler
    http.Open "GET", url, False
    http.setRequestHeader "Accept", "application/json"
    http.Send

    If http.Status = 200 Then
        icons = ParseIconsJSON(http.responseText)
    Else
        ReDim icons(0): icons(0) = ""
    End If
    SearchIcons = icons
    Exit Function

ErrHandler:
    ReDim icons(0): icons(0) = ""
    SearchIcons = icons
End Function

' Parse le JSON Iconify : extrait le tableau "icons"
Private Function ParseIconsJSON(json As String) As String()
    Dim result() As String
    Dim startPos As Long, endPos As Long

    startPos = InStr(json, """icons"":[")
    If startPos = 0 Then
        ReDim result(0): result(0) = ""
        ParseIconsJSON = result
        Exit Function
    End If

    startPos = startPos + 9
    endPos = InStr(startPos, json, "]")

    Dim iconsStr As String
    iconsStr = Mid(json, startPos, endPos - startPos)
    iconsStr = Replace(iconsStr, """", "")

    result = Split(iconsStr, ",")
    ParseIconsJSON = result
End Function

' Telecharge le SVG et l'insere dans la slide active
Public Sub InsertIcon(iconFullName As String, sizePx As Integer)
    Dim parts() As String
    parts = Split(Trim(iconFullName), ":")
    If UBound(parts) < 1 Then
        MsgBox "Nom invalide : " & iconFullName, vbExclamation, "Icon Scraper"
        Exit Sub
    End If

    Dim prefix As String, iconName As String
    prefix = Trim(parts(0))
    iconName = Trim(parts(1))

    Dim url As String
    url = ICONIFY_API & "/" & prefix & "/" & iconName & ".svg" & _
          "?width=" & sizePx & "&height=" & sizePx & "&color=%23333333"

    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP60")

    On Error GoTo ErrDownload
    http.Open "GET", url, False
    http.Send

    If http.Status <> 200 Then
        MsgBox "Erreur HTTP " & http.Status, vbExclamation, "Icon Scraper"
        Exit Sub
    End If

    ' Sauvegarde dans un fichier temporaire
    Dim tempFile As String
    tempFile = Environ("TEMP") & "\iconscrap_" & prefix & "_" & iconName & ".svg"

    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1
    stream.Open
    stream.Write http.responseBody
    stream.SaveToFile tempFile, 2
    stream.Close

    ' Insertion dans la slide courante
    If Application.Presentations.Count = 0 Then
        MsgBox "Aucune presentation ouverte.", vbExclamation, "Icon Scraper"
        Exit Sub
    End If

    Dim slide As Slide
    Set slide = Application.ActivePresentation.Slides( _
        Application.ActiveWindow.View.Slide.SlideIndex)

    Dim sizeInPts As Single
    sizeInPts = sizePx * 0.75

    Dim shp As Shape
    Set shp = slide.Shapes.AddPicture( _
        Filename:=tempFile, _
        LinkToFile:=False, _
        SaveWithDocument:=True, _
        Left:=50, Top:=50, _
        Width:=sizeInPts, Height:=sizeInPts)

    shp.Name = "Icon_" & prefix & "_" & iconName

    On Error Resume Next
    Kill tempFile
    On Error GoTo 0
    Exit Sub

ErrDownload:
    MsgBox "Erreur reseau : " & Err.Description, vbExclamation, "Icon Scraper"
End Sub

' Encodage URL basique
Private Function EncodeURL(s As String) As String
    Dim i As Integer, c As String, result As String
    For i = 1 To Len(s)
        c = Mid(s, i, 1)
        Select Case c
            Case "A" To "Z", "a" To "z", "0" To "9", "-", "_", "."
                result = result & c
            Case " "
                result = result & "+"
            Case Else
                result = result & "%" & Right("0" & Hex(Asc(c)), 2)
        End Select
    Next i
    EncodeURL = result
End Function
'@

# ── Code du UserForm ─────────────────────────────────────────────────────────
$formCode = @'
Option Explicit

Private Sub UserForm_Initialize()
    With cboCollection
        .AddItem "Toutes les collections"
        .AddItem "mdi  -  Material Design Icons"
        .AddItem "fa   -  Font Awesome"
        .AddItem "logos  -  Logos de marques"
        .AddItem "fluent  -  Microsoft Fluent"
        .AddItem "ph  -  Phosphor Icons"
        .AddItem "tabler  -  Tabler Icons"
        .AddItem "heroicons  -  Heroicons"
        .ListIndex = 0
    End With

    With cboSize
        .AddItem "32"
        .AddItem "64"
        .AddItem "128"
        .AddItem "256"
        .ListIndex = 2
    End With
End Sub

Private Sub btnSearch_Click()
    Dim query As String
    query = Trim(txtSearch.Text)
    If query = "" Then
        lblStatus.Caption = "Saisis un mot-cle."
        Exit Sub
    End If

    Dim prefix As String
    Select Case cboCollection.ListIndex
        Case 1: prefix = "mdi"
        Case 2: prefix = "fa"
        Case 3: prefix = "logos"
        Case 4: prefix = "fluent"
        Case 5: prefix = "ph"
        Case 6: prefix = "tabler"
        Case 7: prefix = "heroicons"
        Case Else: prefix = ""
    End Select

    lblStatus.Caption = "Recherche en cours..."
    lstResults.Clear
    Me.Repaint

    On Error GoTo ErrHandler
    Dim icons() As String
    icons = SearchIcons(query, prefix)

    Dim i As Integer, count As Integer
    count = 0
    For i = 0 To UBound(icons)
        If Trim(icons(i)) <> "" Then
            lstResults.AddItem Trim(icons(i))
            count = count + 1
        End If
    Next i

    If count = 0 Then
        lblStatus.Caption = "Aucun resultat."
    Else
        lblStatus.Caption = count & " icone(s) trouvee(s)."
    End If
    Exit Sub

ErrHandler:
    lblStatus.Caption = "Erreur : " & Err.Description
End Sub

Private Sub btnInsert_Click()
    If lstResults.ListIndex = -1 Then
        lblStatus.Caption = "Selectionne une icone d'abord."
        Exit Sub
    End If

    lblStatus.Caption = "Insertion en cours..."
    Me.Repaint

    On Error GoTo ErrHandler
    InsertIcon lstResults.Text, CInt(cboSize.Text)
    lblStatus.Caption = "Insere : " & lstResults.Text
    Exit Sub

ErrHandler:
    lblStatus.Caption = "Erreur : " & Err.Description
End Sub

Private Sub txtSearch_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then btnSearch_Click
End Sub
'@

# ── Build ─────────────────────────────────────────────────────────────────────

Write-Host ""
Write-Host "================================================" -ForegroundColor Cyan
Write-Host "   IconScraper.ppam  —  Build" -ForegroundColor Cyan
Write-Host "================================================" -ForegroundColor Cyan
Write-Host ""

# Activation temporaire de l'acces au projet VBA via registre
$regPath = "HKCU:\Software\Microsoft\Office\16.0\PowerPoint\Security"
$origValue = $null
try {
    $origValue = (Get-ItemProperty -Path $regPath -Name "AccessVBOM" -ErrorAction Stop).AccessVBOM
} catch {}
Set-ItemProperty -Path $regPath -Name "AccessVBOM" -Value 1 -Type DWord -Force
Write-Host "[1/5] Acces VBA active temporairement" -ForegroundColor Green

$pptApp = $null
$pres   = $null

try {
    Write-Host "[2/5] Demarrage de PowerPoint..." -ForegroundColor Yellow
    $pptApp = New-Object -ComObject PowerPoint.Application
    $pptApp.Visible = [Microsoft.Office.Core.MsoTriState]::msoTrue

    Write-Host "[3/5] Creation de la presentation..." -ForegroundColor Yellow
    $pres = $pptApp.Presentations.Add([Microsoft.Office.Core.MsoTriState]::msoTrue)

    $vbProj = $pres.VBProject

    # Module principal
    $mod = $vbProj.VBComponents.Add(1)   # vbext_ct_StdModule
    $mod.Name = "IconScraper"
    $mod.CodeModule.AddFromString($moduleCode)

    Write-Host "[4/5] Creation du formulaire..." -ForegroundColor Yellow

    # UserForm
    $form = $vbProj.VBComponents.Add(3)  # vbext_ct_MSForm
    $form.Name = "frmIconScraper"

    $d = $form.Designer
    $d.Caption         = "Icon Scraper — Iconify"
    $d.Width           = 356
    $d.Height          = 430
    $d.StartUpPosition = 1  # CenterOwner

    $c = $d.Controls

    # Ligne 1 : recherche
    $l1 = $c.Add("Forms.Label.1",         "lblSearch",     $true)
    $l1.Caption = "Rechercher :"; $l1.Left = 8;  $l1.Top = 12; $l1.Width = 72; $l1.Height = 16

    $tx = $c.Add("Forms.TextBox.1",        "txtSearch",     $true)
    $tx.Left = 84; $tx.Top = 9; $tx.Width = 172; $tx.Height = 20

    $bs = $c.Add("Forms.CommandButton.1",  "btnSearch",     $true)
    $bs.Caption = "Chercher"; $bs.Left = 262; $bs.Top = 8; $bs.Width = 76; $bs.Height = 22

    # Ligne 2 : filtre collection
    $l2 = $c.Add("Forms.Label.1",         "lblCol",        $true)
    $l2.Caption = "Collection :"; $l2.Left = 8; $l2.Top = 38; $l2.Width = 72; $l2.Height = 16

    $cb = $c.Add("Forms.ComboBox.1",       "cboCollection", $true)
    $cb.Left = 84; $cb.Top = 36; $cb.Width = 254; $cb.Height = 20; $cb.Style = 2

    # Ligne 3 : liste des resultats
    $lb = $c.Add("Forms.ListBox.1",        "lstResults",    $true)
    $lb.Left = 8; $lb.Top = 64; $lb.Width = 332; $lb.Height = 272

    # Ligne 4 : statut
    $l3 = $c.Add("Forms.Label.1",         "lblStatus",     $true)
    $l3.Caption = "Pret."; $l3.Left = 8; $l3.Top = 344; $l3.Width = 332; $l3.Height = 16

    # Ligne 5 : taille + bouton insertion
    $l4 = $c.Add("Forms.Label.1",         "lblSize",       $true)
    $l4.Caption = "Taille (px) :"; $l4.Left = 8; $l4.Top = 370; $l4.Width = 76; $l4.Height = 16

    $cs = $c.Add("Forms.ComboBox.1",       "cboSize",       $true)
    $cs.Left = 88; $cs.Top = 368; $cs.Width = 60; $cs.Height = 20; $cs.Style = 2

    $bi = $c.Add("Forms.CommandButton.1",  "btnInsert",     $true)
    $bi.Caption = "Inserer dans la slide"
    $bi.Left = 156; $bi.Top = 366; $bi.Width = 184; $bi.Height = 24

    # Code du formulaire
    $form.CodeModule.AddFromString($formCode)

    # Sauvegarde en .ppam
    Write-Host "[5/5] Sauvegarde en .ppam..." -ForegroundColor Yellow
    $pres.SaveAs($OutputPath, 25)   # 25 = ppSaveAsAddin

    Write-Host ""
    Write-Host "  SUCCES !" -ForegroundColor Green
    Write-Host "  Fichier : $OutputPath" -ForegroundColor White
    Write-Host ""
    Write-Host "  Pour installer dans PowerPoint :" -ForegroundColor Cyan
    Write-Host "  Fichier -> Options -> Complements -> Gerer : Complements PowerPoint -> OK" -ForegroundColor White
    Write-Host "  -> Ajouter -> selectionner IconScraper.ppam" -ForegroundColor White
    Write-Host ""
    Write-Host "  Ensuite : Affichage -> Macros -> ShowIconScraper -> Executer" -ForegroundColor White
    Write-Host ""

} catch {
    Write-Host ""
    Write-Host "  ERREUR : $_" -ForegroundColor Red
    Write-Host ""
} finally {
    if ($null -ne $pres)   { try { $pres.Close() }   catch {} }
    if ($null -ne $pptApp) { try { $pptApp.Quit() }  catch {} }

    # Restauration du registre
    if ($null -eq $origValue) {
        Remove-ItemProperty -Path $regPath -Name "AccessVBOM" -ErrorAction SilentlyContinue
    } else {
        Set-ItemProperty -Path $regPath -Name "AccessVBOM" -Value $origValue -Type DWord -Force
    }
    Write-Host "  Securite VBA restauree." -ForegroundColor Green
}
