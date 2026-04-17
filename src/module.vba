Option Explicit

Private Const ICONIFY_API As String = "https://api.iconify.design"
Private Const MAX_RESULTS As Integer = 20

' ═══════════════════════════════════════════════
'  POINT D'ENTREE PRINCIPAL
' ═══════════════════════════════════════════════
Public Sub ShowIconScraper()
    ' 1. Mot-cle
    Dim query As String
    query = Trim(InputBox("Que recherches-tu ?" & vbCrLf & "(ex: home, arrow, github, microsoft...)", "Icon Scraper - Recherche"))
    If query = "" Then Exit Sub

    ' 2. Collection (optionnel)
    Dim prefixChoice As String
    prefixChoice = InputBox( _
        "Choisis une collection (laisser vide = toutes) :" & vbCrLf & vbCrLf & _
        "  mdi       Material Design Icons" & vbCrLf & _
        "  fa        Font Awesome" & vbCrLf & _
        "  logos     Logos de marques" & vbCrLf & _
        "  fluent    Microsoft Fluent" & vbCrLf & _
        "  ph        Phosphor" & vbCrLf & _
        "  tabler    Tabler Icons" & vbCrLf & _
        "  heroicons Heroicons", _
        "Icon Scraper - Collection", "")
    If prefixChoice = "False" Then Exit Sub
    Dim prefix As String
    prefix = Trim(prefixChoice)

    ' 3. Recherche
    Dim icons() As String
    icons = SearchIcons(query, prefix)

    If UBound(icons) = 0 And icons(0) = "" Then
        MsgBox "Aucun resultat pour """ & query & """." & vbCrLf & "Essaie un autre mot-cle.", vbInformation, "Icon Scraper"
        Exit Sub
    End If

    ' 4. Affichage des resultats + selection
    Dim resultList As String
    Dim i As Integer, count As Integer
    Dim validIcons(19) As String
    count = 0
    For i = 0 To UBound(icons)
        Dim ic As String
        ic = Trim(icons(i))
        If ic <> "" And count < MAX_RESULTS Then
            validIcons(count) = ic
            resultList = resultList & " " & (count + 1) & ".  " & ic & vbCrLf
            count = count + 1
        End If
    Next i

    Dim numStr As String
    numStr = InputBox( _
        count & " icone(s) trouvee(s) pour """ & query & """ :" & vbCrLf & vbCrLf & _
        resultList & vbCrLf & _
        "Entre le numero de l'icone a inserer :", _
        "Icon Scraper - Resultats", "1")
    If numStr = "" Or numStr = "False" Then Exit Sub

    Dim num As Integer
    On Error Resume Next
    num = CInt(numStr)
    On Error GoTo 0
    If num < 1 Or num > count Then
        MsgBox "Numero invalide.", vbExclamation, "Icon Scraper"
        Exit Sub
    End If

    ' 5. Taille
    Dim sizeStr As String
    sizeStr = InputBox("Taille de l'icone en pixels :", "Icon Scraper - Taille", "128")
    If sizeStr = "" Or sizeStr = "False" Then Exit Sub
    Dim sizePx As Integer
    On Error Resume Next
    sizePx = CInt(sizeStr)
    On Error GoTo 0
    If sizePx < 16 Then sizePx = 128

    ' 6. Insertion
    InsertIcon validIcons(num - 1), sizePx
End Sub

' ═══════════════════════════════════════════════
'  RECHERCHE VIA API ICONIFY
' ═══════════════════════════════════════════════
Public Function SearchIcons(query As String, Optional prefix As String = "") As String()
    Dim http As Object
    Dim url As String
    Dim icons() As String

    url = ICONIFY_API & "/search?query=" & EncodeURL(query) & "&limit=" & MAX_RESULTS
    If prefix <> "" Then url = url & "&prefix=" & prefix

    On Error GoTo ErrHandler
    Set http = CreateObject("MSXML2.XMLHTTP60")
    http.Open "GET", url, False
    http.setRequestHeader "Accept", "application/json"
    http.Send

    If http.Status = 200 Then
        icons = ParseIconsJSON(http.responseText)
    Else
        MsgBox "Erreur API : HTTP " & http.Status, vbExclamation, "Icon Scraper"
        ReDim icons(0): icons(0) = ""
    End If
    SearchIcons = icons
    Exit Function

ErrHandler:
    MsgBox "Erreur reseau : " & Err.Description, vbExclamation, "Icon Scraper"
    ReDim icons(0): icons(0) = ""
    SearchIcons = icons
End Function

' ═══════════════════════════════════════════════
'  PARSE JSON ICONIFY
' ═══════════════════════════════════════════════
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

' ═══════════════════════════════════════════════
'  INSERTION DANS LA SLIDE
' ═══════════════════════════════════════════════
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

    ' URL SVG Iconify
    Dim url As String
    url = ICONIFY_API & "/" & prefix & "/" & iconName & ".svg" & _
          "?width=" & sizePx & "&height=" & sizePx & "&color=%23333333"

    ' Telechargement
    Dim http As Object
    On Error GoTo ErrDownload
    Set http = CreateObject("MSXML2.XMLHTTP60")
    http.Open "GET", url, False
    http.Send

    If http.Status <> 200 Then
        MsgBox "Erreur HTTP " & http.Status & " pour : " & iconFullName, vbExclamation, "Icon Scraper"
        Exit Sub
    End If

    ' Sauvegarde SVG en fichier temp
    Dim tempFile As String
    tempFile = Environ("TEMP") & "\iconscrap_" & prefix & "_" & iconName & ".svg"

    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1
    stream.Open
    stream.Write http.responseBody
    stream.SaveToFile tempFile, 2
    stream.Close

    ' Verification presentation ouverte
    If Application.Presentations.Count = 0 Then
        MsgBox "Aucune presentation ouverte.", vbExclamation, "Icon Scraper"
        Exit Sub
    End If

    ' Insertion dans la slide active
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

    MsgBox "Insere : " & iconFullName, vbInformation, "Icon Scraper"
    Exit Sub

ErrDownload:
    MsgBox "Erreur : " & Err.Description, vbExclamation, "Icon Scraper"
End Sub

' ═══════════════════════════════════════════════
'  ENCODAGE URL
' ═══════════════════════════════════════════════
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
