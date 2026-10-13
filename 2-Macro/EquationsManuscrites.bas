Attribute VB_Name = "EquationsManuscrites"
Option Explicit

' Met toutes les équations du document dans l'écriture manuscrite.
' Changez ces deux valeurs si vous voulez une autre police ou une autre couleur.
Private Const POLICE As String = "Playpen Manuscrit"
Private Const COULEUR As String = "1B3A8C"   ' bleu stylo (code hexadécimal)

Sub EquationsManuscrites()
    Dim i As Long, rg As Range, prefixe As String

    If ActiveDocument.OMaths.Count = 0 Then
        MsgBox "Aucune équation dans ce document.", vbInformation
        Exit Sub
    End If

    prefixe = "<m:r><m:rPr><m:nor/></m:rPr><w:rPr><w:rFonts w:ascii=""" & POLICE & _
              """ w:hAnsi=""" & POLICE & """ w:cs=""" & POLICE & """/><w:color w:val=""" & _
              COULEUR & """/></w:rPr>"

    Application.ScreenUpdating = False
    For i = ActiveDocument.OMaths.Count To 1 Step -1
        Set rg = ActiveDocument.OMaths(i).Range
        rg.InsertXML RemplacerRuns(rg.WordOpenXML, prefixe)
    Next i
    Application.ScreenUpdating = True

    MsgBox ActiveDocument.OMaths.Count & " équation(s) passée(s) en écriture manuscrite.", vbInformation
End Sub

' Remplace le début de chaque <m:r> (et son ancien formatage) par le nouveau formatage.
Private Function RemplacerRuns(ByVal x As String, ByVal prefixe As String) As String
    Dim res As String, p As Long, q As Long
    p = 1
    Do
        q = InStr(p, x, "<m:r>")
        If q = 0 Then
            res = res & Mid$(x, p)
            Exit Do
        End If
        res = res & Mid$(x, p, q - p) & prefixe
        p = q + 5
        p = SauterBalise(x, p, "m:rPr")
        p = SauterBalise(x, p, "w:rPr")
    Loop
    RemplacerRuns = res
End Function

Private Function SauterBalise(ByVal x As String, ByVal p As Long, ByVal nom As String) As Long
    Dim fin As Long
    If Mid$(x, p, Len(nom) + 3) = "<" & nom & "/>" Then
        SauterBalise = p + Len(nom) + 3
    ElseIf Mid$(x, p, Len(nom) + 2) = "<" & nom & ">" Then
        fin = InStr(p, x, "</" & nom & ">")
        SauterBalise = fin + Len(nom) + 3
    Else
        SauterBalise = p
    End If
End Function
