Option Explicit

' Remplacez par le nom EXACT de la police tel qu'il apparaît dans le menu déroulant de Word
Private Const POLICE As String = "Playpen Manuscrit"

Sub EquationsManuscrites()
    Dim i As Long, rg As Range, oMath As oMath

    If ActiveDocument.OMaths.Count = 0 Then
        MsgBox "Aucune équation dans ce document.", vbInformation
        Exit Sub
    End If

    Application.ScreenUpdating = False

    For i = ActiveDocument.OMaths.Count To 1 Step -1
        Set oMath = ActiveDocument.OMaths(i)
        Set rg = oMath.Range
        
        ' 1. Active le mode "Texte normal" mathématique (<m:nor/>) dans le XML pour débloquer la police
        rg.InsertXML PasserEnTexteNormalSansEcraser(rg.WordOpenXML)
        
        ' 2. Force l'application de la police sur le Range (sans toucher aux couleurs ni gras/italique)
        Set rg = ActiveDocument.OMaths(i).Range
        rg.Font.Name = POLICE
    Next i

    Application.ScreenUpdating = True
    MsgBox ActiveDocument.OMaths.Count & " équation(s) traitée(s).", vbInformation
End Sub

Private Function PasserEnTexteNormalSansEcraser(ByVal xml As String) As String
    Dim p As Long, q As Long, finR As Long
    Dim res As String, bloc As String
    res = ""
    p = 1

    Do
        q = InStr(p, xml, "<m:r>")
        If q = 0 Then
            res = res & Mid$(xml, p)
            Exit Do
        End If

        res = res & Mid$(xml, p, q - p)
        finR = InStr(q, xml, "</m:r>")
        If finR = 0 Then
            res = res & Mid$(xml, q)
            Exit Do
        End If

        bloc = Mid$(xml, q + 5, finR - (q + 5))

        ' Ajoute <m:nor/> (Texte Normal) pour débrayer Cambria Math
        If InStr(bloc, "<m:rPr>") > 0 Then
            If InStr(bloc, "<m:nor/>") = 0 Then
                bloc = Replace(bloc, "<m:rPr>", "<m:rPr><m:nor/>", 1, 1)
            End If
        Else
            bloc = "<m:rPr><m:nor/></m:rPr>" & bloc
        End If

        res = res & "<m:r>" & bloc & "</m:r>"
        p = finR + 6
    Loop

    PasserEnTexteNormalSansEcraser = res
End Function
