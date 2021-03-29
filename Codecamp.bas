Attribute VB_Name = "Module1"
Function removeSpecial(sInput As String) As String
       Dim sSpecialChars As String
       Dim i As Long
       sSpecialChars = "\/:*?""<>|." 'This is your list of characters to be removed
       For i = 1 To Len(sSpecialChars)
           sInput = Replace$(sInput, Mid$(sSpecialChars, i, 1), " ") 'this will remove spaces
       Next
       removeSpecial = sInput
End Function
Sub Codecamp()
'
' Codecamp Macro
'
    Columns("AO:AT").ColumnWidth = ActiveSheet.StandardWidth
    Dim longueur_colonne As Integer
    longueur_colonne = Application.WorksheetFunction.CountA(Range("A:A"))
    Sheets("00 Reprise Sales report").Activate
    Range(("A2:K2")).Sort Header:=xlYes, _
    Key1:=Range("K2"), Order1:=xlAscending, Key2:=Range("A2") _
    , Order2:=xlAscending
    Dim saveLocation As String
    Dim valeur As Integer
    Dim nbr_pdf As String
    Dim nbr_page As String
    Dim nbr_ligne_pdf As String
    Dim colonne As Integer
    Dim nbr_bon As String
    valeur = 2
    Do While valeur <= longueur_colonne
        nbr_pdf = 0
        nbr_page = 0
        nbr_ligne_pdf = 69
        nbr_bon = 0
        Range("AO" & 2 + nbr_pdf + (nbr_page * nbr_ligne_pdf)) = Application.Proper(Range("K" & valeur))
        Range("AO" & 2 + nbr_pdf + (nbr_page * nbr_ligne_pdf)).Font.Bold = True
        Range("AQ" & 2 + nbr_pdf + (nbr_page * nbr_ligne_pdf)) = ",livraison du "
        Range("AQ" & 2 + nbr_pdf + (nbr_page * nbr_ligne_pdf)).Font.Bold = True
        Columns("AQ").EntireColumn.AutoFit
        Range("AO" & 4 + nbr_pdf + (nbr_page * nbr_ligne_pdf)) = Range("C" & valeur)
        Range("AO" & 4 + nbr_pdf + (nbr_page * nbr_ligne_pdf)).Font.Bold = True
        Range("AO" & 6 + nbr_pdf + (nbr_page * nbr_ligne_pdf)) = Range("D" & valeur)
        If Range("E" & valeur) = "" Then
            Range("AO" & 7 + nbr_pdf + (nbr_page * nbr_ligne_pdf)) = Range("F" & valeur)
            Range("AO" & 8 + nbr_pdf + (nbr_page * nbr_ligne_pdf)) = Range("G" & valeur)
        Else
            Range("AO" & 7 + nbr_pdf + (nbr_page * nbr_ligne_pdf)) = Range("E" & valeur)
            Range("AO" & 8 + nbr_pdf + (nbr_page * nbr_ligne_pdf)) = Range("F" & valeur)
            Range("AO" & 9 + nbr_pdf + (nbr_page * nbr_ligne_pdf)) = Range("G" & valeur)
        End If
        Range("AO" & 11 + nbr_pdf + (nbr_page * nbr_ligne_pdf)) = "Tel : "
        Range("AQ" & 11 + nbr_pdf + (nbr_page * nbr_ligne_pdf)) = Range("H" & valeur)
        Range("AO" & 13 + nbr_pdf + (nbr_page * nbr_ligne_pdf)) = "Mode de retrait : "
        Range("AO" & 13 + nbr_pdf + (nbr_page * nbr_ligne_pdf)).Font.Bold = True
        Range("AQ" & 13 + nbr_pdf + (nbr_page * nbr_ligne_pdf)) = Range("S" & valeur)
        If Range("S" & valeur) = "pickup" Then
            Range("AO" & 14 + nbr_pdf + (nbr_page * nbr_ligne_pdf)) = "Lieu de retrait : "
            Range("AO" & 14 + nbr_pdf + (nbr_page * nbr_ligne_pdf)).Font.Bold = True
            Range("AQ" & 14 + nbr_pdf + (nbr_page * nbr_ligne_pdf)) = Range("T" & valeur)
            Range("AR" & 2 + nbr_pdf + (nbr_page * nbr_ligne_pdf)) = Range("U" & valeur)
            Range("AQ" & 15 + nbr_pdf + (nbr_page * nbr_ligne_pdf)) = Range("V" & valeur)
        Else
            Range("AO" & 14 + nbr_pdf + (nbr_page * nbr_ligne_pdf)) = "Lieu de livraison : "
            Range("AO" & 14 + nbr_pdf + (nbr_page * nbr_ligne_pdf)).Font.Bold = True
            Range("AQ" & 14 + nbr_pdf + (nbr_page * nbr_ligne_pdf)) = Range("Y" & valeur)
            Range("AR" & 2 + nbr_pdf + (nbr_page * nbr_ligne_pdf)) = Range("W" & valeur)
            Range("AQ" & 15 + nbr_pdf + (nbr_page * nbr_ligne_pdf)) = Range("X" & valeur)
        End If
        Range("AO" & 15 + nbr_pdf + (nbr_page * nbr_ligne_pdf)) = "Heure de retrait :"
        Range("AO" & 15 + nbr_pdf + (nbr_page * nbr_ligne_pdf)).Font.Bold = True
        Range("AO" & 16 + nbr_pdf + (nbr_page * nbr_ligne_pdf)) = "Commande :"
        Range("AQ" & 16 + nbr_pdf + (nbr_page * nbr_ligne_pdf)) = Range("AD" & valeur)
        Range("AQ" & 16 + nbr_pdf + (nbr_page * nbr_ligne_pdf)).Font.Color = RGB(255, 0, 0)
        Range("AQ" & 16 + nbr_pdf + (nbr_page * nbr_ligne_pdf)).Font.Bold = True
        Range("AO" & 18 + nbr_pdf + (nbr_page * nbr_ligne_pdf)) = Range("AE" & valeur)
        Range("AO" & 19 + nbr_pdf + (nbr_page * nbr_ligne_pdf)) = "Produit"
        Range("AR" & 19 + nbr_pdf + (nbr_page * nbr_ligne_pdf)) = "Variante"
        Range("AT" & 19 + nbr_pdf + (nbr_page * nbr_ligne_pdf)) = "Quantite"
        Range("AO" & 19 + nbr_pdf + (nbr_page * nbr_ligne_pdf) & ":AT" & 19 + nbr_pdf + (nbr_page * nbr_ligne_pdf)).Font.Color = RGB(0, 0, 255)
        colonne = 20
        Range("AO" & colonne + nbr_pdf + (nbr_page * nbr_ligne_pdf)) = Range("AH" & valeur)
        Range("AR" & colonne + nbr_pdf + (nbr_page * nbr_ligne_pdf)) = Range("AI" & valeur)
        Range("AT" & colonne + nbr_pdf + (nbr_page * nbr_ligne_pdf)) = Range("AJ" & valeur)
        Range("AO" & colonne + nbr_pdf + (nbr_page * nbr_ligne_pdf) & ":AT" & colonne + nbr_pdf + (nbr_page * nbr_ligne_pdf)).Borders(xlEdgeTop).Weight = xlThin
        Range("AO" & colonne + nbr_pdf + (nbr_page * nbr_ligne_pdf) & ":AT" & colonne + nbr_pdf + (nbr_page * nbr_ligne_pdf)).Borders(xlEdgeBottom).Weight = xlThin
        Do While Range("A" & valeur) = Range("A" & valeur + 1) And valeur <= longueur_colonne
            valeur = valeur + 1
            colonne = colonne + 1
            Range("AO" & colonne + nbr_pdf + (nbr_page * nbr_ligne_pdf)) = Range("AH" & valeur)
            Range("AR" & colonne + nbr_pdf + (nbr_page * nbr_ligne_pdf)) = Range("AI" & valeur)
            Range("AT" & colonne + nbr_pdf + (nbr_page * nbr_ligne_pdf)) = Range("AJ" & valeur)
            Range("AO" & colonne + nbr_pdf + (nbr_page * nbr_ligne_pdf) & ":AT" & colonne + nbr_pdf + (nbr_page * nbr_ligne_pdf)).Borders(xlEdgeTop).Weight = xlThin
            Range("AO" & colonne + nbr_pdf + (nbr_page * nbr_ligne_pdf) & ":AT" & colonne + nbr_pdf + (nbr_page * nbr_ligne_pdf)).Borders(xlEdgeBottom).Weight = xlThin
        Loop
        Range("AO" & 19 + nbr_pdf + (nbr_page * nbr_ligne_pdf) & ":AT" & 19 + nbr_pdf + (nbr_page * nbr_ligne_pdf)).Borders(xlEdgeTop).Weight = xlMedium
        Range("AO" & 19 + nbr_pdf + (nbr_page * nbr_ligne_pdf) & ":AT" & 19 + nbr_pdf + (nbr_page * nbr_ligne_pdf)).Borders(xlEdgeBottom).Weight = xlMedium
        'MkDir (ActiveWorkbook.Path & "\Etiquettes\")
        valeur = valeur + 1
        nbr_pdf = nbr_pdf + colonne + 4
        nbr_bon = nbr_bon + 1
        Do While Application.Proper(Range("K" & valeur)) = Application.Proper(Range("K" & valeur - 1)) And valeur <= longueur_colonne
            If nbr_bon = 2 Then
                nbr_page = nbr_page + 1
                nbr_pdf = 0
                nbr_bon = 0
            End If
            Range("AO" & 2 + nbr_pdf + (nbr_page * nbr_ligne_pdf)) = Application.Proper(Range("K" & valeur))
            Range("AO" & 2 + nbr_pdf + (nbr_page * nbr_ligne_pdf)).Font.Bold = True
            Range("AQ" & 2 + nbr_pdf + (nbr_page * nbr_ligne_pdf)) = ", livraison du "
            Range("AQ" & 2 + nbr_pdf + (nbr_page * nbr_ligne_pdf)).Font.Bold = True
            Range("AO" & 4 + nbr_pdf + (nbr_page * nbr_ligne_pdf)) = Range("C" & valeur)
            Range("AO" & 4 + nbr_pdf + (nbr_page * nbr_ligne_pdf)).Font.Bold = True
            Range("AO" & 6 + nbr_pdf + (nbr_page * nbr_ligne_pdf)) = Range("D" & valeur)
            If Range("E" & valeur) = "" Then
                Range("AO" & 7 + nbr_pdf + (nbr_page * nbr_ligne_pdf)) = Range("F" & valeur)
                Range("AO" & 8 + nbr_pdf + (nbr_page * nbr_ligne_pdf)) = Range("G" & valeur)
            Else
                Range("AO" & 7 + nbr_pdf + (nbr_page * nbr_ligne_pdf)) = Range("E" & valeur)
                Range("AO" & 8 + nbr_pdf + (nbr_page * nbr_ligne_pdf)) = Range("F" & valeur)
                Range("AO" & 9 + nbr_pdf + (nbr_page * nbr_ligne_pdf)) = Range("G" & valeur)
            End If
            Range("AO" & 11 + nbr_pdf + (nbr_page * nbr_ligne_pdf)) = "Tel : "
            Range("AQ" & 11 + nbr_pdf + (nbr_page * nbr_ligne_pdf)) = Range("H" & valeur)
            Range("AO" & 13 + nbr_pdf + (nbr_page * nbr_ligne_pdf)) = "Mode de retrait : "
            Range("AO" & 13 + nbr_pdf + (nbr_page * nbr_ligne_pdf)).Font.Bold = True
            Range("AQ" & 13 + nbr_pdf + (nbr_page * nbr_ligne_pdf)) = Range("S" & valeur)
            If Range("S" & valeur) = "pickup" Then
                Range("AO" & 14 + nbr_pdf + (nbr_page * nbr_ligne_pdf)) = "Lieu de retrait : "
                Range("AO" & 14 + nbr_pdf + (nbr_page * nbr_ligne_pdf)).Font.Bold = True
                Range("AQ" & 14 + nbr_pdf + (nbr_page * nbr_ligne_pdf)) = Range("T" & valeur)
                Range("AR" & 2 + nbr_pdf + (nbr_page * nbr_ligne_pdf)) = Range("U" & valeur)
                Range("AQ" & 15 + nbr_pdf + (nbr_page * nbr_ligne_pdf)) = Range("V" & valeur)
            Else
                Range("AO" & 14 + nbr_pdf + (nbr_page * nbr_ligne_pdf)) = "Lieu de livraison : "
                Range("AO" & 14 + nbr_pdf + (nbr_page * nbr_ligne_pdf)).Font.Bold = True
                Range("AQ" & 14 + nbr_pdf + (nbr_page * nbr_ligne_pdf)) = Range("Y" & valeur)
                Range("AR" & 2 + nbr_pdf + (nbr_page * nbr_ligne_pdf)) = Range("W" & valeur)
                Range("AQ" & 15 + nbr_pdf + (nbr_page * nbr_ligne_pdf)) = Range("X" & valeur)
            End If
            Range("AO" & 15 + nbr_pdf + (nbr_page * nbr_ligne_pdf)) = "Heure de retrait :"
            Range("AO" & 15 + nbr_pdf + (nbr_page * nbr_ligne_pdf)).Font.Bold = True
            Range("AO" & 16 + nbr_pdf + (nbr_page * nbr_ligne_pdf)) = "Commande :"
            Range("AQ" & 16 + nbr_pdf + (nbr_page * nbr_ligne_pdf)) = Range("AD" & valeur)
            Range("AQ" & 16 + nbr_pdf + (nbr_page * nbr_ligne_pdf)).Font.Color = RGB(255, 0, 0)
            Range("AQ" & 16 + nbr_pdf + (nbr_page * nbr_ligne_pdf)).Font.Bold = True
            Range("AO" & 18 + nbr_pdf + (nbr_page * nbr_ligne_pdf)) = Range("AE" & valeur)
            Range("AO" & 19 + nbr_pdf + (nbr_page * nbr_ligne_pdf)) = "Produit"
            Range("AR" & 19 + nbr_pdf + (nbr_page * nbr_ligne_pdf)) = "Variante"
            Range("AT" & 19 + nbr_pdf + (nbr_page * nbr_ligne_pdf)) = "Quantite"
            Range("AO" & 19 + nbr_pdf + (nbr_page * nbr_ligne_pdf) & ":AT" & 19 + nbr_pdf + (nbr_page * nbr_ligne_pdf)).Font.Color = RGB(0, 0, 255)
            colonne = 20
            Range("AO" & colonne + nbr_pdf + (nbr_page * nbr_ligne_pdf)) = Range("AH" & valeur)
            Range("AR" & colonne + nbr_pdf + (nbr_page * nbr_ligne_pdf)) = Range("AI" & valeur)
            Range("AT" & colonne + nbr_pdf + (nbr_page * nbr_ligne_pdf)) = Range("AJ" & valeur)
            Range("AO" & colonne + nbr_pdf + (nbr_page * nbr_ligne_pdf) & ":AT" & colonne + nbr_pdf + (nbr_page * nbr_ligne_pdf)).Borders(xlEdgeTop).Weight = xlThin
            Range("AO" & colonne + nbr_pdf + (nbr_page * nbr_ligne_pdf) & ":AT" & colonne + nbr_pdf + (nbr_page * nbr_ligne_pdf)).Borders(xlEdgeBottom).Weight = xlThin
            Do While Range("A" & valeur) = Range("A" & valeur + 1) And valeur <= longueur_colonne
                valeur = valeur + 1
                colonne = colonne + 1
                Range("AO" & colonne + nbr_pdf + (nbr_page * nbr_ligne_pdf)) = Range("AH" & valeur)
                Range("AR" & colonne + nbr_pdf + (nbr_page * nbr_ligne_pdf)) = Range("AI" & valeur)
                Range("AT" & colonne + nbr_pdf + (nbr_page * nbr_ligne_pdf)) = Range("AJ" & valeur)
                Range("AO" & colonne + nbr_pdf + (nbr_page * nbr_ligne_pdf) & ":AT" & colonne + nbr_pdf + (nbr_page * nbr_ligne_pdf)).Borders(xlEdgeTop).Weight = xlThin
                Range("AO" & colonne + nbr_pdf + (nbr_page * nbr_ligne_pdf) & ":AT" & colonne + nbr_pdf + (nbr_page * nbr_ligne_pdf)).Borders(xlEdgeBottom).Weight = xlThin
            Loop
            Range("AO" & 19 + nbr_pdf + (nbr_page * nbr_ligne_pdf) & ":AT" & 19 + nbr_pdf + (nbr_page * nbr_ligne_pdf)).Borders(xlEdgeTop).Weight = xlMedium
            Range("AO" & 19 + nbr_pdf + (nbr_page * nbr_ligne_pdf) & ":AT" & 19 + nbr_pdf + (nbr_page * nbr_ligne_pdf)).Borders(xlEdgeBottom).Weight = xlMedium
            'MkDir (ActiveWorkbook.Path & "\Etiquettes\")
            valeur = valeur + 1
            nbr_pdf = nbr_pdf + colonne + 4
            nbr_bon = nbr_bon + 1
        Loop
        saveLocation = ActiveWorkbook.Path & "\" & "Etiquettes " & removeSpecial(Range("AO2")) & " " & removeSpecial(Range("AR2")) & ".pdf"
        'Save Active Sheet(s) as PDF
        Range("AO2:AT" & colonne + nbr_pdf + (nbr_page * nbr_ligne_pdf)).ExportAsFixedFormat Type:=xlTypePDF, _
            Filename:=saveLocation
        Range("AO2:AT" & colonne + nbr_pdf + (nbr_page * nbr_ligne_pdf)).Clear
    Loop
    Range(("A2:K2")).Sort Header:=xlYes, _
    Key1:=Range("K2"), Order1:=xlAscending, Key2:=Range("L2") _
    , Order2:=xlAscending, Key3:=Range("M2") _
    , Order3:=xlAscending
    Dim calcul As Integer
    Dim vieux_valeur As Integer
    vieux_valeur = 0
    calcul = 0
    valeur = 2
    Range("AO2") = Range("W1")
    Range("AP2") = Range("U1")
    Range("AQ2") = "Vendor"
    Range("AR2") = Range("L1")
    Range("AS2") = Range("M1")
    Range("AT2") = Range("AB1")
    Range("AO2:AT2").Font.Bold = True
    colonne = 3
    vieux_valeur = valeur
    Do While valeur <= longueur_colonne
        calcul = 0
        Do
            calcul = 0
            vieux_valeur = valeur
            Do While Range("L" & valeur) = Range("L" & valeur + 1) And Range("AI" & valeur) = Range("AI" & valeur + 1) And valeur <= longueur_colonne
                valeur = valeur + 1
            Loop
            Do While vieux_valeur <= valeur
                calcul = calcul + Range("AB" & vieux_valeur)
                vieux_valeur = vieux_valeur + 1
            Loop
            If calcul <> 0 Then
                Range("AQ" & colonne) = Application.Proper(Range("K" & valeur))
                Range("AO" & colonne) = Range("W" & valeur)
                Range("AP" & colonne) = Range("U" & valeur)
                Range("AR" & colonne) = Range("L" & valeur)
                Range("AS" & colonne) = Range("M" & valeur)
                Range("AT" & colonne) = calcul
                colonne = colonne + 1
            End If
            valeur = valeur + 1
            calcul = 0
        Loop While Application.Proper(Range("K" & valeur)) = Application.Proper(Range("K" & valeur - 1)) And valeur <= longueur_colonne
        If vieux_valeur = valeur And Range("K" & valeur + 1) <> Range("AQ" & colonne - 1) And colonne <> 3 Then
            Columns("AO:AT").EntireColumn.AutoFit
            saveLocation = ActiveWorkbook.Path & "\Bon de commande " & removeSpecial(Range("AQ3")) & ".pdf"
            'Save Active Sheet(s) as PDF
            Range("AO2:AT" & colonne).ExportAsFixedFormat Type:=xlTypePDF, _
                Filename:=saveLocation
            Range("AO3:AT" & colonne).Clear
            colonne = 3
        End If
    Loop
    Range("AO2:AT2").Clear
    Columns("AO:AT").ColumnWidth = ActiveSheet.StandardWidth
    MsgBox "Vos etiquettes et vos bons de commandes sont prets !"
'
End Sub


