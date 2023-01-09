Attribute VB_Name = "MajPickupXFModule"
Option Explicit

'The MIT License (MIT)
'
'Copyright (c) 2022 FORREST
' Mateusz Milewski mateusz.milewski@mpsa.com aka FORREST
'
'Permission is hereby granted, free of charge, to any person obtaining a copy
'of this software and associated documentation files (the "Software"), to deal
'in the Software without restriction, including without limitation the rights
'to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
'copies of the Software, and to permit persons to whom the Software is
'furnished to do so, subject to the following conditions:
'
'The above copyright notice and this permission notice shall be included in all
'copies or substantial portions of the Software.
'
'THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
'IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
'FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
'AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
'LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
'OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
'SOFTWARE.
'
'
' NEW PUS For xF features



' Option Explicit


Public majGlobale As Boolean



Sub maj_pickup__xF()

    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    'Application.StatusBar = "2/5 : format et calculs Sechel"
    Call mise_forme_sechel
    majGlobale = True
    'Application.StatusBar = "3/5 : mise à jour onglet BASE"
    ' -------------------------------
    Call copie_reference_base
    ' -------------------------------
    'Application.StatusBar = "4/5 : récupération infos BASE_temp"
    Call copie_informations_cpl
    'Application.StatusBar = "5/5 : paramétrage onglet PUS creation"
    majGlobale = True
    ' MOD
    ' ==================================================================
    Call useDeliveryProgramColumnAndFillItWithKindOfPUSnumber
    ' ==================================================================
    Call majNomFNR__xF
    Application.EnableEvents = True
    majGlobale = False
    m = MsgBox("Fin de la mise à jour.", vbOKOnly + vbExclamation, "Attention")
    Application.Calculation = xlCalculationAutomatic
End Sub




Private Sub useDeliveryProgramColumnAndFillItWithKindOfPUSnumber()
     
    Debug.Print "useDeliveryProgramColumnAndFillItWithKindOfPUSnumber"
    ' --------------------------------------------------------------------
    
    '    Public Property Get col_del_prog_b() As Variant ' 8 ' H
    '    Public Property Get col_po_numb_b() As Variant ' 9 ' I
    ' Cells(2, 16).Value
    
    
    Dim bs As Worksheet, i As Variant, strCofor As String, strRef As String, longCofor As Long, longRef As Double
    Set bs = ThisWorkbook.Sheets("BASE")
    
    
    For i = 3 To bs.Cells(500000, col_ref_s).End(xlUp).Row
    
    
        strCofor = bs.Cells(i, col_cofor_b).Value
        'longCofor = CLng(strCofor)
        strRef = bs.Cells(i, col_ref_b).Value
        'longRef = CDbl(strRef)
        'longRef = longRef / 1000
        
        
        ' tmp
        'strCofor = Hex(longCofor) ' OK
        'longCofor = CDec("&H" & strCofor) ' OK - coming back from HEX to Decimal again - just in case...
    
        bs.Cells(i, col_del_prog_b) = "" & bs.Cells(i, col_base_mag).Value & "_" & _
            Trim(strCofor) & "_" & _
            Trim(strRef)
        bs.Cells(i, col_po_numb_b) = "" & bs.Cells(i, col_base_mag).Value & "_" & _
            Trim(strCofor) & "_" & _
            Trim(strRef)
    Next i
    
    
    
    ' --------------------------------------------------------------------
End Sub





Sub mise_forme_sechel()
    
    Sheets("sechel").Visible = True
    Sheets("progliv").Visible = True
    Sheets("RLF").Visible = True
    Sheets("ME9E").Visible = True
    Sheets("BASE_temp").Visible = True
    
    Application.ScreenUpdating = True
    
    
    Sheets("RLF").Select
    If Not ActiveSheet.AutoFilter Is Nothing Then
    Selection.AutoFilter
    Else
    End If
    
    
    'Filtre des EL1
    Sheets("sechel").Select
    If Not ActiveSheet.AutoFilter Is Nothing Then
        Selection.AutoFilter
    End If
    
    
    Dim dom As String
    Dim dateEL1 As Variant
    
    i = 6 '--> ligne des domaines
    For j = 2 To Sheets("Macro & projects infos").Cells(i - 1, 1).End(xlToRight).Column
        'On vérifie si le domaine et la date sont bien remplis avant de les utiliser
        If Sheets("Macro & projects infos").Cells(i, j).Value <> "" And Sheets("Macro & projects infos").Cells(i + 2, j).Value <> "" Then
            dom = Sheets("Macro & projects infos").Cells(i, j).Value
            dateEL1 = Format(Sheets("Macro & projects infos").Cells(i + 2, j).Value, "mm / dd / yyyy")
            
            'Filtre des EL1
            Sheets("sechel").Cells.AutoFilter Field:=10, Criteria1:=dom
            Sheets("sechel").Cells.AutoFilter Field:=12, Criteria1:="<" & dateEL1
            nbre_ligne_filtree = [_filterdatabase].Resize(, 1).SpecialCells(xlCellTypeVisible).count - 1
            If nbre_ligne_filtree >= 1 Then
                Range("_FilterDataBase").Offset(1, 0).Resize(Range("_FilterDataBase"). _
                    Rows.count - 1).SpecialCells(xlCellTypeVisible).Delete Shift:=xlUp
            Else
                ' Debug.Print Range("_FilterDataBase").Address
            End If
            ActiveSheet.AutoFilterMode = False
        End If
    Next j
    
    
    ' Si plusieurs échéances à la même date pour une même référence => arrêt macro et pop up
    Columns(15).Select
    Selection.ClearContents
    Cells(1, 15) = "conca ref et écheance"
    
    Columns(16).Select
    Selection.ClearContents
    Cells(1, 16) = "test doublon écheance"
    
    For i = 2 To Cells(500000, col_ref_s).End(xlUp).Row
        Cells(i, 15) = Cells(i, 2) & Cells(i, 5) & Cells(i, 7) & Cells(i, 12)
    Next i
    
    For i = 2 To Cells(500000, col_ref_s).End(xlUp).Row
        Cells(i, 16) = Application.WorksheetFunction.CountIf(Range("O2:O" & Cells(500000, col_ref_s).End(xlUp).Row), Cells(i, 15))
    Next i
    
    For i = 2 To Cells(500000, col_ref_s).End(xlUp).Row
         If Cells(i, 16) = 1 Then
         Else
            If (MsgBox("La référence " & Cells(i, col_ref_s) & " se répète pour une même date d'échéance, sur le même mag." & Chr(10) & "Voulez-vous poursuivre le process ?", vbYesNo) = vbYes) Then
                'on poursuit
                GoTo keepOnGoing
            Else
                'Restaure la donnée par défaut de la barre d'état
                Application.StatusBar = False
                Sheets("sechel").Cells(i, col_ref_s).Select
                End
            End If
         End If
    Next i
    
    
keepOnGoing:
    
    '''''''''''''''''''''''''''''
    Columns(15).Select
    Selection.ClearContents
    Cells(1, 15) = "quantité à livrer"
    
    Columns(16).Select
    Selection.ClearContents
    Cells(1, 16) = "numéro contrat"
    
    Columns(17).Select
    Selection.ClearContents
    Cells(1, 17) = "semaine de l'échéance"
     
    Columns(18).Select
    Selection.ClearContents
    Cells(1, 18) = "mail appro"
    
    Columns(19).Select
    Selection.ClearContents
    Cells(1, 19) = "échéancier maj- pickupsheet à renvoyer"
    
    Columns(20).Select
    Selection.ClearContents
    Cells(1, 20) = "Type"
    
    Columns(21).Select
    Selection.ClearContents
    Cells(1, 21) = "Mois"
    
    ''' identification des fnrs avec des noms identiques sous des COFORs différent (dans ce cas ajout du cofor à coté du nom)
    Sheets("sechel").Select
    If Not ActiveSheet.AutoFilter Is Nothing Then
    Selection.AutoFilter
    Else
    End If
    
    Columns("D:E").Select
    Selection.Copy
    Range("AX1").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    ActiveSheet.Range(Cells(1, 50), Cells(Cells(500000, 51).End(xlUp).Row, 51)).RemoveDuplicates Columns:=1, Header:=xlYes
    
    For i = 2 To Cells(500000, 50).End(xlUp).Row
        cofor_fnr_s = Cells(i, 50)
        nom_fnr_s = Cells(i, 51)
        
        For j = i + 1 To Cells(500000, 50).End(xlUp).Row
            If Cells(j, 51) = nom_fnr_s And Cells(j, 50) <> cofor_fnr_s Then 'dénomination fnr à discriminer
                For k = 2 To Cells(500000, 2).End(xlUp).Row
                    If Cells(k, 5) = nom_fnr_s Then
                        Cells(k, 5) = nom_fnr_s & " - " & Cells(k, 4)
                    Else: End If
                Next k
            Else: End If
        Next j
    Next i
    
    Columns("AX:AY").Select
    Selection.ClearContents
    Sheets("sechel").Select
    If Not ActiveSheet.AutoFilter Is Nothing Then
    Selection.AutoFilter
    Else
    End If
    Range("A1").Select
    
    
    Sheets("sechel").Select
    If Not ActiveSheet.AutoFilter Is Nothing Then
    Selection.AutoFilter
    Else
    End If
    
    
    Dim endingLineSechel As Long
    endingLineSechel = Cells(100000, 1).End(xlUp).Row
    
    Columns("U:Z").Select
    Selection.ClearContents
    Columns("AX:BC").Select
    Selection.ClearContents
    Sheets("sechel").Select
    If Not ActiveSheet.AutoFilter Is Nothing Then
    Selection.AutoFilter
    Else
    End If
    Range("A1").Select
    
    '0.Tableau de correspondance TMC
    Dim p As Long
    Dim q As Long
    p = 0
    q = 0
    i = 6
    For j = 2 To Sheets("Macro & projects infos").Cells(5, 1).End(xlToRight).Column
        If Sheets("Macro & projects infos").Cells(i, j).Value <> "" Then
            'Partie TMC
            For k = 13 To Sheets("Macro & projects infos").Cells(1000, 1).End(xlUp).Row
                If Sheets("Macro & projects infos").Cells(k, j).Value <> "" Then
                    q = q + 1
                    Sheets("Sechel").Cells(q, 53).Value = Sheets("Macro & projects infos").Cells(i, j).Value & Sheets("Macro & projects infos").Cells(k, j).Value
                    Sheets("Sechel").Cells(q, 54).Value = Sheets("Macro & projects infos").Cells(k, 1).Value
                End If
            Next k
        End If
    Next j


    ' column 10 -> domain -> for now from PCV perspective taking cost center number
    'forcer en chiffre la colonne mag
    'Columns(10).TextToColumns , DataType:=xlDelimited,
    '        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
    '        Semicolon:=True, Comma:=True, Space:=False, Other:=False, FieldInfo:= _
    '        Array(1, 1), TrailingMinusNumbers:=True
    
    
    ' column 8 -> doc achat column - no need to make any instruction!
    'Sheets("Sechel").Columns(8).TextToColumns , DataType:=xlDelimited, _
    '        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
    '        Semicolon:=True, Comma:=True, Space:=False, Other:=False, FieldInfo:= _
    '        Array(1, 2), TrailingMinusNumbers:=True
    
    '1.Ecrire formules sur la première ligne
    Sheets("sechel").Select
    
    ' do we really want to double check code of suplpier planner?
    'Cells(2, 1).NumberFormat = "General"
    'Cells(2, 1).FormulaR1C1 = "=IF(IFERROR(VLOOKUP(RC11,NOA!C1:C4,4,0),"""")="""",""à compléter - onglet NOA"",VLOOKUP(RC11,NOA!C1:C4,4,0))"
    
    Cells(2, 15).FormulaR1C1 = "=RC13-RC14"
    
    ' PROG LIV !!!
    ' --------------------------------------------------------------------------------------------
    ' tmp not used - excpetion I made a seperate fomrula
    ' Cells(2, 16).FormulaR1C1 = "=IFERROR(VLOOKUP(RC8,progliv!C1:C2,2,0),"""")"
    ' --------------------------------------------------------------------------------------------
    
    
    Cells(2, 17).FormulaR1C1 = "=IF(WEEKNUM(RC12,21)<10,RIGHT(YEAR(RC12),2) & ""-CW0"" & WEEKNUM(RC12,21),RIGHT(YEAR(RC12),2) & ""-CW"" & WEEKNUM(RC12,21))"
    Cells(2, 18).FormulaR1C1 = "=IFERROR(VLOOKUP(RC11,NOA!C1:C4,3,0),"""")"
    
    ' ME9E !!!
    ' ---------------------------------------------------------------------------------------------
    ' tmp not used
    ' Cells(2, 19).FormulaR1C1 = "=IF(IFERROR(VLOOKUP(RC8,ME9E!C2,1,0),"""")="""","""",""oui"")"
    ' ---------------------------------------------------------------------------------------------
    
    
    Cells(2, 20).FormulaR1C1 = "=IFERROR(HLOOKUP(RC10,'Macro & projects infos'!R6:R7,2,0),"""")"
    Cells(2, 21).FormulaR1C1 = "=IFERROR(VLOOKUP(RC10 & RC17,C53:C54,2,0),""Erreur : semaine non paramétrée"")"
    
    '2.Etirer les formules
    endingLineSechel = Cells(100000, 1).End(xlUp).Row
    
    If endingLineSechel > 2 Then
    
        'Cells(2, 1).AutoFill Destination:=Range(Cells(2, 1), Cells(endingLineSechel, 1)), Type:=xlFillDefault
        Cells(2, 15).AutoFill Destination:=Range(Cells(2, 15), Cells(endingLineSechel, 15)), Type:=xlFillDefault
        Cells(2, 16).AutoFill Destination:=Range(Cells(2, 16), Cells(endingLineSechel, 16)), Type:=xlFillDefault
        Cells(2, 17).AutoFill Destination:=Range(Cells(2, 17), Cells(endingLineSechel, 17)), Type:=xlFillDefault
        Cells(2, 18).AutoFill Destination:=Range(Cells(2, 18), Cells(endingLineSechel, 18)), Type:=xlFillDefault
        Cells(2, 19).AutoFill Destination:=Range(Cells(2, 19), Cells(endingLineSechel, 19)), Type:=xlFillDefault
        Cells(2, 20).AutoFill Destination:=Range(Cells(2, 20), Cells(endingLineSechel, 20)), Type:=xlFillDefault
        Cells(2, 21).AutoFill Destination:=Range(Cells(2, 21), Cells(endingLineSechel, 21)), Type:=xlFillDefault
    End If
    
    '3.Forcer le calcul
    Application.Calculate
    
    '4.Convertir les formules en valeurs brutes
    ' Range(Cells(2, 1), Cells(endingLineSechel, 1)).Value = Range(Cells(2, 1), Cells(endingLineSechel, 1)).Value
    Range(Cells(2, 15), Cells(endingLineSechel, 15)).Value = Range(Cells(2, 15), Cells(endingLineSechel, 15)).Value
    Range(Cells(2, 16), Cells(endingLineSechel, 16)).Value = Range(Cells(2, 16), Cells(endingLineSechel, 16)).Value
    Range(Cells(2, 17), Cells(endingLineSechel, 17)).Value = Range(Cells(2, 17), Cells(endingLineSechel, 17)).Value
    Range(Cells(2, 18), Cells(endingLineSechel, 18)).Value = Range(Cells(2, 18), Cells(endingLineSechel, 18)).Value
    Range(Cells(2, 19), Cells(endingLineSechel, 19)).Value = Range(Cells(2, 19), Cells(endingLineSechel, 19)).Value
    Range(Cells(2, 20), Cells(endingLineSechel, 20)).Value = Range(Cells(2, 20), Cells(endingLineSechel, 20)).Value
    Range(Cells(2, 21), Cells(endingLineSechel, 21)).Value = Range(Cells(2, 21), Cells(endingLineSechel, 21)).Value
    
    '5.Afficher toutes les lignes d'un TMC/mois (et d'une pièce) sur la même semaine, la première existante. Ex: S34-35-36. Seules des pièces sur 35 et 36 --> on note tout à S35.
        'en formule et non en code avec un For, bcp plus rapide
        
        '5.1. Isole le numéro de semaine --> 37 si CW37
        Cells(2, 22).FormulaR1C1 = "=LEFT(RC17,2)&RIGHT(RC17,LEN(RC17)-5)"
        
        If endingLineSechel > 2 Then
        
            Cells(2, 22).AutoFill Destination:=Range(Cells(2, 22), Cells(endingLineSechel, 22)), Type:=xlFillDefault
        Else
            Cells(2, 22).Value = Cells(2, 22).Value
        End If
        Application.Calculate
        
        Range(Cells(2, 22), Cells(endingLineSechel, 22)).Value = Range(Cells(2, 22), Cells(endingLineSechel, 22)).Value
        Columns(22).TextToColumns , DataType:=xlDelimited, _
                TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
                Semicolon:=True, Comma:=True, Space:=False, Other:=False, FieldInfo:= _
                Array(1, 1), TrailingMinusNumbers:=True
        
        '5.2. On récupère le numéro de semaine minimum avec comme critère Réf de la pièce et le TMC
               
        'On concatene les 3 critères pour simplifier les recherches
        Cells(2, 23).FormulaR1C1 = "=RC[-21]&RC[-19]&RC[-2]&RC[-16]"
        
        If endingLineSechel > 2 Then
            Cells(2, 23).AutoFill Destination:=Range(Cells(2, 23), Cells(endingLineSechel, 23)), Type:=xlFillDefault
        Else
            Cells(2, 23).Value = Cells(2, 23).Value
        End If
        Application.Calculate
        Range(Cells(2, 23), Cells(endingLineSechel, 23)).Value = Range(Cells(2, 23), Cells(endingLineSechel, 23)).Value
        
        
        Dim weekmin As Long
        Dim obj As Object
        Dim codesearched As String
        Dim jOld As Long
        Dim found As Boolean
        
        Columns(24).Clear
        Columns(24).NumberFormat = "@"
        'recherche et écriture
        For i = 2 To endingLineSechel
            jOld = 0
            weekmin = Cells(i, 22).Value
            codesearched = Cells(i, 23).Value
            
            On Error Resume Next 'si aucune n'est trouvé, on passe directement à l'écriture
            j = 0
            j = Columns(23).Find(What:=codesearched, LookAt:=xlWhole).Row
            On Error GoTo 0
            Set obj = Columns(23).Find(What:=codesearched, LookAt:=xlWhole) 'utilisation objet pour findnext
            
            While j > 0 And j > jOld And found = False
                If Cells(j, 22).Value < weekmin Then
                    weekmin = Cells(j, 22).Value
                End If
        
                jOld = j 'conserve l'ancien -> car quand on a fait le tour, Excel repart au 1er -> on garde que si j > jOld
                Set obj = Columns(23).FindNext(obj)
                j = obj.Row 'cherche le suivant
            Wend
            On Error GoTo 0
                Cells(i, 24).Value = weekmin
        Next i
        
    
        '5.3. On renote le nouveau CW
    
        Cells(2, 25).FormulaR1C1 = "=LEFT(RC24,2) & ""-CW"" & RIGHT(RC24,2)"
        If endingLineSechel > 2 Then
            Cells(2, 25).AutoFill Destination:=Range(Cells(2, 25), Cells(endingLineSechel, 25)), Type:=xlFillDefault
        Else
            Cells(2, 25).Value = Cells(2, 25).Value
        End If
        Application.Calculate
        Range(Cells(2, 25), Cells(endingLineSechel, 25)).Value = Range(Cells(2, 25), Cells(endingLineSechel, 25)).Value
    
    Columns("AX:BC").Select
    ' Selection.ClearContents
    Selection.EntireColumn.Delete xlToLeft
    
    
    Range("A1").Select
    Range("A1").CurrentRegion.Sort key1:=Range("D1"), order1:=xlAscending, dataoption1:=xlSortNormal, key2:=Range("B1"), order2:=xlAscending, dataoption2:=xlSortNormal, key3:=Range("H1"), order3:=xlAscending, dataoption3:=xlSortNormal, Header:=xlYes
    Range("A1").Select
    
    If Not ActiveSheet.AutoFilter Is Nothing Then
    Selection.AutoFilter
    Else
    End If
    
    'Restaure la donnée par défaut de la barre d'état
    Application.CutCopyMode = False
    Application.StatusBar = False

End Sub



Sub copie_reference_base()

    Application.ScreenUpdating = True
    Sheets("BASE").DisplayPageBreaks = False
    Sheets("BASE_temp").DisplayPageBreaks = False

    Sheets("RLF").Select
    If Not ActiveSheet.AutoFilter Is Nothing Then
        Selection.AutoFilter
    Else
    End If

    'parametrage
    Sheets("BASE").Select
    If Not ActiveSheet.AutoFilter Is Nothing Then
    Selection.AutoFilter
    Else
    End If
    
    'enlève les filtres
    Sheets("BASE_temp").Select
    If Not ActiveSheet.AutoFilter Is Nothing Then
    Selection.AutoFilter
    Else
    End If
    
    Sheets("BASE").Select
    If Not ActiveSheet.AutoFilter Is Nothing Then
    Selection.AutoFilter
    Else
    End If
    
    ' copier infos BASE dans BASE_temp
    Sheets("BASE_temp").Cells.ClearContents
    Sheets("BASE_temp").Cells.ClearFormats
    
    Sheets("BASE").Cells.Copy
    Sheets("BASE_temp").Cells(1, 1).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
    
    ' copie des infos de l'onglet Sechel
    Range("A3:BO" & Range("A1000000").End(xlUp).Row).Select
    Selection.ClearContents
    
    ' BASE optimise count of rows
    Debug.Print "adr to delete xlshiftup: " & ":A" & (Range("A1000000").End(xlUp).Row + 10) & "A1000000"
    Range("A3:A1000000").EntireRow.Delete xlShiftUp
    
    ' this formatting should be later
    'Selection.Interior.ColorIndex = xlColorIndexNone
    'Selection.ClearComments
    'Selection.NumberFormat = "General"
    
    'Range("A3:AY50000").Select
    'Selection.ClearFormats
    
    'Range("AR3:AS50000").Select
    'Selection.NumberFormat = "dd-mm-yyyy hh:mm:ss"
    Range("AR:AS").NumberFormat = "dd-mm-yyyy hh:mm:ss"


    ' copy data from sechel worksheet to empty BASE
    ' ---------------------------------------------------------------------------------------------------------------------------------------
    ' ---------------------------------------------------------------------------------------------------------------------------------------
    Sheets("sechel").Select
    der_lig_ref_s = Cells(500000, 2).End(xlUp).Row

    Range(Cells(prem_lig_ref_s, 2), Cells(der_lig_ref_s, 3)).Select
    Selection.Copy
    Sheets("BASE").Select
    Cells(prem_lig_ref_b, col_ref_b).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
    
    ' OLD LOGIC provide already calculated data - risk of making data confusion
    '
    'Sheets("sechel").Select
    'Range(Cells(prem_lig_ref_s, 15), Cells(der_lig_ref_s, 15)).Select
    'Selection.Copy
    'Sheets("BASE").Select
    'Cells(prem_lig_ref_b, col_qte_theo_b).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
    
    ' primary theo qty from input file
    ' ---------------------------------------------------------------------------------------------------------------------------------------
    Sheets("sechel").Select
    Range(Cells(prem_lig_ref_s, 13), Cells(der_lig_ref_s, 13)).Select
    Selection.Copy
    Sheets("BASE").Select
    Cells(prem_lig_ref_b, col_qte_theo_b).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    ' ---------------------------------------------------------------------------------------------------------------------------------------
    
    ' qty confirmed from input file
    ' ---------------------------------------------------------------------------------------------------------------------------------------
    Sheets("sechel").Select
    Range(Cells(prem_lig_ref_s, 14), Cells(der_lig_ref_s, 14)).Select
    Selection.Copy
    Sheets("BASE").Select
    Cells(prem_lig_ref_b, col_qte_theo_b + 1).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    ' ---------------------------------------------------------------------------------------------------------------------------------------
    
    Sheets("sechel").Select 'regroupement semaines
    Range(Cells(prem_lig_ref_s, 25), Cells(der_lig_ref_s, 25)).Select
    Selection.Copy
    Sheets("BASE").Select
    Cells(prem_lig_ref_b, col_base_ech).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
    Sheets("sechel").Select 'semaines échéancier
    Range(Cells(prem_lig_ref_s, 17), Cells(der_lig_ref_s, 17)).Select
    Selection.Copy
    Sheets("BASE").Select
    Cells(prem_lig_ref_b, 5).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
    Sheets("sechel").Select
    Range(Cells(prem_lig_ref_s, 7), Cells(der_lig_ref_s, 7)).Select
    Selection.Copy
    Sheets("BASE").Select
    Cells(prem_lig_ref_b, col_base_mag).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
    
    Sheets("sechel").Select
    Range(Cells(prem_lig_ref_s, 4), Cells(der_lig_ref_s, 4)).Select
    Selection.Copy
    Sheets("BASE").Select
    Cells(prem_lig_ref_b, col_cofor_b).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
    
    ' supplier name
    
    Sheets("sechel").Select
    Range(Cells(prem_lig_ref_s, 5), Cells(der_lig_ref_s, 5)).Select
    Selection.Copy
    Sheets("BASE").Select
    Cells(prem_lig_ref_b, col_base_fnr).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
    ' supplier name - blue

    Sheets("sechel").Select
    Range(Cells(prem_lig_ref_s, 5), Cells(der_lig_ref_s, 5)).Select
    Selection.Copy
    Sheets("BASE").Select
    Cells(prem_lig_ref_b, 7).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
    
    ' in xF no contract #
    ' ---------------------------------------------------------------------------------------------------------------------------------------
    'Sheets("sechel").Select
    'Range(Cells(prem_lig_ref_s, 16), Cells(der_lig_ref_s, 16)).Select
    'Selection.Copy
    'Sheets("BASE").Select
    'Cells(prem_lig_ref_b, col_po_numb_b).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    ' ---------------------------------------------------------------------------------------------------------------------------------------
    
    ' in xF no del prog
    ' ---------------------------------------------------------------------------------------------------------------------------------------
    'Sheets("sechel").Select
    'Range(Cells(prem_lig_ref_s, 8), Cells(der_lig_ref_s, 8)).Select
    'Selection.Copy
    'Sheets("BASE").Select
    'Cells(prem_lig_ref_b, col_del_prog_b).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    ' ---------------------------------------------------------------------------------------------------------------------------------------
    
    Sheets("sechel").Select
    Range(Cells(prem_lig_ref_s, 1), Cells(der_lig_ref_s, 1)).Select
    Selection.Copy
    Sheets("BASE").Select
    ' paste values removes leading zeros! NOK!
    ' Cells(prem_lig_ref_b, col_nom_appro_b).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Cells(prem_lig_ref_b, col_nom_appro_b).PasteSpecial Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
    Sheets("sechel").Select
    Range(Cells(prem_lig_ref_s, 18), Cells(der_lig_ref_s, 18)).Select
    Selection.Copy
    Sheets("BASE").Select
    Cells(prem_lig_ref_b, col_mail_appro_b).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
    Sheets("sechel").Select
    Range(Cells(prem_lig_ref_s, 20), Cells(der_lig_ref_s, 20)).Select
    Selection.Copy
    Sheets("BASE").Select
    Cells(prem_lig_ref_b, col_perimetre_b).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
    Sheets("sechel").Select
    Range(Cells(prem_lig_ref_s, 21), Cells(der_lig_ref_s, 21)).Select
    Selection.Copy
    Sheets("BASE").Select
    Cells(prem_lig_ref_b, col_tmc_b).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
    'Mise à jour GZA (Alten) : Suppression de la boucle principale et de ses sous-boucles (complexité Ncarré)
    'Remplacement par des formules
    '0.Forcer la référence en texte côté RLF
    '0.Forcer la référence en texte côté RLF
    Sheets("RLF").Columns(1).TextToColumns , DataType:=xlDelimited, _
            TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
            Semicolon:=True, Comma:=True, Space:=False, Other:=False, FieldInfo:= _
            Array(1, 2), TrailingMinusNumbers:=True
     
    Sheets("BASE").Select
    Range(Cells(3, 1), Cells(Cells(100000, 1).End(xlUp).Row, 1)).TextToColumns , DataType:=xlDelimited, _
            TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
            Semicolon:=True, Comma:=True, Space:=False, Other:=False, FieldInfo:= _
            Array(1, 2), TrailingMinusNumbers:=True
        


    
    
    '1.a.Ecrire formules sur la première ligne
    Cells(3, 13).FormulaR1C1 = "=IF(IFERROR(VLOOKUP(RC8,RLF!C5:C15,5,0),"""")=0,"""",IFERROR(VLOOKUP(RC8,RLF!C5:C15,5,0),""""))"
    Cells(3, 14).FormulaR1C1 = "=IF(IFERROR(VLOOKUP(RC8,RLF!C5:C15,6,0),"""")=0,"""",IFERROR(VLOOKUP(RC8,RLF!C5:C15,6,0),""""))"
    Cells(3, 15).FormulaR1C1 = "=IF(IFERROR(VLOOKUP(RC8,RLF!C5:C15,7,0),"""")=0,"""",IFERROR(VLOOKUP(RC8,RLF!C5:C15,7,0),""""))"
    
    '1.b.Etirer les formules
    Dim endingLine As Long
    endingLine = Cells(100000, 1).End(xlUp).Row

    If endingLine > 3 Then
        Cells(3, 13).AutoFill Destination:=Range(Cells(3, 13), Cells(endingLine, 13)), Type:=xlFillDefault
        Cells(3, 14).AutoFill Destination:=Range(Cells(3, 14), Cells(endingLine, 14)), Type:=xlFillDefault
        Cells(3, 15).AutoFill Destination:=Range(Cells(3, 15), Cells(endingLine, 15)), Type:=xlFillDefault
    Else
        Cells(3, 13).Value = Cells(3, 13).Value
        Cells(3, 14).Value = Cells(3, 14).Value
        Cells(3, 15).Value = Cells(3, 15).Value
    End If
    
    '1.c.Forcer le calcul
    Application.Calculate
    
    '1.d.Convertir les formules en valeurs brutes
    Range(Cells(3, 13), Cells(endingLine, 15)).Value = Range(Cells(3, 13), Cells(endingLine, 15)).Value 'les 3 colonnes d'un coup

    ' NOK!
    '1.e.Cas spécifique de la différenciation des noms de fournisseurs = on ajoute le nom du contact log SI il y a plus d'un contact log relié à ce FNR
    ' Range(Cells(3, 7), Cells(endingLine, 7)).Value = Range(Cells(3, 78), Cells(endingLine, 78)).Value
    
    
    ' identification des quantités confirmées et des TMCs de l'onglet temp
    'Mise à jour GZA (Alten)
    '2.a.Ecrire en dur les rassemblements des données 1)Ref+QtéTheo+Echéancier 2)Ref+Echéancier
    'L'objectif est de créer une clé de recherche unique côté base et sa correspondance côté temp
    Sheets("BASE_temp").Select
    
    ' left 10 digits is xP world
    ' Cells(3, 59).FormulaR1C1 = "=left(RC1,10)&RC5&RC" & col_base_mag
    
    ' take all REFERENCE
    Cells(3, 59).FormulaR1C1 = "=RC1&RC5&RC" & col_base_mag
    endingLine = Cells(100000, 1).End(xlUp).Row
    
    If endingLine > 3 Then
        Cells(3, 59).AutoFill Destination:=Range(Cells(3, 59), Cells(endingLine, 59)), Type:=xlFillDefault
    Else
        Cells(3, 59).Value = Cells(3, 59).Value
    End If
    
    Application.Calculate
    Range(Cells(3, 59), Cells(endingLine, 59)).Value = Range(Cells(3, 59), Cells(endingLine, 59)).Value
    Sheets("BASE").Select
    ' left 10 digits is xP world
    ' Cells(3, 59).FormulaR1C1 = "=left(RC1,10)&RC5&RC" & col_base_mag
    
    ' take all REFERENCE
    Cells(3, 59).FormulaR1C1 = "=RC1&RC5&RC" & col_base_mag
    endingLine = Cells(100000, 1).End(xlUp).Row


    If endingLine > 3 Then
        Cells(3, 59).AutoFill Destination:=Range(Cells(3, 59), Cells(endingLine, 59)), Type:=xlFillDefault
    Else
        Cells(3, 59).Value = Cells(3, 59).Value
    End If
    Application.Calculate
    Range(Cells(3, 59), Cells(endingLine, 59)).Value = Range(Cells(3, 59), Cells(endingLine, 59)).Value

    '2.b.Cas des quantités confirmées
    Sheets("BASE").Select
    endingLine = Cells(100000, 1).End(xlUp).Row
    
    
    ' this logic is NOK
    ' ----------------------------------------------------------------------------------------
    ' ----------------------------------------------------------------------------------------
    
    ' this formula destroying available information from input file from decorator
    ' Cells(3, 4).FormulaR1C1 = "=IF(IFERROR(INDEX(BASE_temp!R1:R" & endingLine & ",MATCH(RC59,BASE_temp!C59,0),4),"""")=0,"""",IFERROR(INDEX(BASE_temp!R1:R" & endingLine & ",MATCH(RC59,BASE_temp!C59,0),4),""""))"

    'Dim i As Long
    'Dim foundLine As Long
    'Dim critere As String
    'For i = 3 To endingLine
    '    critere = Cells(i, 59).Value
    '    On Error Resume Next
    '    foundLine = 0
    '    foundLine = Sheets("BASE_temp").Columns(59).Find(What:=critere, LookAt:=xlWhole).Row
    '    On Error GoTo 0
    '
    '    If foundLine > 0 Then
    '        Cells(i, 4).Value = Sheets("BASE_temp").Cells(foundLine, 4).Value
    '    End If
    'Next i
    ' ----------------------------------------------------------------------------------------
    ' ----------------------------------------------------------------------------------------
    
    ' new idea how to synchro confirmed qty required!
    ' ----------------------------------------------------------------------------------------
    ' basic idea is to have double synchro on confirmed qty
    ' first step is on the input file from decorator - but here is the thing - confirmed most
    ' probably comes from PUS file as well - so should really have this one?
    ' also the problem appears when we will have to big time gap between
    ' ----------------------------------------------------------------------------------------
    
    Dim iter1 As Long
    With ThisWorkbook.Sheets("BASE")
        For iter1 = 3 To endingLine
        
        
            If Trim(.Cells(iter1, 4).Value) <> "" Then
                If IsNumeric(.Cells(iter1, 4).Value) Then

                    If CLng(.Cells(iter1, 4).Value) > 0 Then
                        ' nop or check if we having the same data ?
                    End If
                    
                    If Trim(.Cells(iter1, 4).Value) = "0" Then
                        ' this scenario is zero from input file - check if maybe there is some confirmed qty in meantime
                        .Cells(iter1, 4).Value = assignConfQtyFrom_BASE_temp(iter1)
                    End If
                Else
                    .Cells(iter1, 4).Value = assignConfQtyFrom_BASE_temp(iter1)
                End If
            End If
        Next iter1
    End With
    
    
    ' ----------------------------------------------------------------------------------------
    
    
    
    ' ----------------------------------------------------------------------------------------




    '2.c. Ajout et report des infos : commentaires RPO, VOR et Acteur
    Sheets("BASE").Select
    Cells(3, col_base_commSechel).FormulaR1C1 = "=IF(IFERROR(INDEX(BASE_temp!C1:C59,MATCH(RC59,BASE_temp!C59,0)," & col_base_commSechel & "),"""")=0,"""",IFERROR(INDEX(BASE_temp!C1:C59,MATCH(RC59,BASE_temp!C59,0)," & col_base_commSechel & "),""""))"

    If endingLine > 3 Then
        Cells(3, col_base_commSechel).AutoFill Destination:=Range(Cells(3, col_base_commSechel), Cells(endingLine, col_base_commSechel)), Type:=xlFillDefault
    Else
        Cells(3, col_base_commSechel).Value = Cells(3, col_base_commSechel).Value
    End If
    Sheets("BASE").Calculate
    Range(Cells(3, col_base_commSechel), Cells(endingLine, col_base_commSechel)).Value = Range(Cells(3, col_base_commSechel), Cells(endingLine, col_base_commSechel)).Value
    
    Sheets("BASE").Select
    Cells(3, col_base_imputationCodeI).FormulaR1C1 = "=IF(IFERROR(INDEX(BASE_temp!C1:C59,MATCH(RC59,BASE_temp!C59,0)," & col_base_imputationCodeI & "),"""")=0,"""",IFERROR(INDEX(BASE_temp!C1:C59,MATCH(RC59,BASE_temp!C59,0)," & col_base_imputationCodeI & "),""""))"
    
    If endingLine > 3 Then
        Cells(3, col_base_imputationCodeI).AutoFill Destination:=Range(Cells(3, col_base_imputationCodeI), Cells(endingLine, col_base_imputationCodeI)), Type:=xlFillDefault
    Else
        Cells(3, col_base_imputationCodeI).Value = Cells(3, col_base_imputationCodeI).Value
    End If
    Sheets("BASE").Calculate
    Range(Cells(3, col_base_imputationCodeI), Cells(endingLine, col_base_imputationCodeI)).Value = Range(Cells(3, col_base_imputationCodeI), Cells(endingLine, col_base_imputationCodeI)).Value
    
    Sheets("BASE").Select
    Cells(3, col_base_commRPO).FormulaR1C1 = "=IF(IFERROR(INDEX(BASE_temp!C1:C59,MATCH(RC59,BASE_temp!C59,0)," & col_base_commRPO & "),"""")=0,"""",IFERROR(INDEX(BASE_temp!C1:C59,MATCH(RC59,BASE_temp!C59,0)," & col_base_commRPO & "),""""))"
    If endingLine > 3 Then
        Cells(3, col_base_commRPO).AutoFill Destination:=Range(Cells(3, col_base_commRPO), Cells(endingLine, col_base_commRPO)), Type:=xlFillDefault
    Else
        Cells(3, col_base_commRPO).Value = Cells(3, col_base_commRPO).Value
    End If
    Sheets("BASE").Calculate
    Range(Cells(3, col_base_commRPO), Cells(endingLine, col_base_commRPO)).Value = Range(Cells(3, col_base_commRPO), Cells(endingLine, col_base_commRPO)).Value
    
    Sheets("BASE").Select
    Cells(3, col_base_VOR).FormulaR1C1 = "=IF(IFERROR(INDEX(BASE_temp!C1:C59,MATCH(RC59,BASE_temp!C59,0)," & col_base_VOR & "),"""")=0,"""",IFERROR(INDEX(BASE_temp!C1:C59,MATCH(RC59,BASE_temp!C59,0)," & col_base_VOR & "),""""))"
    If endingLine > 3 Then
        Cells(3, col_base_VOR).AutoFill Destination:=Range(Cells(3, col_base_VOR), Cells(endingLine, col_base_VOR)), Type:=xlFillDefault
    Else
        Cells(3, col_base_VOR).Value = Cells(3, col_base_VOR).Value
    End If
    Sheets("BASE").Calculate
    Range(Cells(3, col_base_VOR), Cells(endingLine, col_base_VOR)).Value = Range(Cells(3, col_base_VOR), Cells(endingLine, col_base_VOR)).Value
    
    Sheets("BASE").Select
    Cells(3, col_base_acteur).FormulaR1C1 = "=IF(IFERROR(INDEX(BASE_temp!C1:C59,MATCH(RC59,BASE_temp!C59,0)," & col_base_acteur & "),"""")=0,"""",IFERROR(INDEX(BASE_temp!C1:C59,MATCH(RC59,BASE_temp!C59,0)," & col_base_acteur & "),""""))"
    endingLine = Cells(100000, 1).End(xlUp).Row
    If endingLine > 3 Then
        Cells(3, col_base_acteur).AutoFill Destination:=Range(Cells(3, col_base_acteur), Cells(endingLine, col_base_acteur)), Type:=xlFillDefault
    Else
        Cells(3, col_base_acteur).Value = Cells(3, col_base_acteur).Value
    End If
    Sheets("BASE").Calculate
    Range(Cells(3, col_base_acteur), Cells(endingLine, col_base_acteur)).Value = Range(Cells(3, col_base_acteur), Cells(endingLine, col_base_acteur)).Value
    
    
    
    
    '2.d.Nettoyer construction
    Sheets("BASE").Select
    Range(Cells(3, 59), Cells(endingLine, 59)).ClearContents
    Range(Cells(3, 60), Cells(endingLine, 60)).ClearContents
    Sheets("BASE_temp").Select
    Range(Cells(3, 59), Cells(endingLine, 59)).ClearContents
    Range(Cells(3, 60), Cells(endingLine, 60)).ClearContents


    ''''''''''''''''''''''''''''''''''''''''''''''''''
    ' dans la colonne des quantités confirmées : filtre sur les vides et suppression (au cas où une cellule ne possède qu'un espace)
    Sheets("BASE").Select
    If Not ActiveSheet.AutoFilter Is Nothing Then
    Selection.AutoFilter
    Else
    End If
    
    Range(Cells(prem_lig_ref_b - 1, col_ref_b), Cells(Cells(500000, col_ref_b).End(xlUp).Row, Cells(2, 1).End(xlToRight).Column)).Select
    Selection.AutoFilter
    Range("A1").Select
    
    
    ThisWorkbook.Sheets("BASE").Range("A" & (endingLine + 10) & ":A1000000").EntireRow.Delete xlShiftUp
    
    
    
    ' hmmm...??
    'Sheets("BASE").Cells.AutoFilter Field:=col_qte_conf_b, Criteria1:=""
    '
    'Dim Maplagevisible As Range
    'Set Maplagevisible = Range("D2", Cells(Rows.count, "D").End(xlUp)).SpecialCells(xlCellTypeVisible)
    'Maplagevisible.Select
    'If Maplagevisible.Cells(1, 1) = "QUANTITE confirmée" Then
    'Range("D2", Cells(Rows.count, "D").End(xlUp)).SpecialCells(xlCellTypeVisible).Select
    'Selection.ClearContents
    'Cells(2, 4) = "QUANTITE confirmée"
    'Else: End If
    
    ' test si une quantité confirmée ou une capacité d'UC n'est pas une valeur numérique
    Sheets("BASE").Select
    If Not ActiveSheet.AutoFilter Is Nothing Then
    Selection.AutoFilter
    Else
    End If
    
    ' identification des fournisseurs avec même nom et même cofor mais avec interlocuteurs différents
    Sheets("BASE").Select
    If Not ActiveSheet.AutoFilter Is Nothing Then
    Selection.AutoFilter
    Else
    End If
    
    Dim found As Boolean
    found = False
    ' test si un contact fnr est vide
    For i = 2 To Cells(500000, col_nom_log_b).End(xlUp).Row
         If Cells(i, col_nom_log_b) <> "" And Cells(i, col_nom_log_b) <> 0 Then
         Else
            found = True
            Exit For
        End If
    Next i
    
    If found = True Then
        If (MsgBox("Au moins un contact fournisseur n'est pas renseigné." & Chr(10) & "Voulez-vous poursuivre le process ?", vbYesNo) = vbYes) Then
        Else
            'Restaure la donnée par défaut de la barre d'état
            Application.StatusBar = False
            Sheets("BASE").Select
            End
        End If
    End If
    
    
    'Restaure la donnée par défaut de la barre d'état
    Application.StatusBar = False

End Sub



Private Function assignConfQtyFrom_BASE_temp(miter1 As Long) As Variant
    
    assignConfQtyFrom_BASE_temp = 0
    
    Dim tmp As Worksheet, bs
    Set tmp = ThisWorkbook.Sheets("BASE_temp")
    Set bs = ThisWorkbook.Sheets("BASE")
    
    
    Dim BG_column As Long
    BG_column = tmp.Range("BG1").Column
    
    
    Dim tmpr1 As Range
    Set tmpr1 = tmp.Cells(3, BG_column)
    Do
        
        If tmpr1.Value = bs.Cells(miter1, BG_column).Value Then
            
            If IsNumeric(tmp.Cells(tmpr1.Row, 4).Value) Then
                If tmp.Cells(tmpr1.Row, 4).Value > 0 Then
                    assignConfQtyFrom_BASE_temp = tmp.Cells(tmpr1.Row, 4).Value
                End If
            End If
        End If
        
        Set tmpr1 = tmpr1.Offset(1, 0)
    Loop Until Trim(tmpr1.Value) = ""
    
    
    
    
End Function


Sub copie_informations_cpl()

    Application.ScreenUpdating = True


    'Mise à jour GZA (Alten)
    '1.Ecrire en dur les rassemblements des données Ref+Echéancier+COFOR
    'L'objectif est de créer une clé de recherche unique côté base et sa correspondance côté temp
    Sheets("BASE_temp").Select
    Cells(3, 59).FormulaR1C1 = "=RC1&RC5&RC" & col_base_mag
    Dim endingLine As Long
    endingLine = Cells(100000, 1).End(xlUp).Row
    Cells(3, 59).AutoFill Destination:=Range(Cells(3, 59), Cells(endingLine, 59)), Type:=xlFillDefault
    Sheets("BASE_temp").Calculate
    Range(Cells(3, 59), Cells(endingLine, 59)).Value = Range(Cells(3, 59), Cells(endingLine, 59)).Value
    Sheets("BASE").Select
    Cells(3, 59).FormulaR1C1 = "=left(RC1,10)&RC5&RC" & col_base_mag
    endingLine = Cells(100000, 1).End(xlUp).Row
    
    If endingLine > 3 Then
        Cells(3, 59).AutoFill Destination:=Range(Cells(3, 59), Cells(endingLine, 59)), Type:=xlFillDefault
    Else
        Cells(3, 59).Value = Cells(3, 59).Value
    End If
    
    Sheets("BASE").Calculate
    Range(Cells(3, 59), Cells(endingLine, 59)).Value = Range(Cells(3, 59), Cells(endingLine, 59)).Value

    '2.Calculer puis figer successivement les colonnes
    Sheets("BASE").Select

    For i = 17 To 52
        If i <> 23 Then
            Cells(3, i).FormulaR1C1 = "=IF(IFERROR(INDEX(BASE_temp!C1:C59,MATCH(RC59,BASE_temp!C59,0)," & i & "),"""")=0,"""",IFERROR(INDEX(BASE_temp!C1:C59,MATCH(RC59,BASE_temp!C59,0)," & i & "),""""))"
            endingLine = Cells(100000, 1).End(xlUp).Row
            
            If endingLine > 3 Then
                Cells(3, i).AutoFill Destination:=Range(Cells(3, i), Cells(endingLine, i)), Type:=xlFillDefault
            Else
                Cells(3, i).Value = Cells(3, i).Value
            End If
            Sheets("BASE").Calculate
            Range(Cells(3, i), Cells(endingLine, i)).Value = Range(Cells(3, i), Cells(endingLine, i)).Value
        End If
    Next i

    '4.Nettoyer construction
    Sheets("BASE").Select
    Range(Cells(3, 59), Cells(endingLine, 59)).ClearContents
    Sheets("BASE_temp").Select
    Range(Cells(3, 59), Cells(endingLine, 59)).ClearContents
    
    
    Sheets("BASE").Select
    
    
    ' repositionnement des commentaires
     Dim cmt As Comment
     For Each cmt In ActiveSheet.Comments
        cmt.Shape.Top = cmt.Parent.Top + 5
        cmt.Shape.Left = cmt.Parent.Offset(0, 1).Left + 5
        cmt.Shape.TextFrame.AutoSize = True
     Next
    
    
    Columns("CA:CC").Select
    Selection.Delete Shift:=xlToLeft
    
    'Restaure la donnée par défaut de la barre d'état
    Application.StatusBar = False

End Sub

Sub liste_pickupsheet_a_mettre_a_jour()
    Application.ScreenUpdating = False
    lig_ref_ac = prem_lig_ref_ac
    Sheets("PUS creation").Select
    If ActiveSheet.AutoFilter Is Nothing Then
        Sheets("PUS creation").Range("B6:G6").AutoFilter
    Else
        ' nop
    End If

    Range(Cells(lig_ref_ac, col_noa_ac), Cells(Cells(500000, col_noa_ac).End(xlUp).Row, col_prog_liv_ac)).Select
    Selection.ClearContents
    Selection.Interior.ColorIndex = xlColorIndexNone
    Selection.Borders.Value = 0

    Sheets("sechel").Visible = True
    Sheets("sechel").Select
    If Not ActiveSheet.AutoFilter Is Nothing Then
        Selection.AutoFilter
    Else
    End If


    ' ??????
    ' TEMP_REMARK_FOR_DEMO
    ' Sheets("sechel").Cells.AutoFilter Field:=19, Criteria1:="oui"

    Dim Maplage As Range
    Set Maplage = Sheets("sechel").UsedRange.SpecialCells(xlCellTypeVisible)
    Dim Ligne As Range
    
    Dim tmpTxtNoa As String, tmpTxtCofor As String, tmpTxtRef As String
    tmpTxtNoa = ""
    tmpTxtCofor = ""
    tmpTxtRef = ""
    
    For Each Ligne In Maplage.Rows
        If Ligne.Row > 1 Then
            If Ligne.Cells(col_noa_s).Value = "GAc_Nom_NOA" Then
            '
            Else
                Sheets("PUS creation").Select
                tmpTxtNoa = CStr(Ligne.Cells(col_noa_s).Value)
                tmpTxtRef = CStr(Ligne.Cells(col_ref_s).Value)
                Cells(lig_ref_ac, col_noa_ac).Value = tmpTxtNoa
                Cells(lig_ref_ac, col_ref_ac).Value = CStr(tmpTxtRef)
                Cells(lig_ref_ac, col_desi_ac) = Ligne.Cells(col_desi_s).Value
                tmpTxtCofor = CStr(Ligne.Cells(col_cofor_s).Value)
                Cells(lig_ref_ac, col_cofor_ac).Value = tmpTxtCofor
                Cells(lig_ref_ac, col_fnr_ac).Value = Ligne.Cells(col_fnr_s).Value
                ' mag7 = 7
                ' Cells(lig_ref_ac, col_prog_liv_ac) = "" & Ligne.Cells(7).Value & "_" & Ligne.Cells(col_cofor_s).Value
                Cells(lig_ref_ac, col_prog_liv_ac) = "" & Ligne.Cells(7).Value & "_" & _
                    Ligne.Cells(4).Value & "_" & Ligne.Cells(2).Value
                '
                Rows(lig_ref_ac).Select
                Selection.RowHeight = 15
                lig_ref_ac = lig_ref_ac + 1
            End If
        End If
    Next
    
    ' this one is NOK becuase we are missing info about col_prog_liv_ac
    ' --------------------------------------------------------------------------------------------------------------------------
    ' TEMP_REMARK_FOR_DEMO
    'Range(Cells(prem_lig_ref_ac - 1, col_noa_ac), Cells(Cells(500000, col_noa_ac).End(xlUp).Row, col_prog_liv_ac)).Select
    'Selection.RemoveDuplicates Columns:=Array(6), Header:=xlYes
    ' --------------------------------------------------------------------------------------------------------------------------
    
    Range(Cells(prem_lig_ref_ac - 1, col_noa_ac), Cells(Cells(500000, col_noa_ac).End(xlUp).Row, col_prog_liv_ac)).Select
    Selection.Borders.Value = 1
    Selection.Sort key1:=Cells(prem_lig_ref_ac - 1, col_noa_ac), order1:=xlAscending, dataoption1:=xlSortNormal, key2:=Cells(prem_lig_ref_ac - 1, col_cofor_ac), order2:=xlAscending, dataoption2:=xlSortNormal, key3:=Cells(prem_lig_ref_ac - 1, col_desi_ac), order3:=xlAscending, dataoption3:=xlSortNormal, Header:=xlYes
    
    '''' création des formes
    
    For Each x In ActiveSheet.Shapes
        If x.AutoShapeType = msoShapeRectangle Then x.Delete
    Next x
    
    For i = prem_lig_ref_ac To Cells(500000, col_ref_ac).End(xlUp).Row
        Cells(i, col_prog_liv_ac + 1).Activate
        With ActiveSheet.Shapes.AddShape(msoShapeRectangle, 5, 5, 20, 5)
           .Left = ActiveCell.Left
           .Top = ActiveCell.Top
           .Height = Range(ActiveCell, ActiveCell.Offset(0.5, 0)).Height
           .OnAction = "reference_pickup"
        End With
    Next i
    
    ' renunmération des shapes par ligne
    num_shape = prem_lig_ref_ac
    For Each x In ActiveSheet.Shapes
        If x.AutoShapeType = msoShapeRectangle Then
            x.Name = num_shape
        num_shape = num_shape + 1
        Else: End If
    Next x
    
    ' prise en compte des modifications des noms des fournisseurs si plusieurs interlocuteurs
    ' this loop is becuase 1 step in this sub is to take data from sechel worksheet
    ' and now we are coming back to the BASE and re-asign data (supplier name)
    Sheets("PUS creation").Select
    For i = prem_lig_ref_ac To Cells(500000, col_ref_ac).End(xlUp).Row
        Sheets("PUS creation").Select
        reference_ac = Cells(i, col_ref_ac)
        Sheets("BASE").Select
            For j = prem_lig_ref_b To Cells(500000, col_ref_b).End(xlUp).Row
                If CStr(Cells(j, col_ref_b).Value) = CStr(reference_ac) Then
                reference = Cells(j, col_nom_fnr_b)
                Sheets("PUS creation").Select
                Cells(i, col_fnr_ac) = reference
                Exit For
                Else: End If
            Next j
    Next i
    
    ' nice and auto - but for some reason i deleted labels and this line below wont work
    ' Range(Cells(prem_lig_ref_ac - 1, col_noa_ac), Cells(Cells(500000, col_noa_ac).End(xlUp).Row, Cells(2, 1).End(xlToRight).Column)).Select
    'quick and dirty solution
    Range("B6:G6").Select
    Selection.AutoFilter
    Range("A1").Select
    
    Sheets("sechel").Select
    If Not ActiveSheet.AutoFilter Is Nothing Then
    Selection.AutoFilter
    Else
    End If
    
    Sheets("PUS creation").Select
    
    Sheets("sechel").Visible = False
    Sheets("progliv").Visible = False
    Sheets("RLF").Visible = False
    Sheets("ME9E").Visible = False
    Sheets("BASE_temp").Visible = False

End Sub

Sub list_deroul_fnr()

    Application.ScreenUpdating = False
    Sheets("PUS creation").Select
    Columns(100).Select
    Selection.ClearContents
    Sheets("BASE").Select
    If Not ActiveSheet.AutoFilter Is Nothing Then
    Selection.AutoFilter
    Else
    End If
    der_lig_ref_s = Cells(500000, 2).End(xlUp).Row
    Range(Cells(prem_lig_ref_b - 1, col_nom_fnr_b), Cells(Cells(500000, col_nom_fnr_b).End(xlUp).Row, col_nom_fnr_b)).Select
    Selection.Copy
    Sheets("PUS creation").Select
    Cells(1, 100).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

    ActiveSheet.Range(Cells(1, 100), Cells(Cells(500000, 100).End(xlUp).Row, 100)).RemoveDuplicates Columns:=1, Header:=xlYes
    ActiveWorkbook.Worksheets("PUS creation").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("PUS creation").Sort.SortFields.Add Key:=Range(Cells(1, 100), Cells(Cells(500000, 100).End(xlUp).Row, 100)) _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
With ActiveWorkbook.Worksheets("PUS creation").Sort
    .SetRange Range(Cells(1, 100), Cells(Cells(500000, 100).End(xlUp).Row, 100))
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With

If test_ref_style = "XL-A1" Then
'
ElseIf test_ref_style = "XL-R1C1" Then
Application.ReferenceStyle = xlR1C1
Else: End If

' on initialise le menu déroulant avec le premier fournisseur de la liste (ou cas où le fnr de la liste n'existe plus)
Cells(lig_cofor_rond, col_fnr_ac) = Cells(2, 100)

End Sub

Sub list_deroul_appro()
Sheets("PUS creation").Select
Columns(101).Select
Selection.ClearContents

Sheets("BASE").Select
der_lig_ref_s = Cells(500000, 2).End(xlUp).Row
Range(Cells(prem_lig_ref_b - 1, col_nom_appro_b), Cells(Cells(500000, col_nom_appro_b).End(xlUp).Row, col_nom_appro_b)).Select
Selection.Copy
Sheets("PUS creation").Select
Cells(1, 101).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

ActiveSheet.Range(Cells(1, 101), Cells(Cells(500000, 101).End(xlUp).Row, 101)).RemoveDuplicates Columns:=1, Header:=xlYes
ActiveWorkbook.Worksheets("PUS creation").Sort.SortFields.Clear
ActiveWorkbook.Worksheets("PUS creation").Sort.SortFields.Add Key:=Range(Cells(1, 101), Cells(Cells(500000, 101).End(xlUp).Row, 101)) _
    , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
With ActiveWorkbook.Worksheets("PUS creation").Sort
    .SetRange Range(Cells(1, 101), Cells(Cells(500000, 101).End(xlUp).Row, 101))
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With

If test_ref_style = "XL-A1" Then
'
ElseIf test_ref_style = "XL-R1C1" Then
Application.ReferenceStyle = xlR1C1
Else: End If

' on initialise le menu déroulant avec le premier appro de la liste (ou cas où l'appro de la liste n'existe plus)
Cells(lig_cofor_rond, col_noa_ac) = Cells(2, 101)

End Sub
Sub list_appro_fnr_cofor()
Application.ScreenUpdating = False
Sheets("PUS creation").Select
Columns(102).Select
Selection.ClearContents

Sheets("BASE").Select
der_lig_ref_s = Cells(500000, 2).End(xlUp).Row
Range(Cells(prem_lig_ref_b - 1, col_nom_fnr_b), Cells(Cells(500000, col_nom_fnr_b).End(xlUp).Row, col_nom_fnr_b)).Select
Selection.Copy
Sheets("PUS creation").Select
Cells(1, 102).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
Sheets("BASE").Select
Range(Cells(prem_lig_ref_b - 1, col_nom_appro_b), Cells(Cells(500000, col_nom_appro_b).End(xlUp).Row, col_nom_appro_b)).Select
Selection.Copy
Sheets("PUS creation").Select
Cells(1, 103).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
Sheets("BASE").Select
Range(Cells(prem_lig_ref_b - 1, col_cofor_b), Cells(Cells(500000, col_cofor_b).End(xlUp).Row, col_cofor_b)).Select
Selection.Copy
Sheets("PUS creation").Select
Cells(1, 104).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

For i = 2 To Cells(500000, 102).End(xlUp).Row
Cells(i, 102) = Cells(i, 102) & Cells(i, 103) & Cells(i, 104)
Next i

Columns(103).Select
Selection.ClearContents
Columns(104).Select
Selection.ClearContents


If Cells(1, 102) <> "" Then
ActiveSheet.Range(Cells(1, 102), Cells(Cells(500000, 102).End(xlUp).Row, 102)).RemoveDuplicates Columns:=1, Header:=xlYes
Else
ActiveSheet.Range(Cells(2, 102), Cells(Cells(500000, 102).End(xlUp).Row, 102)).RemoveDuplicates Columns:=1, Header:=xlYes
End If

ActiveWorkbook.Worksheets("PUS creation").Sort.SortFields.Clear
ActiveWorkbook.Worksheets("PUS creation").Sort.SortFields.Add Key:=Range(Cells(1, 102), Cells(Cells(500000, 102).End(xlUp).Row, 102)) _
    , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
With ActiveWorkbook.Worksheets("PUS creation").Sort
    .SetRange Range(Cells(1, 102), Cells(Cells(500000, 102).End(xlUp).Row, 102))
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With

If test_ref_style = "XL-A1" Then
'
ElseIf test_ref_style = "XL-R1C1" Then
Application.ReferenceStyle = xlR1C1
Else: End If

End Sub

Sub list_appro_par_fnr()
Application.ScreenUpdating = False
Sheets("PUS creation").Select
Columns(110).Select
Selection.ClearContents
Columns(111).Select
Selection.ClearContents
Cells(1, 112) = "concatenation"

Sheets("BASE").Select
der_lig_ref_s = Cells(500000, 2).End(xlUp).Row
Range(Cells(prem_lig_ref_b - 1, col_nom_appro_b), Cells(Cells(500000, col_nom_appro_b).End(xlUp).Row, col_nom_appro_b)).Select
Selection.Copy
Sheets("PUS creation").Select
Cells(1, 110).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

Sheets("BASE").Select
Range(Cells(prem_lig_ref_b - 1, col_nom_fnr_b), Cells(Cells(500000, col_nom_fnr_b).End(xlUp).Row, col_nom_fnr_b)).Select
Selection.Copy
Sheets("PUS creation").Select
Cells(1, 111).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

For i = 2 To Cells(500000, 110).End(xlUp).Row
Cells(i, 112) = Cells(i, 110) & Cells(i, 111)
Next i

ActiveSheet.Range(Cells(1, 110), Cells(Cells(500000, 110).End(xlUp).Row, 112)).RemoveDuplicates Columns:=3, Header:=xlYes
ActiveWorkbook.Worksheets("PUS creation").Sort.SortFields.Clear
ActiveWorkbook.Worksheets("PUS creation").Sort.SortFields.Add Key:=Range(Cells(1, 110), Cells(Cells(500000, 110).End(xlUp).Row, 110)) _
    , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
ActiveWorkbook.Worksheets("PUS creation").Sort.SortFields.Add Key:=Range(Cells(1, 111), Cells(Cells(500000, 111).End(xlUp).Row, 111)) _
    , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
With ActiveWorkbook.Worksheets("PUS creation").Sort
    .SetRange Range(Cells(1, 110), Cells(Cells(500000, 112).End(xlUp).Row, 112))
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With

Columns(112).Select
Selection.ClearContents

If test_ref_style = "XL-A1" Then
'
ElseIf test_ref_style = "XL-R1C1" Then
Application.ReferenceStyle = xlR1C1
Else: End If

Sheets("BASE").Select
If Not ActiveSheet.AutoFilter Is Nothing Then
Selection.AutoFilter
Else
End If
Range(Cells(prem_lig_ref_b - 1, col_ref_b), Cells(Cells(500000, col_ref_b).End(xlUp).Row, Cells(2, 1).End(xlToRight).Column)).Select
Selection.AutoFilter
Range("A1").Select

Sheets("PUS creation").Select
Cells(1, 1).Select

End Sub


Sub majNomFNR__xF()
    
    
    Sheets("BASE").Select
    If Not ActiveSheet.AutoFilter Is Nothing Then
    Selection.AutoFilter
    Else
    End If


    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    Sheets("BASE").Activate
    'écrire combinaisons FNR/mag/type pièce
    For i = 3 To Sheets("BASE").Cells(500000, col_nom_fnr_b).End(xlUp).Row
    
        ' Debug.Print CStr(Sheets("BASE").Cells(i, 6).Value) & CStr(Sheets("BASE").Cells(i, col_base_ech).Value) & CStr(Sheets("BASE").Cells(i, 23).Value)
        Sheets("BASE").Cells(i, 90).Value = Sheets("BASE").Cells(i, 6).Value & Sheets("BASE").Cells(i, col_base_ech).Value & Sheets("BASE").Cells(i, 23).Value
    Next i

    Dim lastLine As Long
    lastLine = Sheets("BASE").Cells(500000, 1).End(xlUp).Row

    'For i = 3 To lastLine
    '    If majGlobale = False Then
    '        Application.StatusBar = "traitement lignes BASE : " & i & " sur " & lastLine
    '    End If
    '
    '    'nom contact log
    '    Dim nom_contact_log As String
    '    nom_contact_log = Sheets("BASE").Cells(i, col_nom_log_b).Value
    '    If nom_contact_log <> "" Then
    '        nom_contact_log = " - " & nom_contact_log
    '    End If
    '
    '    'cofor exp
    '    Dim cofor_exp As String
    '    nom_cofor_exp_b = Sheets("BASE").Cells(i, col_cofor_exp_b)
    '    If nom_cofor_exp_b <> "" Then
    '        cofor_exp = " - COFOR EXP " & Cells(i, col_cofor_exp_b)
    '        Else
    '        cofor_exp = " - COFOR EXP " & "to precise"
    '    End If
    '
    '    'Type pièce
    '   ' Dim type_piece As String
    '    'type_piece = Sheets("BASE").Cells(i, 23).Value
    '
    '    'nom final
    '    ' Sheets("BASE").Cells(i, col_nom_fnr_b) = Sheets("BASE").Cells(i, col_base_fnr).Value & nom_contact_log & cofor_exp
    'Next i
    
    If majGlobale = False Then
        Application.StatusBar = "Etape 2 : mise à jour feuille PUS Creation"
        
    End If
    
    liste_pickupsheet_a_mettre_a_jour 'MaJ feuille PUS Creation
    list_deroul_fnr
    list_deroul_appro
    list_appro_fnr_cofor
    list_appro_par_fnr
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = ""
    
    
    If majGlobale = False Then
        MsgBox "Mise à jour du nom des fournisseurs terminée."
    End If
End Sub




