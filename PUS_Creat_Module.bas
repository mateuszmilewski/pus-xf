Attribute VB_Name = "PUS_Creat_Module"
Option Explicit

' note:
' slight changes in this version on 2020-07-21
' emergency implementation - issue with order with bigger parts list - macro error!
' fckg no-pro :)

' change in lines ~138 - wrong clearing selections!

' added private sub and function (visible only in scope of this module) - previous code was not working.
' On bottom:
' findRow
' clearBreaks

'
' used in lines:
' ~ 326 for findRow
' ~ > 420 managing h page breaks...




' next sub version 2.23 - added 6th static page




' second fix beginning comment
' --------------------------------------------------
' next sub version 2.24 - big change from line 353:
' tmc issue - dictionary solution!
' more private function on bottom of this module
' for TMC management
' --------------------------------------------------


Global Const G_FIRST_NUMBER_ROW = 33

Public testPUS As Long

Public Enum E_PUS_LIST_FOR_ARTICLES
    E_PUS_LIST_NUMBER = 3
    E_PUS_LIST_NAME
    E_PUS_LIST_PU_WEEK
    E_PUS_LIST_PU_DATE
    E_PUS_LIST_TMC
    E_PUS_LIST_CONTRACT_NM
    E_PUS_LIST_DELIVERY_PROGRAM
    E_PUS_LIST_QTY
    E_PUS_LIST_PCKG_NUM
    E_PUS_LIST_DIM
    E_PUS_LIST_WEIGHT
    E_PUS_LIST_QTY_IN_CONT
End Enum



Sub reference_pickup()
    '''''''''''''''''''''''''test existence nom appro et nom fournisseur

    testPUS = 1
    
    Sheets("BASE").Select
    Columns("A:BA").Select
    Range("A2").Activate
    Selection.EntireColumn.Hidden = False
    
    Sheets("PUS creation").Select
    If ActiveSheet.Shapes(Application.Caller).AutoShapeType = msoShapeRoundedRectangle Then
        noa_a = Cells(lig_cofor_rond, col_noa_ac)
        cofor_a = Cells(lig_cofor_rond, col_cofor_ac)
        nom_fnr_a = Cells(lig_cofor_rond, col_fnr_ac)
        perimetre_analyse = "per_cofor"
    Else
        lig_cofor_carre = ActiveSheet.Shapes(Application.Caller).TopLeftCell.Row
        reference_a = Cells(lig_cofor_carre, col_ref_ac)
        noa_a = Cells(lig_cofor_carre, col_noa_ac)
        cofor_a = Cells(lig_cofor_carre, col_cofor_ac)
        nom_fnr_a = Cells(lig_cofor_carre, col_fnr_ac)
        perimetre_analyse = "per_reference"
    End If
    
    index_fnr = 0
    index_noa = 0
    index_noa_fnr = 0

    ' les variables ci dessous sont utilisée en cas de filtre mis par qq sur le tableau ME9E
    ' exemple, un noa veut lister ses références. A ce moment le filtre limite les colonnes 100,101 et 102
    nb_lig_fnr = 2
    nb_lig_noa = 2
    nb_lig_noa_fnr = 2
    
    Do While Cells(nb_lig_fnr, 100) <> ""
        nb_lig_fnr = nb_lig_fnr + 1
    Loop
    
    Do While Cells(nb_lig_noa, 101) <> ""
        nb_lig_noa = nb_lig_noa + 1
    Loop
    
    Do While Cells(nb_lig_noa_fnr, 102) <> ""
        nb_lig_noa_fnr = nb_lig_noa_fnr + 1
    Loop

    For i = 2 To nb_lig_fnr - 1
    
        If UCase(Cells(i, 100)) = UCase(nom_fnr_a) Then
            index_fnr = 0
        Exit For
        Else
            index_fnr = 1
        End If
    Next i

    For j = 2 To nb_lig_noa - 1
    
        If UCase(Cells(j, 101)) = UCase(noa_a) Then
            index_noa = 0
        Exit For
        Else
            index_noa = 1
        End If
    Next j

    For k = 2 To nb_lig_noa_fnr - 1
    
        If UCase(Cells(k, 102)) = UCase(nom_fnr_a & noa_a & cofor_a) Then
            index_noa_fnr = 0
        Exit For
        Else
            index_noa_fnr = 1
        End If
    Next k

    '' si les test sont validés, on continue
    Application.ScreenUpdating = False
    
    '''''''''''''''''''''''''parametrage
    Sheets("PUS creation").Select ' accueil

    ' changement des couleurs des puces carrées si clic sur puce carrée
    'If ActiveSheet.Shapes(Application.Caller).AutoShapeType = msoShapeRectangle Then
    '    If ActiveSheet.Shapes(Application.Caller).Fill.ForeColor.RGB = RGB(144, 238, 144) Then 'couleur verte
    '    ActiveSheet.Shapes(Application.Caller).Fill.ForeColor.RGB = RGB(91, 155, 213) ' on passe bleu et rien d'autre
    '    Exit Sub
    '    ElseIf ActiveSheet.Shapes(Application.Caller).Fill.ForeColor.RGB <> RGB(144, 238, 144) Then 'différent de vert
    '    ActiveSheet.Shapes(Application.Caller).Fill.ForeColor.RGB = RGB(144, 238, 144) ' on passe vert et on crée la pick up sheet
    '    Else: End If
    'Else: End If

    ' changement des couleurs des puces carrées si clic sur puce ronde
    'If ActiveSheet.Shapes(Application.Caller).AutoShapeType = msoShapeRoundedRectangle Then
    '    For i = prem_lig_ref_ac To Cells(500000, col_cofor_ac).End(xlUp).Row
    '    If UCase((Cells(i, col_cofor_ac) & Cells(i, col_fnr_ac) & Cells(i, col_noa_ac))) = UCase((Cells(lig_cofor_rond, col_cofor_ac) & Cells(lig_cofor_rond, col_fnr_ac) & Cells(lig_cofor_rond, col_noa_ac))) Then
    '    ActiveSheet.Shapes(CStr(i)).Select
    '    Selection.ShapeRange.Fill.ForeColor.RGB = RGB(144, 238, 144)
    '    Else: End If
    '    Next i
    'Else: End If

    Sheets("BASE_temp").Visible = True
    Sheets("READ-ME OBLIGATORY!").Visible = True
    Sheets("subst_pack_form").Visible = True
    Sheets("SERIAL_PACKAGING_PACKMAN").Visible = True


Sheets("pickup_sheet").Select 'construction pickup sheet
If Not ActiveSheet.AutoFilter Is Nothing Then
Selection.AutoFilter
Else
End If

If Cells(500000, col_ref_p - 1).End(xlUp).Row = prem_lig_ref_p - 1 Then
    '
Else:
        'page 1
    Range(Cells(33, 3), Cells(67, lastColPUS)).Select
    Selection.ClearContents
    Selection.Interior.ColorIndex = xlColorIndexNone
        'page 2
    Range(Cells(101, 3), Cells(135, lastColPUS)).Select
    Selection.ClearContents
    Selection.Interior.ColorIndex = xlColorIndexNone
        'page 3
    Range(Cells(169, 3), Cells(203, lastColPUS)).Select
    Selection.ClearContents
    Selection.Interior.ColorIndex = xlColorIndexNone
    
        'page 4
    Range(Cells(237, 3), Cells(271, lastColPUS)).Select
    Selection.ClearContents
    Selection.Interior.ColorIndex = xlColorIndexNone
    
        'page 5
    Range(Cells(305, 3), Cells(339, lastColPUS)).Select
    Selection.ClearContents
    Selection.Interior.ColorIndex = xlColorIndexNone
    
        'page 6
    Range(Cells(373, 3), Cells(407, lastColPUS)).Select
    Selection.ClearContents
    Selection.Interior.ColorIndex = xlColorIndexNone
    
        'page 7
    Range(Cells(441, 3), Cells(475, lastColPUS)).Select
    Selection.ClearContents
    Selection.Interior.ColorIndex = xlColorIndexNone
    
        'page 8
    Range(Cells(509, 3), Cells(543, lastColPUS)).Select
    Selection.ClearContents
    Selection.Interior.ColorIndex = xlColorIndexNone
    
        'page 9
    Range(Cells(577, 3), Cells(611, lastColPUS)).Select
    Selection.ClearContents
    Selection.Interior.ColorIndex = xlColorIndexNone
    
        'page 10
    Range(Cells(645, 3), Cells(679, lastColPUS)).Select
    Selection.ClearContents
    Selection.Interior.ColorIndex = xlColorIndexNone
    
        'page 11
    Range(Cells(713, 3), Cells(747, lastColPUS)).Select
    Selection.ClearContents
    Selection.Interior.ColorIndex = xlColorIndexNone
    
        'page 12
    Range(Cells(781, 3), Cells(815, lastColPUS)).Select
    Selection.ClearContents
    Selection.Interior.ColorIndex = xlColorIndexNone
    
        'page 13
    Range(Cells(849, 3), Cells(883, lastColPUS)).Select
    Selection.ClearContents
    Selection.Interior.ColorIndex = xlColorIndexNone
    
        'page 14
    Range(Cells(917, 3), Cells(951, lastColPUS)).Select
    Selection.ClearContents
    Selection.Interior.ColorIndex = xlColorIndexNone
    
        'page 15
    Range(Cells(985, 3), Cells(1019, lastColPUS)).Select
    Selection.ClearContents
    Selection.Interior.ColorIndex = xlColorIndexNone
    
End If

Sheets("BASE").Rows(2).AutoFilter
' Au cas où, des tris non prévus aient été fait : tri de la feuille au nominal
Sheets("BASE").Select
If ActiveSheet.AutoFilter Is Nothing Then
Sheets("BASE").Rows(2).AutoFilter
Else
End If

'Range(Cells(prem_lig_ref_b - 1, col_ref_b), Cells(Cells(500000, 54).End(xlUp).Row, Cells(2, 1000).End(xlToLeft).Column)).Select
'Selection.Sort key1:=Cells(prem_lig_ref_b - 1, col_cofor_b), order1:=xlAscending, dataoption1:=xlSortNormal, key2:=Cells(prem_lig_ref_b - 1, 1), order2:=xlAscending, dataoption1:=xlSortNormal, key3:=Cells(prem_lig_ref_b - 1, 53), order3:=xlAscending, dataoption2:=xlSortNormal, key4:=Cells(prem_lig_ref_b - 1, 54), order4:=xlAscending, dataoption3:=xlSortNormal, Header:=xlYes


ActiveWorkbook.Worksheets("BASE").AutoFilter.Sort.SortFields.Add Key:=Range("F2"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
ActiveWorkbook.Worksheets("BASE").AutoFilter.Sort.SortFields.Add Key:=Cells(1, col_base_ech), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
'ActiveWorkbook.Worksheets("BASE").AutoFilter.Sort.SortFields.Clear
'ActiveWorkbook.Worksheets("BASE").AutoFilter.Sort.SortFields.Add Key:=Range("E2"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
ActiveWorkbook.Worksheets("BASE").AutoFilter.Sort.SortFields.Add Key:=Cells(1, col_base_mag), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
ActiveWorkbook.Worksheets("BASE").AutoFilter.Sort.SortFields.Add Key:=Range("A2"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal




With ActiveWorkbook.Worksheets("BASE").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With


' dans la colonne des quantités confirmées : filtre sur les vides et suppression (au cas où une cellule ne possède qu'un espace)
Sheets("BASE").Select
If Not ActiveSheet.AutoFilter Is Nothing Then
Selection.AutoFilter
Else
End If

Range(Cells(prem_lig_ref_b - 1, col_ref_b), Cells(Cells(500000, col_ref_b).End(xlUp).Row, Cells(2, 1).End(xlToRight).Column)).Select
Selection.AutoFilter
Range("A1").Select

Sheets("BASE").Cells.AutoFilter Field:=col_qte_conf_b, Criteria1:=""



Dim Maplagevisible As Range
Set Maplagevisible = Range("D2", Cells(Rows.count, "D").End(xlUp)).SpecialCells(xlCellTypeVisible)
Maplagevisible.Select
If Maplagevisible.Cells(1, 1) = "QUANTITE confirmée" Then
Range("D2", Cells(Rows.count, "D").End(xlUp)).SpecialCells(xlCellTypeVisible).Select
Selection.ClearContents
Cells(2, 4) = "QUANTITE confirmée"
Else: End If

''''''''''''''''''''''''' suite
Sheets("BASE").Select
If Not ActiveSheet.AutoFilter Is Nothing Then
Selection.AutoFilter
Else
End If
Rows("1:1").Hidden = True
'Selection.Delete Shift:=xlUp
Dim lastLine As Long


    ' to filtrowanie jeszcze trzeba poprawic tutaj
    ' this filtering is far away to be perfect!!!
    ' TEMP_REMARK_FOR_DEMO
    ' ------------------------------------------------
    
    If perimetre_analyse = "per_cofor" Then
        Sheets("BASE").Range("A2").AutoFilter Field:=col_cofor_b, Criteria1:="*" & cofor_a
        'Sheets("BASE").Range("A2").AutoFilter Field:=col_nom_fnr_b, Criteria1:=nom_fnr_a
        Sheets("BASE").Range("A2").AutoFilter Field:=col_nom_appro_b, Criteria1:=noa_a
    ElseIf perimetre_analyse = "per_reference" Then
        Sheets("BASE").Range("A2").AutoFilter Field:=col_cofor_b, Criteria1:="*" & cofor_a
    ' Sheets("BASE").Range("A2").AutoFilter Field:=col_nom_fnr_b, Criteria1:=nom_fnr_a
        Sheets("BASE").Range("A2").AutoFilter Field:=col_nom_appro_b, Criteria1:=noa_a
        Sheets("BASE").Range("A2").AutoFilter Field:=col_ref_b, Criteria1:=reference_a
    Else: End If

' Forcage d'une TMC à non en cas de transport par le fournisseur
' en effet :
' Pour un fournisseur qui n'est pas en DAP en flux série et qui livre les premières échéances à ses frais, il faut
' mettre " non " dans la colonne TMC. Si cela n'est pas fait, la PUS indiquera que toute la TMC est avec un transport à la charge du fnr
' cela arrive dans le cas où les lignes de la TMC sont sur un même cofor exp en colonne G de la base - cela peut arriver en cas de passage d'une
' échéance en transport fnr entre 2 mises à jour hebdo.

Dim Maplage As Range
Set Maplage = Sheets("BASE").UsedRange.SpecialCells(xlCellTypeVisible)
Dim Ligne As Range

For Each Ligne In Maplage.Rows
    If Ligne.Cells(col_cofor_exp_b).Value = "transport by supplier" And Ligne.Cells(col_tmc_b).Value <> "non" Then
        Ligne.Cells(col_tmc_b).Value = "non"
    Else: End If
Next


    lig_ref_p = prem_lig_ref_p

    Sheets("pickup_sheet").Select


    ' only 3 static pages with only clearing on the top one
    ' on 2nd and 3d you have formulas, which just copy&paste info
    
    'Nettoyage
        'SUPPLIER
        Cells(7, 6) = ""
        Cells(8, 6) = ""
        Cells(9, 6) = ""
        Cells(10, 6) = ""
        Cells(11, 6) = ""
        Cells(12, 6) = ""
        Cells(13, 6) = ""
        'PSA PLANT Contacts
        
        ' PROJECT 16,6 -> F16
        Cells(16, 6) = ""
        Cells(17, 6) = ""
        Cells(18, 6) = ""
        Cells(19, 6) = ""
        Cells(20, 6) = ""
        Cells(21, 6) = ""
        Cells(22, 6) = ""
        'CARRIER
        Cells(23, 6) = ""
        Cells(24, 6) = ""
        Cells(25, 6) = ""
        Cells(26, 6) = ""
        Cells(27, 6) = ""
        
        ActiveSheet.Rows("1:1100").Hidden = False
        
        Dim compteur As Long
        Dim writingLine As Long
        compteur = 0
        writingLine = 1
        
        
        

        
        
        ' copie des valeurs dans la pick up sheet
        
        ' MAIN PART NUMBERS LOOP BEGINING
        ' ===========================================================
        ' ===========================================================
        Dim count As Long
        count = 1
        For Each Ligne In Maplage.Rows
            Application.StatusBar = "ref count : " & count
            count = count + 1
            
            If Ligne.Cells(col_ref_b).Value = "REFERENCE" Then
            '
            Else
            
                Sheets("pickup_sheet").Select
            
            
            '''''''''''''''''''''''''''''''''''''''
            '''''''''''''''''''''''''''''''''''''''
            '''''''''''''''''''''''''''''''''''''''
            '''''''''''''''''''''''''''''''''''''''
            '''''''''''''''''''''''''''''''''''''''
            If Ligne.Cells(1).Value = "" Then
                Exit For
            End If
                    'regroupement + refus regroupement --> "non" dans la colonne P de la base
                If Ligne.Cells(16).Value <> "non" And Cells(writingLine, 3).Value Like Ligne.Cells(1).Value And Cells(writingLine, 5).Value Like Ligne.Cells(col_base_ech).Value And Cells(writingLine, 8).Value Like Ligne.Cells(col_base_mag).Value Then

                    If IsNumeric(Ligne.Cells(col_qte_theo_b).Value) Then
      
                        Cells(writingLine, 12).Value = Cells(writingLine, 12).Value + Ligne.Cells(col_qte_theo_b).Value
                    End If
            
                Else
                    
                    
                    
                    compteur = compteur + 1
                    writingLine = findRow(ThisWorkbook.Sheets("pickup_sheet"), 2, CLng(compteur))
            

                    Dim coforExp As String
                    If Ligne.Cells(38).Value <> "" Then
                        coforExp = Ligne.Cells(38).Value
                    End If
            
            
                
                    'PARTS
                    Cells(writingLine, 3) = Ligne.Cells(col_ref_b).Value
                    Cells(writingLine, 4) = Ligne.Cells(col_desi_b).Value
                    
                    If Ligne.Cells(16).Value = "non" Then
                        Cells(writingLine, 5) = Ligne.Cells(5).Value 'si on refuse le regroupement, on se base sur la date initiale
                    Else
                        Cells(writingLine, 5) = Ligne.Cells(col_base_ech).Value 'sinon sur la date regroupé
                    End If
                    
                    
                    Dim testVal As String
                    testVal = Ligne.Cells(43).Value
                    
                    If InStr(1, testVal, "transport by supplier", vbTextCompare) <> 1 And InStr(1, testVal, "DAP", vbTextCompare) <> 1 Then
                    'If Ligne.Cells(43).Value <> "transport by supplier (serial logistic in DAP) " And Ligne.Cells(43).Value <> "transport by supplier (serial logistic in DAP)" And Ligne.Cells(43).Value <> "DAP" Then
                        Cells(writingLine, 6) = Ligne.Cells(44).Value 'date pickup
                        Else
                        Cells(23, 6) = "transport by supplier (serial logistic in DAP)"
                        Cells(24, 6) = "transport by supplier (serial logistic in DAP)"
                        Cells(25, 6) = "transport by supplier (serial logistic in DAP)"
                        Cells(26, 6) = "transport by supplier (serial logistic in DAP)"
                    End If
                    
                        ' NOK this one
                        'Cells(writingLine, 7) = Ligne.Cells(16).Value
                        'Cells(writingLine, 8) = Ligne.Cells(col_base_mag).Value 'mag
                    
                        'nom du projet - on écrase temporairement le mag par le domaine, pour éviter bug de la formule
                        ' Dim mag As Integer in xp mag is number
                        ' in xF I want to use mag as string
                        Dim mag As String
                        mag = Ligne.Cells(col_base_mag).Value
                        Dim foundCol As Integer
                        foundCol = Sheets("Macro & projects infos").Rows(6).Find(What:=mag, LookAt:=xlWhole).Column
                        
                        Cells(writingLine, 9).Value = Sheets("Macro & projects infos").Cells(5, foundCol).Value
                        
                        Cells(writingLine, 8).Value = CStr(mag)
                        
                        
                        ' contract and del prog!
                        ' ------------------------------------------------------
                        'Cells(writingLine, 10) = Ligne.Cells(9).Value
                        'Cells(writingLine, 11) = Ligne.Cells(8).Value
                        ' ------------------------------------------------------
                    
                    ' QTY calc - minus confirmed
                    ' NEW NEW!   -> TEMP_REMARK_FOR_DEMO
                    'If CLng(Ligne.Cells(col_qte_theo_b + 1).Value) > 0 Then
                    '    Cells(writingLine, 12).Value = Ligne.Cells(col_qte_theo_b).Value - Ligne.Cells(col_qte_theo_b + 1).Value
                    'Else
                    '    Cells(writingLine, 12) = Ligne.Cells(col_qte_theo_b).Value
                    'End If
                    
                    ' always! -> so if it is zero -> then no problem -> just theo will stay as it is!
                    Cells(writingLine, 12).Value = Ligne.Cells(col_qte_theo_b).Value - Ligne.Cells(col_qte_theo_b + 1).Value
                    
                
                
                    Cells(writingLine, 13).Value = Ligne.Cells(24).Value
                    If Ligne.Cells(27).Value <> "" And Ligne.Cells(28).Value <> "" And Ligne.Cells(29).Value <> "" Then 'dimensions
                        Cells(writingLine, 14) = Ligne.Cells(27).Value & "x" & Ligne.Cells(28).Value & "x" & Ligne.Cells(29).Value
                    Else: End If
                
                    Cells(writingLine, 15) = Ligne.Cells(26).Value
                    'Cells(writingLine, 16) -->
                    Cells(writingLine, 18) = Ligne.Cells(25).Value
        
                
                    'SUPPLIER
                    If Cells(7, 6) = "" Then
                        Cells(7, 6) = Ligne.Cells(col_pickup_address_b).Value
                    End If
                
                    If Cells(8, 6) = "" Then
                        Cells(8, 6) = Ligne.Cells(col_nom_fnr_b).Value
                    End If
                    
                    If Cells(9, 6) = "" Then
                        Cells(9, 6) = Ligne.Cells(col_cofor_vend_b).Value
                    End If
                    
                    If Cells(10, 6) = "" Then
                        Cells(10, 6) = Ligne.Cells(col_cofor_exp_b).Value
                    End If
                    
                    If Cells(11, 6) = "" Then
                        Cells(11, 6) = Ligne.Cells(col_psa_supply_1_b).Value
                    End If
                    
                    If Cells(12, 6) = "" Then
                        Cells(12, 6) = Ligne.Cells(col_psa_supply_3_b).Value
                    End If
                    
                    If Cells(13, 6) = "" Then
                        Cells(13, 6) = Ligne.Cells(col_psa_supply_2_b).Value
                    End If
        
                
                    'PSA PLANT Contacts
                    If Cells(19, 6) = "" Then
                        Cells(19, 6) = Ligne.Cells(col_psa_contact_1_b).Value
                    End If
                    If Cells(20, 6) = "" Then
                        Cells(20, 6) = Ligne.Cells(col_psa_contact_3_b).Value
                    End If
                    If Cells(21, 6) = "" Then
                        Cells(21, 6) = Ligne.Cells(col_psa_contact_2_b).Value
                    End If
                    
                    'CARRIER
                    If Cells(23, 6) = "" Then
                        Cells(23, 6) = Ligne.Cells(col_gefco_4_b).Value
                    End If
                    If Cells(24, 6) = "" Then
                        Cells(24, 6) = Ligne.Cells(col_gefco_5_b).Value
                    End If
                    If Cells(25, 6) = "" Then
                        Cells(25, 6) = Ligne.Cells(col_gefco_6_b).Value
                    End If
                    
                    If Cells(26, 6) = "" Then
                        Cells(26, 6) = Ligne.Cells(col_gefco_3_b).Value
                    End If
                
        
                
                End If
            End If
            
        Next
        
        
        Dim qty As Double
        Dim condi As Double
        Dim nbContainer As Long

    Application.StatusBar = "Container qty calculations.."
    lastLine = Cells(writingLine, 2).Value
    For i = 1 To lastLine
        writingLine = findRow(ThisWorkbook.Sheets("pickup_sheet"), 2, CLng(i))
        qty = Cells(writingLine, 12).Value
        condi = Cells(writingLine, 18).Value
        If condi > 0 Then
            On Error Resume Next
            nbContainer = Int(qty / condi)
            If nbContainer > 0 And Cells(writingLine, 15).Value > 0 Then
                Cells(writingLine, 16).Value = (nbContainer + 1) * Cells(writingLine, 15).Value 'poids total
            End If
            On Error GoTo 0
            Cells(writingLine, 17).Value = nbContainer + 1 'nb of pack
        End If
    Next i
    
    
    Dim test As String
    Dim testResult As Long
    Application.StatusBar = "Preparing MAF infos.."
    Dim colMAF As Long
    colMAF = Sheets("Macro & projects infos").Rows(4).Find(What:="MAF infos", LookAt:=xlWhole).Column
    
    Sheets("pickup_sheet").Cells(14, 6).Value = ""
    Sheets("pickup_sheet").Cells(15, 6).Value = ""

    For i = 6 To Sheets("Macro & projects infos").Cells(1000, colMAF).End(xlUp).Row
    test = Sheets("Macro & projects infos").Cells(i, colMAF).Value
    testResult = StrComp(coforExp, test, vbTextCompare)
        If testResult = 0 Then
            Sheets("pickup_sheet").Cells(14, 6).Value = Sheets("Macro & projects infos").Cells(i, colMAF + 1).Value
            Sheets("pickup_sheet").Cells(15, 6).Value = Sheets("Macro & projects infos").Cells(i, colMAF + 2).Value
            Exit For
        End If
    Next i

  
    
    
    
    
    
    ' récupération des paramètres
        'Dim Prj As String
        'Prj = Sheets("pickup_sheet").Cells(33, 9).Value
        Dim colPrj As Long
        colPrj = 2 'toutes les infos sont les mêmes dans les colonnes ' nie prawda!
        
        ' colPrj should be diff in case of dirrent mag/domain
        ' by default colPrj = 2
        
        ' look for the domain in "Macro & projects infos"
        
        Dim r As Range
        Set r = ThisWorkbook.Sheets("Macro & projects infos").Range("B6")
        
        
        ' we will just look into very frist cell with magasin value in pickup_sheet
        Dim magRng1 As Range
        Set magRng1 = ThisWorkbook.Sheets("pickup_sheet").Range("H33")
        
        Do
            If Trim(r.Value) = Trim(magRng1.Value) Then
                    colPrj = CLng(r.Column)
                Exit Do
            End If
            
            Set r = r.Offset(0, 1)
        Loop While r.Value <> ""
        
        
        'PSA PLANT
        Cells(16, 6) = ThisWorkbook.Sheets("Macro & projects infos").Cells(9, colPrj).Value
        
        Cells(2, 8) = "PROJECT " & Sheets("Macro & projects infos").Cells(9, colPrj).Value 'entete
        Cells(2, 13) = Date 'entete date
        
        Cells(17, 6) = Sheets("Macro & projects infos").Cells(10, colPrj).Value
        Cells(18, 6) = Sheets("Macro & projects infos").Cells(11, colPrj).Value

        ' CARRIER
        Cells(27, 6) = Sheets("Macro & projects infos").Cells(12, colPrj).Value




        'ajuster hauteur cellule adresse
        Range("C17:E17").MergeCells = False
     
     Dim hauteur As Integer
     
        With Range("F17:R17")
            .MergeCells = False
            .Rows.EntireRow.AutoFit
            DoEvents
            hauteur = Int(.RowHeight / 5 * 4)
            .WrapText = True
            .MergeCells = True
            .RowHeight = hauteur
            .VerticalAlignment = xlVAlignCenter
        End With
            
        Range("C17:E17").MergeCells = True
  
    
    'OPTIONAL - IF MAF IS USED IN SERIAL FLOW - CONTACT FOR MAF/ IF MAF IS NOT USED FILL CELL WITH "NOT APPLICABLE"
    
    
    Application.StatusBar = "Printing preparation.."
    
        
        ' MAIN PART NUMBER LOOP END
        ' ===========================================================
        ' ===========================================================

    'configuration impressions
    If compteur <= 35 Then
        ActiveSheet.PageSetup.PrintArea = "$A$1:$P$68"
        ActiveSheet.Rows("69:1019").Hidden = True
        
        ActiveSheet.ResetAllPageBreaks
    
    ElseIf compteur <= 75 Then
        ActiveSheet.PageSetup.PrintArea = "$A$1:$P$136"
        
        ActiveSheet.ResetAllPageBreaks
        
        On Error Resume Next
        ActiveSheet.HPageBreaks.Add Range("A69")
        ActiveSheet.Rows("137:1019").Hidden = True
    
    ElseIf compteur <= 105 Then
        ActiveSheet.PageSetup.PrintArea = "$A$1:$P$204"
        ActiveSheet.Rows("205:1019").Hidden = True
        ActiveSheet.ResetAllPageBreaks
        
        
        On Error Resume Next
        ActiveSheet.HPageBreaks.Add Range("A69")
        On Error Resume Next
        ActiveSheet.HPageBreaks.Add Range("A137")
        
        
    ElseIf compteur <= 140 Then
    
        ActiveSheet.PageSetup.PrintArea = "$A$1:$P$271"
        ActiveSheet.Rows("273:1019").Hidden = True
        ActiveSheet.ResetAllPageBreaks
        
        
        On Error Resume Next
        ActiveSheet.HPageBreaks.Add Range("A69")
        On Error Resume Next
        ActiveSheet.HPageBreaks.Add Range("A137")
        On Error Resume Next
        ActiveSheet.HPageBreaks.Add Range("A205")
        
        
    ElseIf compteur <= 175 Then
    
        ActiveSheet.PageSetup.PrintArea = "$A$1:$P$339"
        ActiveSheet.Rows("341:1019").Hidden = True
        
        ActiveSheet.ResetAllPageBreaks
        
        
        On Error Resume Next
        ActiveSheet.HPageBreaks.Add Range("A69")
        On Error Resume Next
        ActiveSheet.HPageBreaks.Add Range("A137")
        On Error Resume Next
        ActiveSheet.HPageBreaks.Add Range("A205")
        On Error Resume Next
        ActiveSheet.HPageBreaks.Add Range("A273")
        
        
        

        
    ElseIf compteur <= 210 Then
    
        ActiveSheet.PageSetup.PrintArea = "$A$1:$P$407"
        ActiveSheet.Rows("409:1019").Hidden = True
        ActiveSheet.ResetAllPageBreaks
        
        
        On Error Resume Next
        ActiveSheet.HPageBreaks.Add Range("A69")
        On Error Resume Next
        ActiveSheet.HPageBreaks.Add Range("A137")
        On Error Resume Next
        ActiveSheet.HPageBreaks.Add Range("A205")
        On Error Resume Next
        ActiveSheet.HPageBreaks.Add Range("A273")
        On Error Resume Next
        ActiveSheet.HPageBreaks.Add Range("A341")
    
    ElseIf compteur <= 245 Then
    
        ActiveSheet.PageSetup.PrintArea = "$A$1:$P$475"
        ActiveSheet.Rows("475:1019").Hidden = True
        ActiveSheet.ResetAllPageBreaks
        
        On Error Resume Next
        ActiveSheet.HPageBreaks.Add Range("A69")
        On Error Resume Next
        ActiveSheet.HPageBreaks.Add Range("A137")
        On Error Resume Next
        ActiveSheet.HPageBreaks.Add Range("A205")
        On Error Resume Next
        ActiveSheet.HPageBreaks.Add Range("A273")
        On Error Resume Next
        ActiveSheet.HPageBreaks.Add Range("A341")
        On Error Resume Next
        ActiveSheet.HPageBreaks.Add Range("A409")
        
    ElseIf compteur <= 280 Then
    
        ActiveSheet.PageSetup.PrintArea = "$A$1:$P$543"
        ActiveSheet.Rows("545:1019").Hidden = True
        ActiveSheet.ResetAllPageBreaks
        
        On Error Resume Next
        ActiveSheet.HPageBreaks.Add Range("A69")
        On Error Resume Next
        ActiveSheet.HPageBreaks.Add Range("A137")
        On Error Resume Next
        ActiveSheet.HPageBreaks.Add Range("A205")
        On Error Resume Next
        ActiveSheet.HPageBreaks.Add Range("A273")
        On Error Resume Next
        ActiveSheet.HPageBreaks.Add Range("A341")
        On Error Resume Next
        ActiveSheet.HPageBreaks.Add Range("A409")
        On Error Resume Next
        ActiveSheet.HPageBreaks.Add Range("A477")
    
    ElseIf compteur <= 315 Then
    
        ActiveSheet.PageSetup.PrintArea = "$A$1:$P$611"
        ActiveSheet.Rows("613:1019").Hidden = True
        ActiveSheet.ResetAllPageBreaks
        
        On Error Resume Next
        ActiveSheet.HPageBreaks.Add Range("A69")
        On Error Resume Next
        ActiveSheet.HPageBreaks.Add Range("A137")
        On Error Resume Next
        ActiveSheet.HPageBreaks.Add Range("A205")
        On Error Resume Next
        ActiveSheet.HPageBreaks.Add Range("A273")
        On Error Resume Next
        ActiveSheet.HPageBreaks.Add Range("A341")
        On Error Resume Next
        ActiveSheet.HPageBreaks.Add Range("A409")
        On Error Resume Next
        ActiveSheet.HPageBreaks.Add Range("A477")
        On Error Resume Next
        ActiveSheet.HPageBreaks.Add Range("A545")
    
    ElseIf compteur <= 350 Then
    
        ActiveSheet.PageSetup.PrintArea = "$A$1:$P$679"
        ActiveSheet.Rows("681:1019").Hidden = True
        ActiveSheet.ResetAllPageBreaks
        
        On Error Resume Next
        ActiveSheet.HPageBreaks.Add Range("A69")
        On Error Resume Next
        ActiveSheet.HPageBreaks.Add Range("A137")
        On Error Resume Next
        ActiveSheet.HPageBreaks.Add Range("A205")
        On Error Resume Next
        ActiveSheet.HPageBreaks.Add Range("A273")
        On Error Resume Next
        ActiveSheet.HPageBreaks.Add Range("A341")
        On Error Resume Next
        ActiveSheet.HPageBreaks.Add Range("A409")
        On Error Resume Next
        ActiveSheet.HPageBreaks.Add Range("A477")
        On Error Resume Next
        ActiveSheet.HPageBreaks.Add Range("A545")
        On Error Resume Next
        ActiveSheet.HPageBreaks.Add Range("A613")
        
     ElseIf compteur <= 385 Then
    
        ActiveSheet.PageSetup.PrintArea = "$A$1:$P$747"
        ActiveSheet.Rows("749:1019").Hidden = True
        ActiveSheet.ResetAllPageBreaks
        
        On Error Resume Next
        ActiveSheet.HPageBreaks.Add Range("A69")
        On Error Resume Next
        ActiveSheet.HPageBreaks.Add Range("A137")
        On Error Resume Next
        ActiveSheet.HPageBreaks.Add Range("A205")
        On Error Resume Next
        ActiveSheet.HPageBreaks.Add Range("A273")
        On Error Resume Next
        ActiveSheet.HPageBreaks.Add Range("A341")
        On Error Resume Next
        ActiveSheet.HPageBreaks.Add Range("A409")
        On Error Resume Next
        ActiveSheet.HPageBreaks.Add Range("A477")
        On Error Resume Next
        ActiveSheet.HPageBreaks.Add Range("A545")
        On Error Resume Next
        ActiveSheet.HPageBreaks.Add Range("A613")
        On Error Resume Next
        ActiveSheet.HPageBreaks.Add Range("A681")
    
    ElseIf compteur <= 420 Then
    
        ActiveSheet.PageSetup.PrintArea = "$A$1:$P$815"
        ActiveSheet.Rows("817:1019").Hidden = True
        ActiveSheet.ResetAllPageBreaks
        
        On Error Resume Next
        ActiveSheet.HPageBreaks.Add Range("A69")
        On Error Resume Next
        ActiveSheet.HPageBreaks.Add Range("A137")
        On Error Resume Next
        ActiveSheet.HPageBreaks.Add Range("A205")
        On Error Resume Next
        ActiveSheet.HPageBreaks.Add Range("A273")
        On Error Resume Next
        ActiveSheet.HPageBreaks.Add Range("A341")
        On Error Resume Next
        ActiveSheet.HPageBreaks.Add Range("A409")
        On Error Resume Next
        ActiveSheet.HPageBreaks.Add Range("A477")
        On Error Resume Next
        ActiveSheet.HPageBreaks.Add Range("A545")
        On Error Resume Next
        ActiveSheet.HPageBreaks.Add Range("A613")
        On Error Resume Next
        ActiveSheet.HPageBreaks.Add Range("A681")
        On Error Resume Next
        ActiveSheet.HPageBreaks.Add Range("A749")
        
    ElseIf compteur <= 420 Then
    
        ActiveSheet.PageSetup.PrintArea = "$A$1:$P$883"
        ActiveSheet.Rows("885:1019").Hidden = True
        ActiveSheet.ResetAllPageBreaks
        
        On Error Resume Next
        ActiveSheet.HPageBreaks.Add Range("A69")
        On Error Resume Next
        ActiveSheet.HPageBreaks.Add Range("A137")
        On Error Resume Next
        ActiveSheet.HPageBreaks.Add Range("A205")
        On Error Resume Next
        ActiveSheet.HPageBreaks.Add Range("A273")
        On Error Resume Next
        ActiveSheet.HPageBreaks.Add Range("A341")
        On Error Resume Next
        ActiveSheet.HPageBreaks.Add Range("A409")
        On Error Resume Next
        ActiveSheet.HPageBreaks.Add Range("A477")
        On Error Resume Next
        ActiveSheet.HPageBreaks.Add Range("A545")
        On Error Resume Next
        ActiveSheet.HPageBreaks.Add Range("A613")
        On Error Resume Next
        ActiveSheet.HPageBreaks.Add Range("A681")
        On Error Resume Next
        ActiveSheet.HPageBreaks.Add Range("A749")
        On Error Resume Next
        ActiveSheet.HPageBreaks.Add Range("A817")
    
    ElseIf compteur <= 455 Then
    
        ActiveSheet.PageSetup.PrintArea = "$A$1:$P$951"
        ActiveSheet.Rows("953:1019").Hidden = True
        ActiveSheet.ResetAllPageBreaks
        
        On Error Resume Next
        ActiveSheet.HPageBreaks.Add Range("A69")
        On Error Resume Next
        ActiveSheet.HPageBreaks.Add Range("A137")
        On Error Resume Next
        ActiveSheet.HPageBreaks.Add Range("A205")
        On Error Resume Next
        ActiveSheet.HPageBreaks.Add Range("A273")
        On Error Resume Next
        ActiveSheet.HPageBreaks.Add Range("A341")
        On Error Resume Next
        ActiveSheet.HPageBreaks.Add Range("A409")
        On Error Resume Next
        ActiveSheet.HPageBreaks.Add Range("A477")
        On Error Resume Next
        ActiveSheet.HPageBreaks.Add Range("A545")
        On Error Resume Next
        ActiveSheet.HPageBreaks.Add Range("A613")
        On Error Resume Next
        ActiveSheet.HPageBreaks.Add Range("A681")
        On Error Resume Next
        ActiveSheet.HPageBreaks.Add Range("A749")
        On Error Resume Next
        ActiveSheet.HPageBreaks.Add Range("A817")
        On Error Resume Next
        ActiveSheet.HPageBreaks.Add Range("A885")
        
    
    ElseIf compteur <= 490 Then
    
        ActiveSheet.PageSetup.PrintArea = "$A$1:$P$1019"
        ActiveSheet.ResetAllPageBreaks
        
        On Error Resume Next
        ActiveSheet.HPageBreaks.Add Range("A69")
        On Error Resume Next
        ActiveSheet.HPageBreaks.Add Range("A137")
        On Error Resume Next
        ActiveSheet.HPageBreaks.Add Range("A205")
        On Error Resume Next
        ActiveSheet.HPageBreaks.Add Range("A273")
        On Error Resume Next
        ActiveSheet.HPageBreaks.Add Range("A341")
        On Error Resume Next
        ActiveSheet.HPageBreaks.Add Range("A409")
        On Error Resume Next
        ActiveSheet.HPageBreaks.Add Range("A477")
        On Error Resume Next
        ActiveSheet.HPageBreaks.Add Range("A545")
        On Error Resume Next
        ActiveSheet.HPageBreaks.Add Range("A613")
        On Error Resume Next
        ActiveSheet.HPageBreaks.Add Range("A681")
        On Error Resume Next
        ActiveSheet.HPageBreaks.Add Range("A749")
        On Error Resume Next
        ActiveSheet.HPageBreaks.Add Range("A817")
        On Error Resume Next
        ActiveSheet.HPageBreaks.Add Range("A885")
        On Error Resume Next
        ActiveSheet.HPageBreaks.Add Range("A953")
    Else

        Stop 'more than 525 parts ! does not fit in the file
    End If




ActiveSheet.AutoFilterMode = False

Sheets("BASE").Select
If Not ActiveSheet.AutoFilter Is Nothing Then
Selection.AutoFilter
Else
End If

'Sheets("BASE_temp").Select
'Rows("1:1").Select
'Selection.Copy
'Sheets("BASE").Select
'Selection.Insert Shift:=xlDown

' activer le format de sortie souhaité : PDF ou excel
'Call PDFActiveSheet
Call excelsheet

Sheets("READ-ME OBLIGATORY!").Visible = False
Sheets("subst_pack_form").Visible = False
Sheets("SERIAL_PACKAGING_PACKMAN").Visible = False
Sheets("BASE_temp").Visible = False


Sheets("pickup_sheet").Select
Cells(1, 1).Select

Sheets("BASE").Select
If Not ActiveSheet.AutoFilter Is Nothing Then
Selection.AutoFilter
Else
End If


Sheets("PUS creation").Select
Cells(1, 1).Select

testPUS = 0

Application.StatusBar = ""

exceptionForSecondPage:

    Debug.Print "problem!"

End Sub



Private Function simplifyTMC(org As String, newOne As String) As String
    
    ' get common
    Dim x As Variant
    x = 1
    Dim commonPattern As String
    commonPattern = ""
    
    For x = 1 To Len(org)
        If Left(org, x) = Left(newOne, x) Then
            commonPattern = Left(org, x)
        Else
            Exit For
        End If
    Next x
    
    simplifyTMC = commonPattern
End Function


Private Function checkPrevLine(wLine As Long, ByRef prev_lineRngRef As Range, ByRef line As Range) As Long
    
    checkPrevLine = -1
    
    If wLine < PUS.G_FIRST_NUMBER_ROW Then
        'nothing to do
    Else
    
    
        ' same part number - first IF check
        If prev_lineRngRef.Value = line.Cells(col_ref_b).Value Then
        
            ' same part name - second IF check
            If prev_lineRngRef.Offset(0, 1).Value = line.Cells(col_desi_b).Value Then
            
                ' 0, 4 by hard tmc in pus sheet output
                '
                checkPrevLine = Int(checkTMC(wLine, prev_lineRngRef.Offset(0, 4).Value, line.Cells(16).Value))
            End If
            
        End If
        
        ' there is still no match
        If checkPrevLine = -1 Then
            ' simple recurence as long as wLine is bigger then first number row in pus sheet - G_FIRST_NUMBER_ROW
            checkPrevLine = checkPrevLine(wLine - 1, prev_lineRngRef.Offset(-1, 0), line)
        End If
    End If
End Function

Private Function checkTMC(wLine As Long, tmc1 As String, tmc2 As String) As Long

    checkTMC = -1
    
    ' tmc1 - prev tmc
    ' tmc2 - actual tmc from BASE worksheet
    
    
    If UCase(tmc2) = "NON" Then
        checkTMC = -1
    Else
    
        Dim wo_ending_tmc1 As String, wo_ending_tmc2 As String
        
        If UCase(tmc1) = UCase(tmc2) Then
            ' most simple scenario
            checkTMC = wLine
        Else
        
            If UCase(tmc1) Like "TMC*" And UCase(tmc2) Like "TMC*" Then

                wo_ending_tmc1 = Left(tmc1, Len(tmc1) - 0)
                wo_ending_tmc2 = Left(tmc2, Len(tmc2) - 0)
                
                If wo_ending_tmc1 = wo_ending_tmc2 Then
                    checkTMC = wLine
                Else
                    checkTMC = -1
                End If
                
            End If
        End If
    
    End If
End Function


Private Function findRow(sh1 As Worksheet, col As Long, m_compteur As Long) As Long

    ' if -1, then calc issue - should never happen!
    findRow = -1
    
    Dim r As Range
    Set r = sh1.Cells(1, col)
    Do
        Set r = r.Offset(1, 0)
        
        If r.Row > 994 Then
        
            MsgBox "No big pn list - can only go up to 500 pn - please contact dev!", vbCritical

            End
        
        End If
    Loop Until r.Value = m_compteur
    
    
    findRow = r.Row
End Function


Private Sub clearBreaks()
 
    Do While ActiveSheet.HPageBreaks.count > 1
        ActiveSheet.HPageBreaks(ActiveSheet.HPageBreaks.count).Delete
    Loop
End Sub





