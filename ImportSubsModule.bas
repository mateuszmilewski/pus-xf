Attribute VB_Name = "ImportSubsModule"
Option Explicit

'Module ajouté par GZA (Alten)
Sub Step1_import_sigapp_sechel()
    
    '1.Désigner les onglets de travail
    '1.1.Ouverture explorateur et selection du fichier d'extract
    Dim Fichier As FileDialog
    Set Fichier = Application.FileDialog(msoFileDialogFilePicker)
    
    Fichier.Title = "Sélection du fichier Sigapp-sechel"
    Fichier.Show
    
    If Fichier.SelectedItems.count = 0 Then
        MsgBox "Aucun fichier sélectionné"
        Exit Sub 'Stopper la macro
    End If
    
    '1.2.Déclaration des variables Excel
    Dim MonAppliExcel As Excel.Application 'Application Excel générale
    Dim fichier_pus As Excel.Workbook 'Fichiers
    Dim fichier_extract As Excel.Workbook
    Dim onglet_sechel As Worksheet 'Onglets
    Dim onglet_lisezmoi As Worksheet
    Dim onglet_extract_ion As Worksheet
    Dim onglet_extract As Worksheet
    '1.3.Fichier et onglets de la pus
    Set fichier_pus = ThisWorkbook
    Set onglet_sechel = fichier_pus.Worksheets("sechel")
    onglet_sechel.Visible = xlSheetVisible
    Set onglet_lisezmoi = fichier_pus.Worksheets("NOA")
    onglet_lisezmoi.Visible = xlSheetVisible
    '1.4.Fichier et onglet de l'extraction
    Set MonAppliExcel = New Excel.Application 'ouverture d'un second excel en parallèle
    MonAppliExcel.Visible = True ' temp
    Set fichier_extract = MonAppliExcel.Workbooks.Add(Fichier.SelectedItems(1))
    Set onglet_extract = fichier_extract.Worksheets(1)
    Application.DisplayAlerts = False
    ' Application.ScreenUpdating = False
    
    onglet_extract.Select
    If Not onglet_extract.AutoFilter Is Nothing Then
    onglet_extract.Cells.AutoFilter
    Else
    End If
    
    
    ' ---------------------------------------------------------------------------------------------------------
    '2.Préparation et vérifications des données
    Dim colNOA_extract, colArticle_extract, colNOA_id_extract, colDesignation_extract, colFourn_extract, colNomFourn_extract, colRU_extract, colDateEch_extract, colQteEch_extract, colQteLiv_extract, colMag_extract, colDocAchat_extract, colGAc_extract, colDomaine_extract As Variant
    Dim colNOA_sechel, colArticle_sechel, colNOA_id_sechel, colDesignation_sechel, colFourn_sechel, colNomFourn_sechel, colRU_sechel, colDateEch_sechel, colQteEch_sechel, colQteLiv_sechel, colMag_sechel, colDocAchat_sechel, colGAc_sechel, colDomaine_sechel As Variant
    Dim lastRow_extract As Long
    Dim lastCol_extract As Long
    Dim lastRow_sechel As Long
    Dim lastCol_sechel As Long
    Dim i As Long
    
    ' ---------------------------------------------------------------------------------------------------------
    '2.1.Tester la validité des colonnes de l'extraction et les repérer
    On Error GoTo erreurFormat
    colNOA_extract = onglet_extract.Rows(1).Find(What:="GAc Nom NOA", LookAt:=xlWhole).Column
    colNOA_id_extract = onglet_extract.Rows(1).Find(What:="NOA", LookAt:=xlWhole).Column
    colArticle_extract = onglet_extract.Rows(1).Find(What:="Article", LookAt:=xlWhole).Column
    colDesignation_extract = onglet_extract.Rows(1).Find(What:="Désignation", LookAt:=xlWhole).Column
    colFourn_extract = onglet_extract.Rows(1).Find(What:="Fourn.", LookAt:=xlWhole).Column
    colNomFourn_extract = onglet_extract.Rows(1).Find(What:="Nom fournisseur", LookAt:=xlWhole).Column
    colRU_extract = onglet_extract.Rows(1).Find(What:="RU", LookAt:=xlWhole).Column
    colDateEch_extract = onglet_extract.Rows(1).Find(What:="Date écheance", LookAt:=xlWhole).Column
    colQteEch_extract = onglet_extract.Rows(1).Find(What:="Qté échéancée", LookAt:=xlWhole).Column
    colQteLiv_extract = onglet_extract.Rows(1).Find(What:="Quantité livrée", LookAt:=xlWhole).Column
    colMag_extract = onglet_extract.Rows(1).Find(What:="Mag.", LookAt:=xlWhole).Column
    colDocAchat_extract = onglet_extract.Rows(1).Find(What:="Doc achat", LookAt:=xlWhole).Column
    colGAc_extract = onglet_extract.Rows(1).Find(What:="GAc", LookAt:=xlWhole).Column
    lastRow_extract = onglet_extract.Cells(1000000, colNOA_extract).End(xlUp).Row
    lastCol_extract = onglet_extract.Cells(1, 100).End(xlToLeft).Column
    '2.3.Repérer les colonnes de l'onglet sechel
    colNOA_sechel = onglet_sechel.Rows(1).Find(What:="GAc Nom NOA", LookAt:=xlWhole).Column
    colNOA_id_sechel = onglet_sechel.Rows(1).Find(What:="NOA", LookAt:=xlWhole).Column
    colArticle_sechel = onglet_sechel.Rows(1).Find(What:="Article", LookAt:=xlWhole).Column
    colDesignation_sechel = onglet_sechel.Rows(1).Find(What:="Désignation", LookAt:=xlWhole).Column
    colFourn_sechel = onglet_sechel.Rows(1).Find(What:="Fourn.", LookAt:=xlWhole).Column
    colNomFourn_sechel = onglet_sechel.Rows(1).Find(What:="Nom fournisseur", LookAt:=xlWhole).Column
    colRU_sechel = onglet_sechel.Rows(1).Find(What:="RU", LookAt:=xlWhole).Column
    colDateEch_sechel = onglet_sechel.Rows(1).Find(What:="Date écheance", LookAt:=xlWhole).Column
    colQteEch_sechel = onglet_sechel.Rows(1).Find(What:="Qté échéancée", LookAt:=xlWhole).Column
    colQteLiv_sechel = onglet_sechel.Rows(1).Find(What:="Quantité livrée", LookAt:=xlWhole).Column
    colMag_sechel = onglet_sechel.Rows(1).Find(What:="Mag.", LookAt:=xlWhole).Column
    colDocAchat_sechel = onglet_sechel.Rows(1).Find(What:="Doc achat", LookAt:=xlWhole).Column
    colGAc_sechel = onglet_sechel.Rows(1).Find(What:="GAc", LookAt:=xlWhole).Column
    lastRow_sechel = onglet_sechel.Cells(1000000, colNOA_sechel).End(xlUp).Row
    lastCol_sechel = onglet_sechel.Cells(1, 100).End(xlToLeft).Column
    
    ' ---------------------------------------------------------------------------------------------------------
    
    
    
    On Error GoTo 0
    If lastRow_sechel < 2 Then
        lastRow_sechel = 2
    End If
    '2.4.Vider contenu actuel onglet sechel
    ' original
    ' onglet_sechel.Range(onglet_sechel.Cells(2, 1), onglet_sechel.Cells(lastRow_sechel, lastCol_sechel)).ClearContents
    ' a bit improvement from Forrest:
    
    onglet_sechel.Range(onglet_sechel.Cells(2, 1), onglet_sechel.Cells(lastRow_sechel, 1)).EntireRow.Delete xlShiftUp
    
    ' for mass clear up - not in use!
    ' ------------------------------------------------------------------------------------------------------------
    'onglet_sechel.Range(onglet_sechel.Cells(2, 1), onglet_sechel.Cells(1000000, 1)).EntireRow.Delete xlShiftUp
    ' ------------------------------------------------------------------------------------------------------------
    
    '2.5.Ajouter filtres
    'Préparation filtrage des Reçus
    onglet_extract.Cells(1, lastCol_extract + 1).Value = "TMP1"
    onglet_extract.Cells(2, lastCol_extract + 1).FormulaR1C1 = "=RC[-2]-RC[-1]"
    onglet_extract.Cells(2, lastCol_extract + 1).AutoFill Destination:=onglet_extract.Range(onglet_extract.Cells(2, lastCol_extract + 1), onglet_extract.Cells(lastRow_extract, lastCol_extract + 1)), Type:=xlFillDefault
    onglet_extract.Calculate
    onglet_extract.Range(onglet_extract.Cells(2, lastCol_extract + 1), onglet_extract.Cells(lastRow_extract, lastCol_extract + 1)).Value = onglet_extract.Range(onglet_extract.Cells(2, lastCol_extract + 1), onglet_extract.Cells(lastRow_extract, lastCol_extract + 1)).Value
    onglet_extract.Cells(1, lastCol_extract + 1) = "filtre reçus"
    
    'Préparation filtrage des FNR non désirés
    Dim lastRowExcludedFNR As Long
    Dim excludedFNR As Boolean
    Dim col_excluded As Integer
    col_excluded = Sheets("Macro & projects infos").Rows(4).Find(What:="To exclude", LookAt:=xlWhole).Column
    lastRowExcludedFNR = Sheets("Macro & projects infos").Cells(1000, col_excluded).End(xlUp).Row
    excludedFNR = True
    
    
    ' THIS PORTION TO BE TESTED!
    ' ================================================================================================================
    
    ' lastRowExcludedFNR is param for managing potential lines from input file to be excluded by cofor and article
    If lastRowExcludedFNR >= 6 Then
        'ajout d'une colonne temporaire pour prépa filtres
        onglet_extract.Range(onglet_extract.Cells(2, 21), onglet_extract.Cells(2 + lastRowExcludedFNR - 6, 21)).Value = _
            Sheets("Macro & projects infos").Range(Sheets("Macro & projects infos").Cells(6, col_excluded), Sheets("Macro & projects infos").Cells(lastRowExcludedFNR, col_excluded)).Value

        On Error Resume Next
        
        'forcer les articles en texte pour que la formules fonctionne dans tous les cas
        onglet_extract.Columns(4).TextToColumns , DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=True, Comma:=True, Space:=False, Other:=False, FieldInfo:= _
        Array(1, 2), TrailingMinusNumbers:=True
        
        onglet_extract.Columns(21).TextToColumns , DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=True, Comma:=True, Space:=False, Other:=False, FieldInfo:= _
        Array(1, 2), TrailingMinusNumbers:=True
        
        On Error GoTo 0
        
        
        'formule avec condition d'exclusion
        onglet_extract.Cells(2, lastCol_extract + 2).FormulaR1C1 = "=IFERROR(VLOOKUP(RC4,C21:C21,1,0),""à conserver"")"
        onglet_extract.Cells(2, lastCol_extract + 2).AutoFill Destination:=onglet_extract.Range(onglet_extract.Cells(2, lastCol_extract + 2), onglet_extract.Cells(lastRow_extract, lastCol_extract + 2)), Type:=xlFillDefault
        onglet_extract.Calculate
        onglet_extract.Range(onglet_extract.Cells(2, lastCol_extract + 2), onglet_extract.Cells(lastRow_extract, lastCol_extract + 2)).Value = onglet_extract.Range(onglet_extract.Cells(2, lastCol_extract + 2), onglet_extract.Cells(lastRow_extract, lastCol_extract + 2)).Value
        onglet_extract.Cells(1, lastCol_extract + 2) = "filtre fnr"
    Else
        excludedFNR = False
    End If
    
    ' ================================================================================================================
    
    
    ' THIS PORTION TO BE TESTED!
    ' ================================================================================================================
    
    'Préparation filtrage des REF non désirées
    Dim lastRowExcludedREF As Long
    Dim excludedREF As Boolean
    lastRowExcludedREF = Sheets("Macro & projects infos").Cells(1000, col_excluded + 1).End(xlUp).Row
    excludedREF = True
    If lastRowExcludedREF >= 6 Then
        'ajout d'une colonne temporaire pour prépa filtres
        onglet_extract.Range(onglet_extract.Cells(2, 21), onglet_extract.Cells(2 + lastRowExcludedREF - 6, 21)).Value = Sheets("Macro & projects infos").Range(Sheets("Macro & projects infos").Cells(6, col_excluded + 1), Sheets("Macro & projects infos").Cells(lastRowExcludedREF, col_excluded + 1)).Value

        On Error Resume Next
        
        'forcer les articles en texte pour que la formules fonctionne dans tous les cas
        onglet_extract.Columns(2).TextToColumns , DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=True, Comma:=True, Space:=False, Other:=False, FieldInfo:= _
        Array(1, 2), TrailingMinusNumbers:=True
        
        onglet_extract.Columns(21).TextToColumns , DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=True, Comma:=True, Space:=False, Other:=False, FieldInfo:= _
        Array(1, 2), TrailingMinusNumbers:=True
        
        On Error GoTo 0
        
        'formule avec condition d'exclusion
        onglet_extract.Cells(2, lastCol_extract + 3).FormulaR1C1 = "=IFERROR(VLOOKUP(RC2,C21:C21,1,0),""à conserver"")"
        onglet_extract.Cells(2, lastCol_extract + 3).AutoFill Destination:=onglet_extract.Range(onglet_extract.Cells(2, lastCol_extract + 3), onglet_extract.Cells(lastRow_extract, lastCol_extract + 3)), Type:=xlFillDefault
        onglet_extract.Calculate
        onglet_extract.Range(onglet_extract.Cells(2, lastCol_extract + 3), onglet_extract.Cells(lastRow_extract, lastCol_extract + 3)).Value = onglet_extract.Range(onglet_extract.Cells(2, lastCol_extract + 3), onglet_extract.Cells(lastRow_extract, lastCol_extract + 3)).Value
        onglet_extract.Cells(1, lastCol_extract + 3) = "filtre REF"
        
        
        
    Else
        excludedREF = False
    End If
    
    ' ================================================================================================================
    
    onglet_extract.Cells(1, lastCol_extract + 2) = "filtre fnr"
    onglet_extract.Cells(1, lastCol_extract + 3) = "filtre REF"
    
    With onglet_extract.Range("A1")
        .AutoFilter Field:=6, Criteria1:="<>4" ' this filter is for RU param - not valid for xF
        '.AutoFilter Field:=11, Criteria1:="E*"
        .AutoFilter Field:=15, Criteria1:="<>0"
        If excludedFNR = True Then
            .AutoFilter Field:=16, Criteria1:="à conserver", Operator:=xlFilterValues
        End If
        If excludedREF = True Then
            .AutoFilter Field:=17, Criteria1:="à conserver", Operator:=xlFilterValues
        End If
    End With
    
    
    ' 3.Transfert des données de l'extract vers l'onglet sechel
    ' fichier_extract.Close (True)
    onglet_extract.Range(onglet_extract.Cells(2, colNOA_extract), onglet_extract.Cells(lastRow_extract, lastCol_extract)).Copy
    onglet_sechel.Activate
    onglet_sechel.Cells(2, colNOA_sechel).Select
    Selection.PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False

    '4.Mettre à jour la liste des codes NOA (en conservant le nom également)
    '4.1.Définition des variables utilisées
    Dim enteteCodeNOA As Long
    Dim lastRowCodeNOA As Long
    Dim CodeNOA As String
    Dim NameNOA As String
    Dim foundCodeNOA As String
    Dim missingCount As Long
    
    '4.2.Côté extract, on supprime les doublons du code NOA, puis redéfinis la dernière ligne
    onglet_extract.Cells.RemoveDuplicates Columns:=Array(colNOA_id_extract), Header:=xlYes
    lastRow_extract = onglet_extract.Cells(1000000, colNOA_id_extract).End(xlUp).Row
    lastRowCodeNOA = onglet_lisezmoi.Cells(1000000, 1).End(xlUp).Row
    

    missingCount = 0
    '4.3.Balayage des codes NOA trouvés dans l'extract --> si non trouvées, on l'ajoute à la fin de la liste prévue pour l'extract ME9E
    i = 2
    While onglet_extract.Cells(i, colNOA_id_extract).Value <> ""
        CodeNOA = onglet_extract.Cells(i, colNOA_id_extract).Value
        'NameNOA = onglet_extract.Cells(i, colNOA_extract).Value
        foundCodeNOA = ""
        On Error Resume Next
        foundCodeNOA = onglet_lisezmoi.Columns(1).Find(What:=CodeNOA, LookAt:=xlWhole).Row
        On Error GoTo 0
        If foundCodeNOA = "" Then 'Si le code NOA n'a pas été trouvé, on l'ajoute
            onglet_lisezmoi.Cells(lastRowCodeNOA + 1, 1) = CodeNOA
            onglet_lisezmoi.Cells(lastRowCodeNOA + 1, 4).Borders.LineStyle = xlContinuous
            onglet_lisezmoi.Cells(lastRowCodeNOA + 1, 3).Borders.LineStyle = xlContinuous
            onglet_lisezmoi.Cells(lastRowCodeNOA + 1, 2).Borders.LineStyle = xlContinuous
            onglet_lisezmoi.Cells(lastRowCodeNOA + 1, 1).Borders.LineStyle = xlContinuous

            lastRowCodeNOA = lastRowCodeNOA + 1
            missingCount = missingCount + 1
        End If
        i = i + 1
    Wend
    
    '5.Fermer le 2e processus Excel ouvert temporairement (sinon il restera et ralentira le PC)
    fichier_extract.Close False
    MonAppliExcel.Quit
    Set MonAppliExcel = Nothing
    '6. Affichage message final + afficher la liste des codes NOA pour ME9E
    
    If missingCount = 0 Then
        MsgBox ("Import Sigapp-Sechel added correctly." & vbLf & "No new NOA code found"), vbInformation
        Sheets("Macro & projects infos").Select
    Else
        onglet_lisezmoi.Visible = True
        onglet_lisezmoi.Select
        On Error Resume Next
        enteteCodeNOA = onglet_lisezmoi.Columns(1).Find(What:="NOM NOA", LookAt:=xlWhole).Row
        Application.GoTo onglet_lisezmoi.Cells(enteteCodeNOA, 1)
        On Error GoTo 0
        MsgBox ("Import Sigapp-Sechel ajouté correctement." & vbLf & "De nouveau(x) code(s) NOA trouvé(s) : " & missingCount), vbInformation
    End If
    Exit Sub
    
erreurFormat:
    MsgBox "Erreur : le format d'entrée n'est pas le bon. La macro se base sur le 1er onglet disponible. Veillez vérifier le format attendu sur le logigramme, puis relancer la macro.", vbCritical
    
End Sub

Sub step2()
    Application.StatusBar = "1/5 : Réalisation des imports"



    ' some comments here:
    ' this is really dirty solution for ignoring for now imports:
    '   ME9E
    '   Delivery Program
    '   RLF - contact database
    ' but on the other hand this is really connected with SIGAPP
    ' and how xP environment is working
    ' so for demo taking only Sechel-Sigapp file "act like" should be enough!
    
    ' additional note for v003 -> prog liv (delivery program) will be most
    ' probably replaced with some kind of PUS number
    
    ' ----------------------------------------------------------------
    'Call import_ME9E
    'Call import_progLiv
    'Call import_RLF
    ' ----------------------------------------------------------------
    
    
    ' for DEMO maj_pickup__xF !!! version 001 - minimal version !!!
    Call PUS.MajPickupXFModule.maj_pickup__xF
    
    Application.StatusBar = False
End Sub


Sub import_ME9E()

    '1.Désigner les onglets de travail
    '1.1.Ouverture explorateur et selection du fichier d'extract
    Dim Fichier As FileDialog
    Set Fichier = Application.FileDialog(msoFileDialogFilePicker)
    Fichier.Title = "Sélection du fichier ME9E"
    Fichier.Show
    If Fichier.SelectedItems.count = 0 Then
        MsgBox "Aucun fichier sélectionné"
        End
    End If
    
    '1.2.Déclaration des variables Excel
    Dim MonAppliExcel As Excel.Application 'Application Excel générale
    Dim fichier_pus As Excel.Workbook 'Fichiers
    Dim fichier_extract As Excel.Workbook
    Dim onglet_ME9E As Worksheet 'Onglets
    Dim onglet_lisezmoi As Worksheet
    Dim onglet_extract_ion As Worksheet
    Dim onglet_extract As Worksheet
    
    '1.3.Fichier et onglets de la pus
    Set fichier_pus = ThisWorkbook
    Set onglet_ME9E = fichier_pus.Worksheets("ME9E")
    onglet_ME9E.Visible = xlSheetVisible
    Set onglet_lisezmoi = fichier_pus.Worksheets("NOA")
    '1.4.Fichier et onglet de l'extraction
    Set MonAppliExcel = New Excel.Application 'ouverture d'un second excel en parallèle
    Set fichier_extract = MonAppliExcel.Workbooks.Add(Fichier.SelectedItems(1))
    MonAppliExcel.Visible = True
    Set onglet_extract = fichier_extract.Worksheets(1)
    Application.DisplayAlerts = False
    Application.ScreenUpdating = True
    
    onglet_extract.Select
    If Not onglet_extract.AutoFilter Is Nothing Then
    onglet_extract.Cells.AutoFilter
    Else
    End If
    
    '2.Préparation et vérifications des données
    Dim colDocAchat_extract As Long
    Dim colDocAchat_ME9E As Long
    Dim firstLine_extract As Long
    Dim lastRow_extract As Long
    Dim lastRow_ME9E As Long
    Dim lastCol_ME9E As Long
    
    '2.1.Tester la validité des colonnes de l'extraction et les repérer
    On Error GoTo erreurFormat
    colDocAchat_extract = onglet_extract.Cells.Find(What:="Doc achat", LookAt:=xlWhole).Column
    firstLine_extract = onglet_extract.Cells.Find(What:="Doc achat", LookAt:=xlWhole).Row + 1
    
    '2.2.Repérer les colonnes de l'onglet ME9E
    colDocAchat_ME9E = onglet_ME9E.Cells.Find(What:="Doc achat", LookAt:=xlWhole).Column
    
    On Error GoTo 0
    '2.3.Dernière ligne active de l'onglet ME9E ' heuristic 300k - idk - but OK!
    lastRow_ME9E = onglet_ME9E.Cells(300000, colDocAchat_ME9E).End(xlUp).Row
    If lastRow_ME9E < 2 Then
        lastRow_ME9E = 2
    End If
    lastRow_extract = onglet_extract.Cells(300000, colDocAchat_extract).End(xlUp).Row
    '2.3.Vider contenu actuel onglet ME9E
    ' this one is not very good
    ' onglet_ME9E.Range(onglet_ME9E.Cells(3, 2), onglet_ME9E.Cells(lastRow_ME9E, 2)).ClearContents
    
    ' forrest proposal:
    onglet_ME9E.Range(onglet_ME9E.Cells(3, 2), onglet_ME9E.Cells(lastRow_ME9E, 2)).EntireRow.Delete xlShiftUp
    
    ' temp global clear:
    ' onglet_ME9E.Range(onglet_ME9E.Cells(3, 2), onglet_ME9E.Cells(1000000, 2)).EntireRow.Delete xlShiftUp
    
    '3.Transfert des données de l'extract vers l'onglet ME9E
    
    ' i do not like this one-liner:
    onglet_ME9E.Range(onglet_ME9E.Cells(3, colDocAchat_ME9E), onglet_ME9E.Cells(lastRow_extract - firstLine_extract + 3, colDocAchat_ME9E)).Value = _
        onglet_extract.Range(onglet_extract.Cells(firstLine_extract, colDocAchat_extract), onglet_extract.Cells(lastRow_extract, colDocAchat_extract)).Value


    ' NOK - taking too long!
    'Dim ir As Range, i As Variant
    'i = 3
    '' better to make quick loop through all data from original file a then put then as normal list
    'For Each ir In onglet_extract.Range(onglet_extract.Cells(firstLine_extract, colDocAchat_extract), onglet_extract.Cells(lastRow_extract, colDocAchat_extract))
    '
    '    If Trim(ir.Value) <> "" Then
    '        onglet_ME9E.Cells(i, colDocAchat_ME9E).Value = ir.Value
    '        i = i + 1
    '    End If
    'Next ir
    
    '4.Fermer le 2e processus Excel ouvert temporairement (sinon il restera et ralentira le PC)
    fichier_extract.Close (False)
    MonAppliExcel.Quit
    Set MonAppliExcel = Nothing
    Exit Sub
    
erreurFormat:
    MsgBox "Erreur : le format d'entrée n'est pas le bon. La macro se base sur le 1er onglet disponible. Veillez vérifier le format attendu sur le logigramme, puis relancer la macro.", vbCritical
    End
End Sub

Sub import_progLiv()
    
    '1.Désigner les onglets de travail
    '1.1.Ouverture explorateur et selection du fichier d'extract
    Dim Fichier As FileDialog
    Set Fichier = Application.FileDialog(msoFileDialogFilePicker)
    Fichier.Title = "Sélection du fichier ProgLiv (ONL_LIST_ACHAT)"
    Fichier.Show
    If Fichier.SelectedItems.count = 0 Then
        MsgBox "Aucun fichier sélectionné"
        End
    End If
    '1.2.Déclaration des variables Excel
    Dim MonAppliExcel As Excel.Application 'Application Excel générale
    Dim fichier_pus As Excel.Workbook 'Fichiers
    Dim fichier_extract As Excel.Workbook
    Dim onglet_prog As Worksheet 'Onglets
    Dim onglet_lisezmoi As Worksheet
    Dim onglet_extract_ion As Worksheet
    Dim onglet_extract As Worksheet
    
    '1.3.Fichier et onglets de la pus
    Set fichier_pus = ThisWorkbook
    Set onglet_prog = fichier_pus.Worksheets("progliv")
    Set onglet_lisezmoi = fichier_pus.Worksheets("NOA")
    
    '1.4.Fichier et onglet de l'extraction
    Set MonAppliExcel = New Excel.Application 'ouverture d'un second excel en parallèle
    Set fichier_extract = MonAppliExcel.Workbooks.Add(Fichier.SelectedItems(1))
    Set onglet_extract = fichier_extract.Worksheets(1)
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    onglet_extract.Select
    If Not onglet_extract.AutoFilter Is Nothing Then
    onglet_extract.Cells.AutoFilter
    Else
    End If
    
'2.Préparation et vérifications des données
    Dim colDocAchat_extract As Long
    Dim colContrat_extract As Long
    Dim colDocAchat_prog As Long
    Dim colContrat_prog As Long
    Dim firstLine_extract As Long
    Dim lastRow_extract As Long
    Dim lastRow_prog As Long
    Dim lastCol_prog As Long
    '2.1.Tester la validité des colonnes de l'extraction et les repérer
    On Error GoTo erreurFormat
    colDocAchat_extract = onglet_extract.Cells.Find(What:="Doc achat", LookAt:=xlWhole).Column
    colContrat_extract = onglet_extract.Cells.Find(What:="Contrat", LookAt:=xlWhole).Column
    firstLine_extract = onglet_extract.Cells.Find(What:="Doc achat", LookAt:=xlWhole).Row + 1
    '2.2.Repérer les colonnes de l'onglet prog
    colDocAchat_prog = onglet_prog.Cells.Find(What:="Doc achat", LookAt:=xlWhole).Column
    colContrat_prog = onglet_prog.Cells.Find(What:="Contrat", LookAt:=xlWhole).Column
    On Error GoTo 0
    '2.3.Dernière ligne active de l'onglet prog
    lastRow_prog = onglet_prog.Cells(1000000, colDocAchat_prog).End(xlUp).Row
    If lastRow_prog < 2 Then
        lastRow_prog = 2
    End If
    lastRow_extract = onglet_extract.Cells(1000000, colDocAchat_extract).End(xlUp).Row
    '2.3.Vider contenu actuel onglet prog
    
    ' this one is not performing
    ' onglet_prog.Range(onglet_prog.Cells(2, 1), onglet_prog.Cells(lastRow_prog, 2)).ClearContents
    
    'shiftup better
    onglet_prog.Range(onglet_prog.Cells(2, 1), onglet_prog.Cells(lastRow_prog, 2)).EntireRow.Delete xlShiftUp
    
    ' temp mass clearance
    ' onglet_prog.Range(onglet_prog.Cells(2, 1), onglet_prog.Cells(1000000, 2)).EntireRow.Delete xlShiftUp
    
    '3.Transfert des données de l'extract vers l'onglet prog
    onglet_prog.Range(onglet_prog.Cells(2, colDocAchat_prog), onglet_prog.Cells(lastRow_extract - firstLine_extract + 2, colDocAchat_prog)).Value = _
        onglet_extract.Range(onglet_extract.Cells(firstLine_extract, colDocAchat_extract), onglet_extract.Cells(lastRow_extract, colDocAchat_extract)).Value
        
    onglet_prog.Range(onglet_prog.Cells(2, colContrat_prog), onglet_prog.Cells(lastRow_extract - firstLine_extract + 2, colContrat_prog)).Value = _
        onglet_extract.Range(onglet_extract.Cells(firstLine_extract, colContrat_extract), onglet_extract.Cells(lastRow_extract, colContrat_extract)).Value

    
'4.Fermer le 2e processus Excel ouvert temporairement (sinon il restera et ralentira le PC)
    fichier_extract.Close (False)
    MonAppliExcel.Quit
    Set MonAppliExcel = Nothing
    Exit Sub
erreurFormat:
    MsgBox "Erreur : le format d'entrée n'est pas le bon. La macro se base sur le 1er onglet disponible. Veillez vérifier le format attendu sur le logigramme, puis relancer la macro.", vbCritical
    End
End Sub

Sub import_RLF()
    
    '1.Désigner les onglets de travail
    '1.1.Ouverture explorateur et selection du fichier d'extract
    Dim Fichier As FileDialog
    Set Fichier = Application.FileDialog(msoFileDialogFilePicker)
    Fichier.Title = "Sélection du fichier RLF"
    Fichier.Show
    If Fichier.SelectedItems.count = 0 Then
        MsgBox "Aucun fichier sélectionné"
        End
    End If
    '1.2.Déclaration des variables Excel
    Dim MonAppliExcel As Excel.Application 'Application Excel générale
    Dim fichier_pus As Excel.Workbook 'Fichiers
    Dim fichier_extract As Excel.Workbook
    Dim onglet_RLF As Worksheet 'Onglets
    Dim onglet_extract As Worksheet
    '1.3.Fichier et onglets de la pus
    Set fichier_pus = ActiveWorkbook
    Set onglet_RLF = fichier_pus.Worksheets("RLF")
    '1.4.Fichier et onglet de l'extraction
    Set MonAppliExcel = New Excel.Application 'ouverture d'un second excel en parallèle
    Set fichier_extract = MonAppliExcel.Workbooks.Add(Fichier.SelectedItems(1))
    Set onglet_extract = fichier_extract.Worksheets("Synthese")
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    
    '2.Préparation et vérifications des données
    Dim firstLine_extract As Long
    Dim lastRow_extract As Long
    Dim lastRow_RLF As Long
    Dim lastCol_RLF As Long
    
    onglet_extract.Select
    If Not onglet_extract.AutoFilter Is Nothing Then
    onglet_extract.Cells.AutoFilter
    Else
    End If
    
    '2.1.Tester la validité des colonnes de l'extraction et les repérer
    On Error GoTo erreurFormat
    firstLine_extract = onglet_extract.Cells.Find(What:="Article 10 C", LookAt:=xlWhole).Row + 1
    On Error GoTo 0
    '2.2.Dernière ligne active de l'onglet RLF
    lastRow_RLF = onglet_RLF.Cells(1000000, 1).End(xlUp).Row
    If lastRow_RLF < 2 Then
        lastRow_RLF = 2
    End If
    lastRow_extract = onglet_extract.Cells(1000000, 11).End(xlUp).Row
    '2.2.Vider contenu actuel onglet RLF
    onglet_RLF.Range(onglet_RLF.Cells(2, 1), onglet_RLF.Cells(lastRow_RLF, 16)).ClearContents
    
'3.Transfert des données de l'extract vers l'onglet RLF
    onglet_extract.Range(onglet_extract.Cells(2, 3), onglet_extract.Cells(lastRow_extract, 13)).Copy
   ' onglet_RLF.Cells(2, 1).Select
    onglet_RLF.Cells(2, 1).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    
    Dim Cible As DataObject
    Set Cible = New DataObject
    Cible.SetText ""
    Cible.PutInClipboard
    Set Cible = Nothing

'4.Fermer le 2e processus Excel ouvert temporairement (sinon il restera et ralentira le PC)
    fichier_extract.Close (False)
    Exit Sub
erreurFormat:
    MsgBox "Erreur : le format d'entrée n'est pas le bon. La macro se base sur le 1er onglet disponible. Veillez vérifier le format attendu sur le logigramme, puis relancer la macro.", vbCritical
    End
End Sub

'Module ajouté par GZA (Alten)
Sub import_sechel_comments()
    
'1.Désigner les onglets de travail
    '1.1.Ouverture explorateur et selection du fichier d'extract
    Dim Fichier As FileDialog
    Set Fichier = Application.FileDialog(msoFileDialogFilePicker)
    Fichier.Title = "Sélection du fichier SECHEL / PROCURE"
    Fichier.Show
    If Fichier.SelectedItems.count = 0 Then
        MsgBox "Aucun fichier sélectionné"
        Exit Sub 'Stopper la macro
    End If
    '1.2.Déclaration des variables Excel
    Dim MonAppliExcel As Excel.Application 'Application Excel générale
    Dim fichier_pus As Excel.Workbook 'Fichiers
    Dim fichier_extract As Excel.Workbook
    Dim onglet_sechel_comments As Worksheet 'Onglets
    Dim onglet_sechel As Worksheet
    Dim onglet_base As Worksheet 'Onglets
    Dim onglet_extract As Worksheet
    '1.3.Fichier et onglets de la pus
    Set fichier_pus = ActiveWorkbook
    Set onglet_sechel_comments = fichier_pus.Worksheets("sechel_comments")
    Set onglet_base = fichier_pus.Worksheets("BASE")
    Set onglet_sechel = fichier_pus.Worksheets("sechel")
    
    '1.4.Fichier et onglet de l'extraction
    Set MonAppliExcel = New Excel.Application 'ouverture d'un second excel en parallèle
    Set fichier_extract = MonAppliExcel.Workbooks.Add(Fichier.SelectedItems(1))
    Dim extractVersion As String

    
    On Error Resume Next
    Set onglet_extract = fichier_extract.Worksheets("Echéances")
    extractVersion = "sechel"
    
    If onglet_extract Is Nothing Then
        
        Set onglet_extract = fichier_extract.Worksheets("Programme")
        extractVersion = "VueActuelleEtGlobale"
        
        If onglet_extract Is Nothing Then 'Version FR ?
            Set onglet_extract = fichier_extract.Worksheets("Included")
            extractVersion = "procureEN"
            
            If onglet_extract Is Nothing Then 'Version FR ?
            Set onglet_extract = fichier_extract.Worksheets("Inclus")
            extractVersion = "procureFR"
            
            
                If onglet_extract Is Nothing Then 'erreur
                    MsgBox "Macro cannot identify the sheet"
                    fichier_extract.Close (False)
                    On Error GoTo 0
                    Exit Sub
                End If
            End If
        End If
    End If
    
    
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    onglet_sechel.Visible = True
    onglet_sechel_comments.Visible = True
    
    onglet_extract.Select
    If Not onglet_extract.AutoFilter Is Nothing Then
    onglet_extract.Cells.AutoFilter
    Else
    End If
    
'2.Préparation et vérifications des données
    Dim colNOA_extract, colDateLundi_extract, colComment_extract, colClé_sechel, colArticle_extract, colDateLundi_sechel, colComment_sechel, colDesignation_extract, colFourn_extract, colNomFourn_extract, colRU_extract, colDateEch_extract, colQteEch_extract, colQteLiv_extract, colMag_extract, colDocAchat_extract, colGAc_extract As Variant
    Dim colNOA_sechel, colCW_sechel, colArticle_sechel, colDesignation_sechel, colFourn_sechel, colNomFourn_sechel, colRU_sechel, colDateEch_sechel, colQteEch_sechel, colQteLiv_sechel, colMag_sechel, colDocAchat_sechel, colGAc_sechel As Variant

    Dim lastRow_extract As Long
    Dim lastCol_extract As Long
    Dim lastRow_sechel As Long
    Dim lastCol_sechel As Long
    Dim i As Long
    '2.1.Tester la validité des colonnes de l'extraction et les repérer
    On Error GoTo erreurFormat
    
    
    
    If extractVersion = "sechel" Then
        colArticle_extract = onglet_extract.Rows(1).Find(What:="Article", LookAt:=xlWhole).Column
        colMag_extract = onglet_extract.Rows(1).Find(What:="Mag", LookAt:=xlWhole).Column
        colDateLundi_extract = onglet_extract.Rows(1).Find(What:="Date_Ech_Lundi", LookAt:=xlWhole).Column
        colComment_extract = onglet_extract.Rows(1).Find(What:="Commentaires", LookAt:=xlWhole).Column
        
        ElseIf extractVersion = "VueActuelleEtGlobale" Then
            colArticle_extract = 0
            On Error Resume Next
            colArticle_extract = onglet_extract.Rows(1).Find(What:="Article", LookAt:=xlWhole).Column
            On Error GoTo 0
            If colArticle_extract = 0 Then
                extractVersion = "procureEN"
                Else
                extractVersion = "procureFR"
            End If
    End If
    
    If extractVersion = "procureFR" Then
            colArticle_extract = onglet_extract.Rows(1).Find(What:="Article", LookAt:=xlWhole).Column
            colMag_extract = onglet_extract.Rows(1).Find(What:="Mag", LookAt:=xlWhole).Column
            colDateLundi_extract = onglet_extract.Rows(1).Find(What:="Date Ech Lundi", LookAt:=xlWhole).Column
            colComment_extract = onglet_extract.Rows(1).Find(What:="Dernier commentaire", LookAt:=xlWhole).Column
            
        ElseIf extractVersion = "procureEN" Then
            colArticle_extract = onglet_extract.Rows(1).Find(What:="Part", LookAt:=xlWhole).Column
            colMag_extract = onglet_extract.Rows(1).Find(What:="Warehouse", LookAt:=xlWhole).Column
            colDateLundi_extract = onglet_extract.Rows(1).Find(What:="PO Date Monday", LookAt:=xlWhole).Column
            colComment_extract = onglet_extract.Rows(1).Find(What:="Latest Comment", LookAt:=xlWhole).Column
    End If
    
    
    lastRow_extract = onglet_extract.Cells(1000000, colArticle_extract).End(xlUp).Row

    On Error GoTo 0
    '2.3.Repérer les colonnes de l'onglet sechel comments
    colClé_sechel = onglet_sechel_comments.Rows(1).Find(What:="Clé de recherche", LookAt:=xlWhole).Column
    colArticle_sechel = onglet_sechel_comments.Rows(1).Find(What:="Article", LookAt:=xlWhole).Column
    colMag_sechel = onglet_sechel_comments.Rows(1).Find(What:="Mag", LookAt:=xlWhole).Column
    colDateLundi_sechel = onglet_sechel_comments.Rows(1).Find(What:="Date_Ech_Lundi", LookAt:=xlWhole).Column
    colComment_sechel = onglet_sechel_comments.Rows(1).Find(What:="Commentaires", LookAt:=xlWhole).Column
    colCW_sechel = onglet_sechel_comments.Rows(1).Find(What:="CW", LookAt:=xlWhole).Column

    '2.4.Vider contenu actuel onglet sechel
    onglet_sechel_comments.Range(onglet_sechel_comments.Cells(2, 1), onglet_sechel_comments.Cells(1000000, 6)).ClearContents
    
'3.Traitement onglet Sechel_comments
    '3.1.Transfert des données
    onglet_sechel_comments.Range(onglet_sechel_comments.Cells(2, colArticle_sechel), onglet_sechel_comments.Cells(lastRow_extract, colArticle_sechel)).Value = onglet_extract.Range(onglet_extract.Cells(2, colArticle_extract), onglet_extract.Cells(lastRow_extract, colArticle_extract)).SpecialCells(xlCellTypeVisible).Value
    onglet_sechel_comments.Range(onglet_sechel_comments.Cells(2, colMag_sechel), onglet_sechel_comments.Cells(lastRow_extract, colMag_sechel)).Value = onglet_extract.Range(onglet_extract.Cells(2, colMag_extract), onglet_extract.Cells(lastRow_extract, colMag_extract)).Value
    onglet_sechel_comments.Range(onglet_sechel_comments.Cells(2, colDateLundi_sechel), onglet_sechel_comments.Cells(lastRow_extract, colDateLundi_sechel)).Value = onglet_extract.Range(onglet_extract.Cells(2, colDateLundi_extract), onglet_extract.Cells(lastRow_extract, colDateLundi_extract)).Value
    onglet_sechel_comments.Range(onglet_sechel_comments.Cells(2, colComment_sechel), onglet_sechel_comments.Cells(lastRow_extract, colComment_sechel)).Value = onglet_extract.Range(onglet_extract.Cells(2, colComment_extract), onglet_extract.Cells(lastRow_extract, colComment_extract)).Value
    
    '3.2. Ajout d'une colonne pour avoir la date en CW 'v2.8.1.1
    onglet_sechel_comments.Cells(2, colCW_sechel).FormulaR1C1 = "=IF(WEEKNUM(RC4,21)<10,RIGHT(YEAR(RC4),2) & ""-CW0"" & WEEKNUM(RC4,21),RIGHT(YEAR(RC4),2) & ""-CW"" & WEEKNUM(RC4,21))"
    onglet_sechel_comments.Cells(2, colCW_sechel).AutoFill Destination:=onglet_sechel_comments.Range(onglet_sechel_comments.Cells(2, colCW_sechel), onglet_sechel_comments.Cells(lastRow_extract, colCW_sechel)), Type:=xlFillDefault
    onglet_sechel_comments.Calculate
    
    '3.3.Création de la clé de recherche côté
    onglet_sechel_comments.Cells(2, colClé_sechel).FormulaR1C1 = "=RC2&RC3&RC6"
    onglet_sechel_comments.Cells(2, colClé_sechel).AutoFill Destination:=onglet_sechel_comments.Range(onglet_sechel_comments.Cells(2, colClé_sechel), onglet_sechel_comments.Cells(lastRow_extract, colClé_sechel)), Type:=xlFillDefault
    onglet_sechel_comments.Calculate
    onglet_sechel_comments.Range(onglet_sechel_comments.Cells(2, colClé_sechel), onglet_sechel_comments.Cells(lastRow_extract, colClé_sechel)).Value = onglet_sechel_comments.Range(onglet_sechel_comments.Cells(2, colClé_sechel), onglet_sechel_comments.Cells(lastRow_extract, colClé_sechel)).Value
    
    
    
    '3.3.Fermer le 2e processus Excel ouvert temporairement (sinon il restera et ralentira le PC)
    fichier_extract.Close (False)

    'v2.8.1.1 - nouvelle méthode simplifiée
''4.Vérifications préliminaires
'    '4.1.Enlever tous les filtres
'    onglet_sechel.Select
'    If Not onglet_sechel.AutoFilter Is Nothing Then
'    onglet_sechel.Cells.AutoFilter
'    Else
'    End If
'
'    onglet_base.Select
'    If Not onglet_base.AutoFilter Is Nothing Then
'    onglet_base.Cells.AutoFilter
'    Else
'    End If
'
'    '4.2. Est-ce que l'onglet Sechel et BASE ont bien le même nombre de ligne ? -->Si non, cela signifie qu'un copier coller sauvage est apparu
'    lastRow_sechel = onglet_sechel.Cells(1000000, 2).End(xlUp).Row
'    Dim lastRow_base As Long
'    lastRow_base = onglet_base.Cells(1000000, 1).End(xlUp).Row
'
'    If lastRow_sechel + 1 <> lastRow_base Then
'        GoTo erreurCopierColler
'    End If
'
''5.Traitement onglet sechel
'    '5.1.Clé de recherche
'
'    onglet_sechel.Cells(2, 26).FormulaR1C1 = "=RC2&RC10&RC12"
'    onglet_sechel.Cells(2, 26).AutoFill Destination:=onglet_sechel.Range(onglet_sechel.Cells(2, 26), onglet_sechel.Cells(lastRow_sechel, 26)), Type:=xlFillDefault
'    onglet_sechel.Calculate
'    onglet_sechel.Range(onglet_sechel.Cells(2, 26), onglet_sechel.Cells(lastRow_sechel, 26)).Value = onglet_sechel.Range(onglet_sechel.Cells(2, 26), onglet_sechel.Cells(lastRow_sechel, 26)).Value
'    '5.2.Commentaires
'    onglet_sechel.Cells(2, 27).FormulaR1C1 = "=IFERROR(IF(VLOOKUP(RC[-1],sechel_comments!C[-26]:C[-21],5,0)=0,"""",VLOOKUP(RC[-1],sechel_comments!C[-26]:C[-21],5,0)),"""")"
'    onglet_sechel.Cells(2, 27).AutoFill Destination:=onglet_sechel.Range(onglet_sechel.Cells(2, 27), onglet_sechel.Cells(lastRow_sechel, 27)), Type:=xlFillDefault
'    onglet_sechel.Calculate
'    onglet_sechel.Range(onglet_sechel.Cells(2, 27), onglet_sechel.Cells(lastRow_sechel, 27)).Value = onglet_sechel.Range(onglet_sechel.Cells(2, 27), onglet_sechel.Cells(lastRow_sechel, 27)).Value
    
'4.Traitement onglet BASE
    Dim lastRow_base As Long
    lastRow_base = onglet_base.Cells(1000000, 1).End(xlUp).Row
        '4.1. Critère comparaison
    onglet_base.Activate
    onglet_base.Range("CJ3").FormulaR1C1 = "=RC1 & RC" & col_base_mag & " & RC5"
    onglet_base.Range("CJ3").AutoFill Destination:=onglet_base.Range("CJ3:CJ" & lastRow_base), Type:=xlFillDefault
    onglet_base.Calculate
    onglet_base.Range("CJ3:CJ" & lastRow_base).Value = onglet_base.Range("CJ3:CJ" & lastRow_base).Value
    
        '4.2. RechercheV
    onglet_base.Range("AW3").FormulaR1C1 = "=IFERROR(IF(VLOOKUP(RC88,sechel_comments!C1:C5,5,0)=0,"""",VLOOKUP(RC88,sechel_comments!C1:C5,5,0)),"""")"
    onglet_base.Range("AW3").AutoFill Destination:=onglet_base.Range("AW3:AW" & lastRow_base), Type:=xlFillDefault
    onglet_base.Calculate
    onglet_base.Range("AW3:AW" & lastRow_base).Value = onglet_base.Range("AW3:AW" & lastRow_base).Value
      
    onglet_sechel_comments.Range(onglet_sechel_comments.Cells(2, 1), onglet_sechel_comments.Cells(1000000, 6)).ClearContents
    
    
'5. Affichage message final
    MsgBox ("Sechel comments correctly imported."), vbInformation
    onglet_sechel.Visible = False
    onglet_sechel_comments.Visible = False
    Worksheets("Macro & projects infos").Select
  
'6. Traitement erreurs
    Exit Sub
erreurFormat:
    MsgBox "Erreur dans l'extraction : il manque une des quatre colonnes nécessaires --> 'Article', 'Domaine', 'Date_Ech_Lundi' et 'Commentaires'. La macro se base sur l'onglet 'Echéances'.", vbCritical
    fichier_extract.Close (False)
    Exit Sub
erreurCopierColler:
    MsgBox "The actual BASE probably came with a copy-paste. Please relaunch the whole process, in order to avoid mistakes.", vbCritical
    
End Sub



