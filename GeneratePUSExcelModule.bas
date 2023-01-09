Attribute VB_Name = "GeneratePUSExcelModule"
Sub PDFActiveSheet()
'www.contextures.com
'for Excel 2010 and later
Dim wsA As Worksheet
Dim wbA As Workbook
Dim strTime As String
Dim strName As String
Dim strPath As String
Dim strFile As String
Dim strPathFile As String
Dim myFile As Variant
On Error GoTo errHandler

Sheets("pickup_sheet").Select

Set wbA = ActiveWorkbook
Set wsA = ActiveSheet
strTime = Format(Now(), "yyyymmdd\_hhmm")

'get active workbook folder, if saved
strPath = wbA.Path
If strPath = "" Then
  strPath = Application.DefaultFilePath
End If
strPath = strPath & "\"

'replace spaces and periods in sheet name
strName = Replace(wsA.Name, " ", "")
strName = Replace(strName, ".", "_")

'create default name for saving file
nom_fnr = Cells(8, 3)
ref_fnr = Cells(31, 2)
If perimetre_analyse = "per_cofor" Then
strFile = strName & "_" & strTime & "_" & nom_fnr & ".xlsx"
ElseIf perimetre_analyse = "per_reference" Then
strFile = strName & "_" & strTime & "_" & nom_fnr & "_" & ref_fnr & ".xlsx"
Else: End If

strPathFile = strPath & strFile

'use can enter name and
' select folder for file
myFile = Application.GetSaveAsFilename _
    (InitialFileName:=strPathFile, _
        FileFilter:="PDF Files (*.pdf), *.pdf", _
        Title:="Select Folder and FileName to save")

'export to PDF if a folder was selected
If myFile <> "False" Then
    wsA.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=myFile, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False
    'confirmation message with file info
    MsgBox "Fichier PDF crée : le document est à envoyer au fournisseur " _
      & vbCrLf _
      & myFile
End If

exitHandler:
    Exit Sub
errHandler:
    MsgBox "Could not create PDF file"
    Resume exitHandler
End Sub

Sub excelsheet()

    'www.contextures.com
    'for Excel 2010 and later
    Dim wsA As Worksheet
    Dim wbA As Workbook
    Dim strTime As String
    Dim strName As String
    Dim strPath As String
    Dim strFile As String
    Dim strPathFile As String
    Dim myFile As Variant
    On Error GoTo errHandler
    
    Sheets("pickup_sheet").Select
    
    Set wbA = ActiveWorkbook
    Set wsA = ActiveSheet
    strTime = Format(Now(), "yyyymmdd\_hhmm")
    
    'get active workbook folder, if saved
    strPath = wbA.Path
    If strPath = "" Then
      strPath = Application.DefaultFilePath
    End If
    strPath = strPath & "\"
    
    'replace spaces and periods in sheet name
    strName = Replace(wsA.Name, " ", "")
    strName = Replace(strName, ".", "_")
    
    'create default name for savng file
    nom_fnr = Cells(8, 6)
    ref_fnr = Cells(33, 3)
    If perimetre_analyse = "per_cofor" Then
    strFile = strName & "_" & strTime & "_" & nom_fnr & ".xlsx"
    ElseIf perimetre_analyse = "per_reference" Then
    strFile = strName & "_" & strTime & "_" & nom_fnr & "_" & ref_fnr & ".xlsx"
    Else: End If

    strPathFile = strPath & strFile
    
    'use can enter name and
    ' select folder for file
    myFile = Application.GetSaveAsFilename _
        (InitialFileName:=strPathFile, _
            FileFilter:="Excel Files (*.xlsx), *.xlsx", _
            Title:="Select Folder and FileName to save")


    'Définir le oules PROJECT_LABEL à ajouter
    Dim onglets(50) As String
    Dim project As Integer
    ' TEMP_REMARK_FOR_DEMO
    Dim strProject As String
    Dim nbProjet As Integer
    Dim fichier_pus As Excel.Workbook 'Fichiers
    Dim fichier_creer As Excel.Workbook
    Set fichier_pus = ActiveWorkbook
    Dim found As Boolean
    
    Dim copiedSh As Worksheet


    ' Debug.Print fichier_pus.Name
   'Création second Excel
    fichier_pus.Sheets(Array("pickup_sheet", "READ-ME OBLIGATORY!", "subst_pack_form", "SERIAL_PACKAGING_PACKMAN")).Copy

    Set fichier_creer = ActiveWorkbook
    nbProjet = 0

    fichier_pus.Activate
    For i = 0 To 524
        currentLine = Sheets("pickup_sheet").Columns(2).Find(i + 1, LookAt:=xlWhole).Row
        If Sheets("pickup_sheet").Cells(currentLine, 3).Value = 0 Then
            Exit For
        Else
                Debug.Print Sheets("pickup_sheet").Cells(currentLine, 8).Value
                strProject = Sheets("pickup_sheet").Cells(currentLine, 8).Value
                found = False
                
                
                onglets(nbProjet) = "PROJECT_LABEL_" & strProject
                Set copiedSh = Nothing
                On Error Resume Next
                Set copiedSh = fichier_creer.Sheets("PROJECT_LABEL_" & strProject)
                
                
                If copiedSh Is Nothing Then
                    On Error Resume Next
                    Sheets(onglets(nbProjet)).Copy After:=fichier_creer.Sheets(fichier_creer.Sheets.count)
                End If
                
        End If
    Next i

    'export to Excel if a folder was selected
    If myFile <> "False" Then
    
    'fichier_creer.Sheets(2).Select
    fichier_creer.SaveAs myFile
    fichier_creer.Close False
        'confirmation message with file info
        MsgBox "Fichier Excel crée : le document est à envoyer au fournisseur " _
          & vbCrLf _
          & myFile
    End If
    
exitHandler:
        Exit Sub
errHandler:
        MsgBox "Could not create Excel file"
        Resume exitHandler
    
errEtiquette:
        MsgBox "Au moins une étiquette manquante"
    Resume exitHandler

End Sub
