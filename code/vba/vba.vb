Sub GetData()

 Application.ScreenUpdating = False
 Application.EnableEvents = False
 Application.DisplayAlerts = False

'Setting workbook and sheets
 Dim wb As Workbook, wsSB As Worksheet, wsHB As Worksheet
 Set wb = ActiveWorkbook
 Set wsSB = wb.Sheets("SB")
 
'User prompt file input
 MsgBox "Vennligst velg plassering for hovedbok", vbInformation, "Plassering av hovedbok"

'Declaring variables and File Dialog
 Dim intChoice As Integer
 Dim strPath As String

    'One file selection limit
     Application.FileDialog(msoFileDialogOpen).AllowMultiSelect = False

    'File Dialog visible
     intChoice = Application.FileDialog(msoFileDialogOpen).Show

    'Controlling selection
     If intChoice <> 0 Then

        'Reading file path
         strPath = Application.FileDialog(msoFileDialogOpen).SelectedItems(1)
     End If

'Controlling input file extension
 Dim Extention, XLSBox
 Extention = Right(strPath, Len(strPath) - InStrRev(strPath, "."))
 If Extention = "xls" Then
    XLSBox = MsgBox(".xls-filer har en begrensning på 65 536 rader. Hovedboken kan derfor være ufullstendig. Vil du likevel kjøre makroen?", vbYesNo + vbQuestion, "Kontroller hovedbok")
    If XLSBox = vbNo Then Exit Sub
 End If
 If Extention <> "xlsx" And Extention <> "xls" And Extention <> "xlsm" And Extention <> "csv" Then
     MsgBox "Ugyldig format på inputdata, følgende filformater støttes:" & vbNewLine & _
     vbNewLine & _
     ".xls" & vbNewLine & _
     ".xlsx" & vbNewLine & _
     ".xlsm" & vbNewLine & _
     ".csv" & vbNewLine & _
     vbNewLine & _
     "Kontakt NO Automation for hjelp", vbInformation
     Exit Sub
 End If

'Updating statusbar and cursor
 Application.Cursor = xlWait
 Application.DisplayStatusBar = True
 Application.StatusBar = "Formaterer hovedbok..."
 frmPleaseWait.Show vbModeless
 frmPleaseWait.Repaint
 
'Copying HB from input file
 Dim fromFile As String
 fromFile = strPath

 Dim inputWB As Workbook
 If Extention = "csv" Then
    Workbooks.OpenText Filename:=fromFile, Origin:=65001, SemiColon:=True, Local:=True
    Set inputWB = ActiveWorkbook
 Else
    Set inputWB = Application.Workbooks.Open(fromFile, _
                 UpdateLinks:=False, _
                ReadOnly:=True, _
                AddToMRU:=False)
 End If

 inputWB.Sheets(1).Copy After:=wb.Sheets(wb.Sheets.Count)

'Closing input file, renaming activesheet and extracting input extention
 inputWB.Close False
 Set wsHB = ActiveSheet
 ActiveSheet.Name = "HB"

'Calling macro to define format
 If Extention = "xlsx" Or Extention = "xls" Or Extention = "xlsm" Then
    'Formatting sheet to "General"
     wsHB.Cells.NumberFormat = "General"
     Call mod2_0_FormatExcel.FormatExcel(wb, wsSB, wsHB)
 ElseIf Extention = "csv" Then
     Call mod3_0_FormatCSV.FormatCSV(wb, wsSB, wsHB)
 End If
           
 If wb.Sheets("Top").Range("A1").Value = "" Then
    MsgBox "Ukjent format på hovedbok, kontakt NO Automation.", vbInformation
    wb.Sheets("HB").Delete
 End If
 
 wb.Sheets("Top").Range("A1").Value = ""
 wb.Sheets("Top").Activate
 frmPleaseWait.Hide
 Application.Cursor = xlDefault
 Application.DisplayAlerts = True
 Application.StatusBar = False
 Application.ScreenUpdating = True
 Application.EnableEvents = True

End Sub

