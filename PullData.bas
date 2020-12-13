Attribute VB_Name = "PullData"
Option Explicit
Dim PDFApp As Variant
Dim Folder, ExportFile As String
Dim ClientRow, CustCol, DataCol, DataRow, LastDataRow As Long
Dim fn As String
Dim Papp As PhantomPDF.Application
Dim Pdoc As PhantomPDF.Document
Dim PExl As Variant
Dim PathNameXL As String
Dim PDFFolder As FileDialog


Dim FiletoOpen As Variant
Dim SelectedBook As Workbook
Dim wbk As Workbook
Dim rng As Integer
Dim YesNo As Integer
Dim MyInp As String
Dim Pathname As String

Dim FileXL As String



Sub GetFolder()


Set PDFFolder = Application.FileDialog(msoFileDialogFolderPicker)
With PDFFolder
.Title = "Select Folder"
If .Show <> -1 Then GoTo NoSet
MainData.Range("FolderLocation").Value = .SelectedItems(1)
End With


Call OpenPDF
NoSet:
End Sub


Sub OpenPDF()

With Application
.StatusBar = "WAIT"
.ScreenUpdating = False
.DisplayAlerts = False


End With

Set Papp = CreateObject("PhantomPDF.Application")

With MainData

 If MainData.Range("FolderLocation").Value = Empty Then
    MsgBox " Please browse for your PDF folder"
    GetFolder
    Exit Sub

 End If

Folder = .Range("FolderLocation").Value ' PDF folder location

On Error GoTo Handle
ChDir Folder

    fn = Application.GetOpenFilename("PDF Files,*.pdf,", _
        1, "Technician Technical Information - Select folder and file to open", , False)
    If TypeName(fn) = "" Then Exit Sub
    ' the user didn't select a file

    Select Case Right(fn, 3)

    Case Is = "pdf"
    
       Set Pdoc = Papp.OpenDocument(fn, "", True, True)
      
       'Call PullData_Data
   Case Is = "xls"
   MsgBox " Excel file was picked and PDF is required"
    Case Else
        MsgBox "No file was selected"
    Exit Sub
    End Select

End With


PathNameXL = fn

Application.Wait Now + 0.00001
 
On Error Resume Next
Call Pdoc.OCRAndExportToExcel(PathNameXL & ".xlsx", 1, 1, True, True)

FileXL = PathNameXL & ".xlsx"


Set SelectedBook = Application.Workbooks.Open(FileXL)



Pathname = SelectedBook.FullName

        SelectedBook.Sheets(1).Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
        SelectedBook.Close
        Kill (Pathname)

        ThisWorkbook.ActiveSheet.Name = "RawData"
Application.DisplayAlerts = False

rng = ThisWorkbook.Sheets(1).Columns(1).End(xlDown).End(xlDown).End(xlUp).Offset(1, 0).Row

Dim Zone, Depth, Fluid, STARTDate, StageStartTime, EndDate, StageEndTime, MannedStandby, PortOpenPressure, _
BreakdownPressure, AveragePressure, MaximumPressure, MinimumPressure, ISIP, AverageSlurryRate, MaximumSlurryRate, _
MinimumSlurryRate, MaximumPropCon, FracClean, FracSlurry, CTClean, TotalClean, BigSand, TotalDesigned, TotalPlaced, MMFR As String


'zone


MainData.Activate
MyInp = VBA.InputBox(" Please enter the item number")
If MyInp = "" Then
ThisWorkbook.Sheets("RawData").Delete
Exit Sub
End If
MainData.Cells(rng, 1).Value = MyInp

'Depth
Set Depth = Worksheets("RawData").Range("A:A").Find(what:=MainData.Cells(2, 2), LookIn:=xlValues, lookat:=xlPart)
MainData.Cells(rng, 2).Value = Depth.Offset(0, 1).Value
'STARTDate
Set STARTDate = Worksheets("RawData").Range("A:A").Find(what:=MainData.Cells(2, 3), LookIn:=xlValues, lookat:=xlPart)
MainData.Cells(rng, 3).Value = Depth.Offset(-5, 1).Value
'StageStartTime
MainData.Cells(rng, 4).Value = Depth.Offset(-3, 1).Value
'EndDate
MainData.Cells(rng, 5).Value = Depth.Offset(-4, 1).Value
'StageEndTime
MainData.Cells(rng, 6).Value = Depth.Offset(-2, 1).Value

'Fluid
Set Fluid = Worksheets("RawData").Range("B:B").Find(what:=MainData.Cells(2, 10), LookIn:=xlValues, lookat:=xlPart)
MainData.Cells(rng, 10).Value = Fluid.Offset(2, 0).Value
'PortOpen
MainData.Cells(rng, 11).Value = Depth.Offset(2, 1).Value
'Breakdown
MainData.Cells(rng, 12).Value = Depth.Offset(3, 1).Value
'AveragePress
MainData.Cells(rng, 13).Value = Depth.Offset(4, 1).Value
'MaxPressure
MainData.Cells(rng, 14).Value = Depth.Offset(5, 1).Value
'MinPressure
MainData.Cells(rng, 15).Value = Depth.Offset(6, 1).Value
'ISIP
MainData.Cells(rng, 16).Value = Depth.Offset(7, 1).Value
'AverageRate
MainData.Cells(rng, 17).Value = Depth.Offset(2, 3).Value
'Max Rate
MainData.Cells(rng, 18).Value = Depth.Offset(3, 3).Value
'Minrate
MainData.Cells(rng, 19).Value = Depth.Offset(4, 3).Value
'CT Volume
MainData.Cells(rng, 23).Value = Depth.Offset(7, 3).Value
' 40/70sand
Set BigSand = Worksheets("RawData").Range("A:A").Find(what:=MainData.Cells(2, 26), LookIn:=xlValues, lookat:=xlWhole, Searchdirection:=xlPrevious)
MainData.Cells(rng, 26).Formula = 1000 * BigSand.Offset(0, 2).Value

'Max propcon
MainData.Cells(rng, 3).Offset(0, 17).Value = BigSand.Offset(0, 3).Value
'Total Design Sand

MainData.Cells(rng, 3).Offset(0, 24).Formula = 1000 * BigSand.Offset(0, 1).Value

'FracClean
MainData.Cells(rng, 3).Offset(0, 18).Value = BigSand.Offset(-3, 2).Value

'fracSlurry
MainData.Cells(rng, 3).Offset(0, 19).Value = BigSand.Offset(-3, 3).Value
'

ThisWorkbook.Sheets("RawData").Delete

Application.DisplayAlerts = True
MainData.Cells(rng, 1).Activate


MsgBox "All information is entered"

With Application
.StatusBar = ""
.ScreenUpdating = True
.DisplayAlerts = True

End With

Handle:

If Err.Number = 76 Then
MsgBox "No path found. Click on Browse Folder to continue", vbInformation, "Please select folder on your computer"
End If
Exit Sub

End Sub


