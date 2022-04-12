# ImportlExport
Excel模板之间数据导入导出
Option Base 1

Public Type AmyRange
    Index As Long
    SheetName As String
    Adress As String
    Value As Variant
    
    RowN As Long
    ColN As Long
End Type

Sub DataExport()


application.ScreenUpdating = False
application.DisplayAlerts = False
application.Calculation = xlCalculationManual

Dim arrList
Dim i As Integer
If Selection.Cells.Count > 1 Then
    arrList = application.WorksheetFunction.Transpose(Selection)
Else
    ReDim arrList(1): arrList(1) = Selection.Value
End If

Dim TargetBook As String
TargetBook = GetExcelName(ActiveWorkbook.path)

Dim ModelRange() As AmyRange
Dim Ic, Ix, Iy, IT As Long
    Ic = 0: Ix = 100: Iy = 50:  IT = Ix * Iy
Dim Iv, Fc, Fx, Fy, Fv As Long
    Iv = 0
    
ReDim ModelRange(IT) As AmyRange
Dim ModelValue()
    
'确定取数范围
Workbooks.Open Filename:=TargetBook, UpdateLinks:=0, ReadOnly:=True
For Each MySheet In Worksheets
If MySheet.Tab.Color = 255 Then
    Ic = Ic + 1
    For Fx = 1 To Ix
    For Fy = 1 To Iy
        If MySheet.Cells(Fx, Fy).Locked = False Then
            Iv = Iv + 1
            ModelRange(Iv).SheetName = MySheet.name
            ModelRange(Iv).RowN = Fx
            ModelRange(Iv).ColN = Fy
         End If
    Next
    Next
End If
Next
ActiveWorkbook.Close
ReDim ModelValue(Iv)

For i = 1 To UBound(arrList)
Range("StoreNo").Value = arrList(i)
    Calculate
    '取模板值
    For Fv = 1 To Iv
        ModelValue(Fv) = Sheets(ModelRange(Fv).SheetName).Cells(ModelRange(Fv).RowN, ModelRange(Fv).ColN).Value
    Next
    
    '填入模板
    Workbooks.Open Filename:=TargetBook, UpdateLinks:=0, ReadOnly:=True
    For Fv = 1 To Iv
        If ModelValue(Fv) <> "" Then
            Sheets(ModelRange(Fv).SheetName).Cells(ModelRange(Fv).RowN, ModelRange(Fv).ColN).Value = ModelValue(Fv)
        End If
    Next
    ActiveWorkbook.SaveAs arrList(i)
    ActiveWorkbook.Close
Next
 
application.Calculation = xlCalculationSemiautomatic
Beep
End Sub


Sub DataImport()


application.ScreenUpdating = False
application.DisplayAlerts = False
application.Calculation = xlCalculationSemiautomatic

Dim TargetRangeList As String:  TargetRangeList = "'" & ActiveSheet.name & "'!RetrieveCell"
Dim NRow, NCol As Long:         NRow = Range(TargetRangeList).Rows.Count: NCol = Range(TargetRangeList).Columns.Count
Dim TargetBook As String
'    ChDir ThisWorkbook.path
    TargetBook = GetExcelName(ThisWorkbook.path)
    
'Dim arrList
    'arrList = application.WorksheetFunction.Transpose(Range(TargetRangeList))
'Dim i As Integer
'If Selection.Cells.Count > 1 Then
'    arrList = application.WorksheetFunction.Transpose(Selection)
'Else
'    ReDim arrList(1): arrList(1) = Selection.Value
'End If


Dim TargetRange() As AmyRange, i As Long
ReDim TargetRange(NRow)
For i = 1 To NRow
    TargetRange(i).SheetName = Range(TargetRangeList).Cells(i, 1)
    TargetRange(i).Adress = Range(TargetRangeList).Cells(i, 2)
Next

    
'打开目标文件 取值
Workbooks.Open Filename:=TargetBook, UpdateLinks:=0, ReadOnly:=True
    For i = 1 To NRow
        If TargetRange(i).Adress <> "" Then TargetRange(i).Value = Sheets(TargetRange(i).SheetName).Range(TargetRange(i).Adress).Value
    Next
ActiveWorkbook.Close

For i = 1 To NRow
    If TargetRange(i).Adress <> "" Then Range(TargetRangeList).Cells(i, 3).Value = TargetRange(i).Value
Next

Beep
End Sub
