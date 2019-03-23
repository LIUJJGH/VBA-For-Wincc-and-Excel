Attribute VB_Name = "模块1"
Sub ImportObjectListFromXLSX_参数表_IOField()
'VBA75
Dim objGDApplication As grafexe.Application
Dim objDoc As grafexe.Document
Dim objHMIObject As grafexe.HMIObject
Dim objXLS As Excel.Application
Dim objWSheet As Excel.Worksheet
Dim objWBook As Excel.Workbook
Dim lRow As Long
Dim strWorkbookName As String
Dim strWorksheetName As String
Dim strSheets As String

Dim objVariableTrigger As HMIVariableTrigger    '动态化变量
Dim objCScript As HMIScriptInfo                 '脚本操作对象
Dim strCode As String

'---------------------- 动作中插入 c脚本 -------------------------
strCode = "#pragma option(mbcs)"
strCode = strCode & vbCrLf & "char *linkvar,szBuffer[50];"
strCode = strCode & vbCrLf & "float newValue,oldValue;"
strCode = strCode & vbCrLf & "int xRet;"
strCode = strCode & vbCrLf & "if ( nChar==13 )" & "{" & vbTab & vbTab & "//if press enter"
strCode = strCode & vbCrLf & vbTab & "linkvar=GetLinkedVariable(lpszPictureName,lpszObjectName,""OutputValue"");"
strCode = strCode & vbCrLf & vbTab & "oldValue=GetTagFloat(linkvar);  //旧值"
strCode = strCode & vbCrLf & vbTab & "newValue=GetInputValueDouble(lpszPictureName,lpszObjectName);//新值"
strCode = strCode & vbCrLf & vbTab & "sprintf( szBuffer,""Input number:%8.3f"",newValue);"
strCode = strCode & vbCrLf & vbTab & "xRet = MessageBox(NULL,szBuffer,""确认窗口"",MB_YESNO|MB_ICONQUESTION|MB_SYSTEMMODAL);//message"
strCode = strCode & vbCrLf & vbTab & "if( xRet == IDYES )" & "{" & "//confirm operate"
strCode = strCode & vbCrLf & vbTab & vbTab & "SetTagFloat(linkvar,newValue);// set data  赋值"
strCode = strCode & vbCrLf & vbTab & vbTab & "ISALG_OperationLog(lpszObjectName,"""",""更改操作"","""",""OK"",oldValue,newValue,""备注"");"
strCode = strCode & vbCrLf & vbTab & "}"
strCode = strCode & vbCrLf & "}"

'define local errorhandler
On Local Error GoTo LocErrTrap
 
'Set references on the applications Excel and GraphicsDesigner
Set objGDApplication = Application
Set objDoc = objGDApplication.ActiveDocument
Set objXLS = New Excel.Application
 
'Open workbook. The workbook have to be in datapath of GraphicsDesigner
'strWorkbookName = InputBox("Name of workbook:", "Import of objects")
Set objWBook = objXLS.Workbooks.Open(objGDApplication.ApplicationDataPath & "..\" & "Import_IOField.xlsx")
On Local Error GoTo LocErrTrap

If objWBook Is Nothing Then
MsgBox "Open workbook fails!" & vbCrLf & "This function is cancled!", vbCritical, "Import od objects"
Set objDoc = Nothing
Set objGDApplication = Nothing
Set objXLS = Nothing
Exit Sub
End If
 
'Read out the names of all worksheets contained in the workbook
For Each objWSheet In objWBook.Sheets
strSheets = strSheets & objWSheet.Name & vbCrLf
Next objWSheet
strWorksheetName = InputBox("Name of table to import:" & vbCrLf & strSheets, "Import of objects")
Set objWSheet = objWBook.Sheets(strWorksheetName)
lRow = 3
 
'Import the worksheet as long as in actual row the first column is empty.
'Add with the outreaded data new objects to the active document and
'assign the values to the objectproperties
With objWSheet
While (.Cells(lRow, 1).value <> vbNullString) And (Not IsEmpty(.Cells(lRow, 1).value))
'Add the objects to the document as its objecttype,
'do nothing by groups, their have to create before.
If (UCase(.Cells(lRow, 2).value) = "HMIGROUP") Then
Else
  If (UCase(.Cells(lRow, 2).value) = "HMIACTIVEXCONTROL") Then
    Set objHMIObject = objDoc.HMIObjects.AddActiveXControl(.Cells(lRow, 1).value, .Cells(lRow, 3).value)
  Else
    Set objHMIObject = objDoc.HMIObjects.AddHMIObject(.Cells(lRow, 1).value, .Cells(lRow, 2).value)
    
    '设置和连接动态变量
    Set objVariableTrigger = objHMIObject.OutputValue.CreateDynamic(hmiDynamicCreationTypeVariableDirect, .Cells(lRow, 9).value)
        objVariableTrigger.CycleType = hmiVariableCycleType_有变化时
  End If
  '创建C动作
  Set objCScript = objHMIObject.Events(7).Actions.AddAction(hmiActionCreationTypeCScript)
  objCScript.SourceCode = strCode
  
'----------------从单元格中获取属性值
  objHMIObject.Left = .Cells(lRow, 4).value
  objHMIObject.Top = .Cells(lRow, 5).value
  objHMIObject.Width = .Cells(lRow, 6).value
  objHMIObject.Height = .Cells(lRow, 7).value
  objHMIObject.Layer = .Cells(lRow, 8).value
  '输出值，数据格式，输出格式
  'objHMIObject.OutputValue = .Cells(lRow, 9).value
  objHMIObject.DataFormat = .Cells(lRow, 10).value
  objHMIObject.OutputFormat = .Cells(lRow, 11).value
  
  '背景颜色，字体，字体大小，x轴对齐,y轴对齐
  objHMIObject.BackColor = .Cells(lRow, 12).value
  objHMIObject.FONTNAME = .Cells(lRow, 13).value
  objHMIObject.FONTSIZE = .Cells(lRow, 14).value
  objHMIObject.AlignmentLeft = .Cells(lRow, 15).value
  objHMIObject.AlignmentTop = .Cells(lRow, 16).value
  objHMIObject.BoxType = .Cells(lRow, 17).value  '0输出域 1输入域 2 输入输出域
  objHMIObject.GlobalColorScheme = False '全局颜色方案设为否
End If
Set objHMIObject = Nothing
lRow = lRow + 1
Wend
End With
objWBook.Close
Set objWBook = Nothing
objXLS.Quit
Set objXLS = Nothing
Set objDoc = Nothing
Set objGDApplication = Nothing
MsgBox "导入完成!"
Exit Sub
LocErrTrap:
MsgBox Err.Description, , Err.Source
Resume Next
End Sub






Sub ExportObjectListToXLS_机运画面窗口()
'VBA74
Dim objGDApplication As grafexe.Application
Dim objDoc As grafexe.Document
Dim objHMIObject As grafexe.HMIObject
Dim objProperty As grafexe.HMIProperty
Dim objXLS As Excel.Application
Dim objWSheet As Excel.Worksheet
Dim objWBook As Excel.Workbook
Dim lRow As Long
 
'Define local errorhandler
On Local Error GoTo LocErrTrap
 
'Set references on the applications Excel and GraphicsDesigner
Set objGDApplication = grafexe.Application
Set objDoc = objGDApplication.ActiveDocument
Set objXLS = New Excel.Application
 
'Create workbook
Set objWBook = objXLS.Workbooks.Add()
objWBook.SaveAs objGDApplication.ApplicationDataPath & "..\" & "CC11.xlsx"
 
'Create worksheet in the new workbook and write headline
'The name of the worksheet is equivalent to the documents name
Set objWSheet = objWBook.Worksheets.Add
objWSheet.Name = objDoc.Name
objWSheet.Cells(1, 1) = "Objectname"
objWSheet.Cells(1, 2) = "Objekttyp"
objWSheet.Cells(1, 3) = "ProgID"
objWSheet.Cells(1, 4) = "Position X"
objWSheet.Cells(1, 5) = "Position Y"
objWSheet.Cells(1, 6) = "Width"
objWSheet.Cells(1, 7) = "Height"
objWSheet.Cells(1, 8) = "Ebene"

objWSheet.Cells(1, 9) = "PictureName"
objWSheet.Cells(1, 10) = "TagPreFix"
objWSheet.Cells(1, 11) = "CaptionText"

lRow = 3
 
'Every objects will be written with their objectproperties width,
'height, pos x, pos y and layer to Excel. If the object is an
'ActiveX-Control the ProgID will be also exported.
For Each objHMIObject In objDoc.HMIObjects
DoEvents
If UCase(objHMIObject.Type) = "HMIPICTUREWINDOW" Then
objWSheet.Cells(lRow, 1).value = objHMIObject.ObjectName
objWSheet.Cells(lRow, 2).value = objHMIObject.Type

If UCase(objHMIObject.Type) = "HMIACTIVEXCONTROL" Then
    objWSheet.Cells(lRow, 3).value = objHMIObject.ProgID
End If

objWSheet.Cells(lRow, 4).value = objHMIObject.Left
objWSheet.Cells(lRow, 5).value = objHMIObject.Top
objWSheet.Cells(lRow, 6).value = objHMIObject.Width
objWSheet.Cells(lRow, 7).value = objHMIObject.Height
objWSheet.Cells(lRow, 8).value = objHMIObject.Layer
objWSheet.Cells(lRow, 9).value = objHMIObject.PictureName
objWSheet.Cells(lRow, 10).value = objHMIObject.TagPrefix
objWSheet.Cells(lRow, 11).value = objHMIObject.CaptionText
lRow = lRow + 1
End If

Next objHMIObject
objWSheet.Columns.AutoFit
Set objWSheet = Nothing
objWBook.Save
objWBook.Close
Set objWBook = Nothing
objXLS.Quit
Set objXLS = Nothing
Set objDoc = Nothing
Set objGDApplication = Nothing
Exit Sub
MsgBox "导出完成"
LocErrTrap:
MsgBox Err.Description, , Err.Source
Resume Next
End Sub
Sub ImportListFromXLSX_机运画面窗口()
Dim objGDApplication As grafexe.Application
Dim objDoc As grafexe.Document
Dim objHMIObject As grafexe.HMIObject
Dim objXLS As Excel.Application
Dim objWSheet As Excel.Worksheet
Dim objWBook As Excel.Workbook
Dim lRow As Long
Dim strWorkbookName As String
Dim strWorksheetName As String
Dim strSheets As String

On Local Error GoTo LocErrTrap
Set objGDApplication = Application
Set objDoc = objGDApplication.ActiveDocument

Set objXLS = New Excel.Application

'strWorkbookName = InputBox("Name of workbook:", "Import of objects")
Set objWBook = objXLS.Workbooks.Open(objGDApplication.ApplicationDataPath & "..\" & "Import_Data.xlsx")

If objWBook Is Nothing Then
    MsgBox "Open workbook fails!" & vbCrLf & "This function is cancled!", vbCritical, "Import od objects"
    Set objDoc = Nothing
    Set objXLS = Nothing
    Set objWBook = Nothing
    Exit Sub
End If

For Each objWSheet In objWBook.Sheets
    strSheets = strSheets & objWSheet.Name & vbCrLf
Next objWSheet

strWorksheetName = InputBox("Name of table to import:" & vbCrLf & strSheets, "Import of objects")
Set objWSheet = objWBook.Sheets(strWorksheetName)
lRow = 2

With objWSheet
    While (.Cells(lRow, 1).value <> vbNullString) And (Not IsEmpty(.Cells(lRow, 1).value))
    Set objHMIObject = ActiveDocument.HMIObjects(.Cells(lRow, 1).value)
    '设置属性值
        
        
        Select Case strWorksheetName
        Case "机运"
            objHMIObject.PictureName = .Cells(lRow, 2).value
            objHMIObject.TagPrefix = .Cells(lRow, 3).value
            objHMIObject.CaptionText = .Cells(lRow, 4).value
        Case "设备"
            objHMIObject.PictureName = .Cells(lRow, 2).value
            objHMIObject.TagPrefix = .Cells(lRow, 3).value
        Case Else
            MsgBox "FielName Error!"
        End Select
        
        lRow = lRow + 1
    Wend
End With
MsgBox "导入完成"
Exit Sub
LocErrTrap:
MsgBox Err.Description, , Err.Source
End Sub



