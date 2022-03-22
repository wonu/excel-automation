VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} VLookupPrintForm 
   Caption         =   "VLOOKUP ���̺� �μ�"
   ClientHeight    =   6855
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10920
   OleObjectBlob   =   "VLookupPrintForm.frx":0000
   StartUpPosition =   1  '������ ���
End
Attribute VB_Name = "VLookupPrintForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const MIN_SERIAL_NUMBER As Integer = 1 ' ��ȿ �ּ� ����
Dim DEFAULT_SAVE_PATH As String ' �⺻ ���� ��ġ
Dim vlookupTargetSheet As Worksheet ' VLOOKUP ��� ��Ʈ

' �� �ʱ�ȭ
Private Sub UserForm_Initialize()

' ���� �Է� ĭ�� ������ �Է��ߴ� �� �ֱ�
SerialNumberStartInput = GetCustomDocumentProperty(SerialNumberStartInput.name, MIN_SERIAL_NUMBER)
SerialNumberEndInput = GetCustomDocumentProperty(SerialNumberEndInput.name, MIN_SERIAL_NUMBER)

' ���� ���� ���� ��ġ�� �⺻ ���� ��ġ�� ����
DEFAULT_SAVE_PATH = ActiveWorkbook.Path & "\"
SavePathInput = DEFAULT_SAVE_PATH

' ��Ʈ ���� ��� ���ڿ� ��� ��Ʈ�� ���
For Each sheet In ActiveWorkbook.Sheets
    SingleSelectSheetListBox.AddItem sheet.name
    MultiSelectSheetListBox.AddItem sheet.name
Next sheet

End Sub

' CustomDocumentProperties�� �� �����ϱ�
Function UpdateCustomDocumentProperty(name As String, value, docType As Office.MsoDocProperties)

On Error Resume Next
ActiveWorkbook.CustomDocumentProperties(name).value = value
If Err.Number > 0 Then
    ActiveWorkbook.CustomDocumentProperties.Add _
        name:=name, _
        LinkToContent:=False, _
        Type:=docType, _
        value:=value
End If
    
End Function

' CustomDocumentProperties�κ��� �� ��������
Function GetCustomDocumentProperty(name, Optional defaultValue As Variant)

If Not IsMissing(arg) Then
    GetCustomDocumentProperty = defaultValue
End If

On Error Resume Next
GetCustomDocumentProperty = ActiveWorkbook.CustomDocumentProperties(name)

End Function

' "���� ����" �� ���� �̺�Ʈ �ڵ鷯
Private Sub SerialNumberStartInput_Change()

UpdateCustomDocumentProperty SerialNumberStartInput.name, SerialNumberStartInput.value, msoPropertyTypeString

End Sub

' "�� ����" �� ���� �̺�Ʈ �ڵ鷯
Private Sub SerialNumberEndInput_Change()

UpdateCustomDocumentProperty SerialNumberEndInput.name, SerialNumberEndInput.value, msoPropertyTypeString

End Sub

' "VLOOKUP ��� ��Ʈ" ��� ���� Ŭ�� �̺�Ʈ �ڵ鷯
Private Sub SingleSelectSheetListBox_Click()

' ��� ���ڿ��� ���õ� ��Ʈ�� vlookupTargetSheet�� ����
Dim i As Integer
For i = 0 To SingleSelectSheetListBox.ListCount - 1
    If SingleSelectSheetListBox.Selected(i) Then
        Set vlookupTargetSheet = Sheets(SingleSelectSheetListBox.List(i))
    End If
Next i

End Sub

' ���� ��ġ ������ ���� "����" ��ư Ŭ�� �̺�Ʈ �ڵ鷯
Private Sub ChangeSavePathButton_Click()

Dim savePath As String: savePath = GetSavePath
If savePath <> "" Then
    SavePathInput.value = savePath & "\"
End If

End Sub

' ���� ���� ���̾�α׷κ��� ���� ��ġ �޾ƿ���
Function GetSavePath() As String

Dim fileDialog As fileDialog
Dim selectedItem As String

Set fileDialog = Application.fileDialog(msoFileDialogFolderPicker)
With fileDialog
    .Title = "���� ��ġ ����"
    .AllowMultiSelect = False
    .InitialFileName = DEFAULT_SAVE_PATH
    If .Show = -1 Then
        selectedItem = .SelectedItems(1)
    End If
End With

' �� ��ȯ
GetSavePath = selectedItem
Set fileDialog = Nothing

End Function

' "PDF�� ����" ��ư Ŭ�� �̺�Ʈ �ڵ鷯
Private Sub SaveAsPdfButton_Click()

PrintVLookupSheet True

End Sub

' "�μ�" ��ư Ŭ�� �̺�Ʈ �ڵ鷯
Private Sub PrintButton_Click()

PrintVLookupSheet False

End Sub

' ��� �Լ�, willExport�� True�� PDF�� ����, False�� �μ�
Function PrintVLookupSheet(willExport As Boolean)

' �������� �Էµ� ���� ��ȿ�� Ȯ��
If CheckSerialNumbersToPrint() = False Then
    Exit Function
End If

If vlookupTargetSheet Is Nothing Then
    MsgBox "VLOOKUP�� ��� ��Ʈ�� �������ּ���."
    Exit Function
End If

If Not IsInteger(EmployerNameColumnNumberInput) Then
    MsgBox "�������� ��� �� ��ȣ�� �Է����ּ���."
    Exit Function
End If

' ArrayList�� VLOOKUP�� ����ϴ� ��Ʈ�� ���
Set selectedSheetNames = CreateObject("System.Collections.ArrayList")
Dim i As Integer
For i = 0 To MultiSelectSheetListBox.ListCount - 1
    If MultiSelectSheetListBox.Selected(i) Then
        selectedSheetNames.Add MultiSelectSheetListBox.List(i)
    End If
Next i

If selectedSheetNames.Count = 0 Then
    MsgBox "����� ��Ʈ�� �������ּ���."
    Exit Function
End If

Dim canceledSerialNumber As Integer: canceledSerialNumber = -1 ' �۾��� �Ϸ����� ���� ����
Dim printed As Boolean: printed = False ' ��� ���� ���� Ȯ�ο� ����

' ��� ���ڿ��� ���õ� �� ��Ʈ�� ���� �ݺ�
For Each sheetName In selectedSheetNames
    Set vlookupSheet = Sheets(sheetName)
    
    Set vlookupCells = CreateObject("System.Collections.ArrayList") ' �ش� ��Ʈ ������ VLOOKUP�� ����ϴ� ��� ���� ��� ArrayList
    Set prevFormulas = CreateObject("System.Collections.ArrayList") ' ������ ���󺹱��� ���� ���� ������ ��Ƶδ� ArrayList
    Set lookupValuePositions = CreateObject("System.Collections.ArrayList") ' ���� ������ ������ ��ġ�� ��Ƶδ� ArrayList
    
    Dim pos As Integer ' ���� ������ ������ ��ġ�� ã�� ���� ���� ����
    
    ' ������ ����ϴ� �� ���� ���� �ݺ�
    For Each c In vlookupSheet.UsedRange.SpecialCells(xlCellTypeFormulas)
        pos = InStr(1, c.Formula, "VLOOKUP", vbTextCompare) ' ���� ������ "VLOOKUP"�� �����ϴ� ��ġ ����
        If pos Then ' ���Ŀ��� "VLOOKUP"�� �����ϸ�
            vlookupCells.Add c ' �ش� �� ����
            prevFormulas.Add c.Formula ' ���� ���� ����
            lookupValuePositions.Add pos + 7 ' Len("LOOKUP(") ' ���� ���� ��ġ ����
        End If
    Next c
    
    Dim j As Integer ' �ݺ����� ���� ���� ����
    
    ' ����ڰ� �Է��� ���� �������� �� �������� �ݺ�
    For serialNumber = GetCustomDocumentProperty(SerialNumberStartInput.name) To GetCustomDocumentProperty(SerialNumberEndInput.name)
        ' �ش� ��Ʈ ���� ��� VLOOKUP �Լ��� ���� (VLOOKUP�� 2�� �̻��� ���� ��� �Ұ�)
        For j = 0 To vlookupCells.Count - 1
            Set c = vlookupCells(j)
            pos = lookupValuePositions(j)
            
            ' ���� ��� c.Formula = "=VLOOKUP(AC4, ��������!A:BI, 2, FALSE)"�� ��� "AC4"�� ���� �������� ����
            ' Left(...) = "=VLOOKUP("
            ' Mid(...) = " ��������!A:BI, 2, FALSE)"
            c.Formula = Left(c.Formula, pos) & serialNumber & Mid(c.Formula, InStr(pos, c.Formula, ",", vbTextCompare))
        Next j
                    
        If willExport Then ' PDF�� �����ϱ�
            ' ����ڰ� �Է��� ������ �� ��ȣ�� VLOOKUP �Լ��� �̿��Ͽ� ������ ��������
            Dim employerName As Variant
            employerName = Application.VLookup( _
                serialNumber, _
                vlookupTargetSheet.Range("A:" & GetLastColumnLetter(vlookupTargetSheet)), _
                EmployerNameColumnNumberInput.value, _
                False _
            )
            
            ' �������� �������� �ʴ� ��� ���� ������ ���� �۾� ���
            If IsError(employerName) Then
                canceledSerialNumber = serialNumber
                Exit For
            End If
            
            ' Ȯ���ڸ� ������ ���ϸ�
            Dim stem As String: stem = GetStem(ActiveWorkbook.name) & "_" & sheetName & "_" & employerName
            
            ' ���� ��ġ, ��Ʈ��, stem �Ű������� ExportAsPdf ȣ��
            ExportAsPdf SavePathInput.value, CStr(sheetName), stem
        Else: vlookupSheet.PrintOut Copies:=1, Collate:=True, IgnorePrintAreas:=False ' �μ�
        End If
        
        printed = True
    Next serialNumber
    
    ' ���� ���󺹱�
    For j = 0 To vlookupCells.Count - 1
        Set c = vlookupCells(j)
        c.Formula = prevFormulas(j)
    Next j
Next sheetName

If canceledSerialNumber <> -1 Then
    MsgBox "���� " & serialNumber & "���ʹ� �۾��� ��ҵǾ����ϴ�."
    Exit Function
End If

If printed Then
    MsgBox "�۾��� �Ϸ��Ͽ����ϴ�."
Else: MsgBox "�۾��� �Ϸ����� ���߽��ϴ�."
End If

End Function

' ����ڰ� �Է��� ������ ��ȿ���� Ȯ���ϱ�
Function CheckSerialNumbersToPrint() As Boolean

' ������ �ƴϰų�, ��ȿ �ּ� �������� ������ False ��ȯ
If Not IsInteger(SerialNumberStartInput.value) _
    Or Not IsInteger(SerialNumberEndInput.value) _
    Or SerialNumberStartInput.value < MIN_SERIAL_NUMBER _
    Or SerialNumberEndInput.value < MIN_SERIAL_NUMBER Then
    MsgBox "������ " & MIN_SERIAL_NUMBER & " �̻��� ������ �Է����ּ���."
    CheckSerialNumbersToPrint = False
    Exit Function
End If

If Int(SerialNumberStartInput.value) > Int(SerialNumberEndInput.value) Then
    MsgBox "���� ������ �� ���� ���Ͽ��� �մϴ�."
    CheckSerialNumbersToPrint = False
ElseIf Int(SerialNumberEndInput.value) - Int(SerialNumberStartInput.value) > 50 Then
    MsgBox "�� ���� ó�� ������ �ִ� ���� ������ 50���Դϴ�."
    CheckSerialNumbersToPrint = False
Else: CheckSerialNumbersToPrint = True
End If

End Function

' �־��� ���� �������� Ȯ���ϱ�
Function IsInteger(value) As Boolean

' numeric�� �ƴϸ� False ��ȯ
If Not IsNumeric(value) Then
    IsInteger = False
    Exit Function
End If

value = CDbl(value)
If Not value = Int(value) Then ' Int(����)�� �ƴϸ� False ��ȯ
    IsInteger = False
Else: IsInteger = True ' True ��ȯ
End If

End Function

' �־��� ��Ʈ ������ ���� ����ִ� ������ �� ����(��: "AC") ��ȯ
Function GetLastColumnLetter(sheet As Worksheet)

' ������ ��
Dim lastCell As Range
Set lastCell = sheet.Cells.Find( _
    What:="*", _
    After:=sheet.Cells(1, 1), _
    LookIn:=xlFormulas, _
    LookAt:=xlPart, _
    SearchOrder:=xlByColumns, _
    SearchDirection:=xlPrevious, _
    MatchCase:=False _
)

' �� ��ȯ
GetLastColumnLetter = GetColumnLetter(lastCell)

End Function

' �־��� ���� �� ����(��: "AC") ��ȯ
Function GetColumnLetter(cell As Range)

' �� ��ȯ
GetColumnLetter = Split(cell.Address(True, False), "$")(0)

End Function

' ������, ��Ʈ��, ���ϸ�(Ȯ���� ����)�� �̿��Ͽ� PDF�� �����ϱ�
Function ExportAsPdf(dirname As String, sheetName As String, stem As String)

On Error GoTo errHandler ' ������ �߻��ϴ� ��� errHandler�� �̵�

' ���� ���ϸ����� ���� ����
Dim workbookDirname As String
workbookDirname = dirname & GetStem(ActiveWorkbook.name) & "\"
GetReadyDir (workbookDirname)

' ���� ���� �ȿ� ��Ʈ������ ���� ����
Dim sheetDirname As String
sheetDirname = workbookDirname & sheetName & "\"
GetReadyDir (sheetDirname)

' ���� ��� �� ���ϸ��� ������ ���� ��� ����
Dim fullPath As String
fullPath = sheetDirname & GetAvailableFilename(sheetDirname, stem, ".pdf")

' �ش� ��Ʈ�� PDF�� ��������
Sheets(sheetName).ExportAsFixedFormat _
    Type:=xlTypePDF, _
    filename:=fullPath, _
    Quality:=xlQualityStandard, _
    IncludeDocProperties:=True, _
    IgnorePrintAreas:=False, _
    OpenAfterPublish:=False

exitHandler:
    Exit Function
errHandler:
    MsgBox "PDF ���� ����: " & fullPath
    Resume exitHandler
    
End Function

' �־��� �������� ������ �������� �ʴ� ��� ����
Function GetReadyDir(dirname As String)

If Dir(dirname, vbDirectory) = "" Then
    MkDir (dirname)
End If

End Function

' ������, ���ϸ�(Ȯ���� ����), Ȯ���ڸ� �̿��Ͽ� ���� ������ ���ϸ� ã��
Function GetAvailableFilename(dirname As String, stem As String, ext As String) As String
' https://stackoverflow.com/a/27220620

Dim testFilename As String: testFilename = "" ' ��� ������ ���ϸ����� Ȯ���ϱ� ���� �׽�Ʈ ���ϸ�
Dim fileCounter As Integer: fileCounter = 1 ' ������ ���ϸ��� ���ϱ� ���� ���� ī����. "abc.pdf", "abc (2).pdf", ...

' ��� ������ ���ϸ��� ã�� ������ �ݺ�
Do While True
    ' ���� ī���͸� �̿��Ͽ� �׽�Ʈ ���ϸ� ����
    testFilename = stem & IIf(fileCounter > 1, " (" & Trim(Str(fileCounter)) & ")", "") & ext

    ' ���ϸ��� ��� ������ ���
    If (Dir(dirname & testFilename) = "") Then
        Exit Do ' �ݺ��� Ż��
    End If
    
    fileCounter = fileCounter + 1 ' ī���� ����
Loop

' �� ��ȯ
GetAvailableFilename = testFilename

End Function

' Ȯ���ڸ� ������ ���ϸ� ��������
Function GetStem(basename As String)

' �� ��ȯ
GetStem = Left(basename, InStrRev(basename, ".") - 1)

End Function

' "�ݱ�" ��ư Ŭ�� �̺�Ʈ �ڵ鷯
Private Sub CloseButton_Click()

' �� �ݱ�
Unload Me

End Sub
