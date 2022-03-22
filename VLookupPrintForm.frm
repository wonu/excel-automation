VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} VLookupPrintForm 
   Caption         =   "VLOOKUP 테이블 인쇄"
   ClientHeight    =   6855
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10920
   OleObjectBlob   =   "VLookupPrintForm.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "VLookupPrintForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const MIN_SERIAL_NUMBER As Integer = 1 ' 유효 최소 연번
Dim DEFAULT_SAVE_PATH As String ' 기본 저장 위치
Dim vlookupTargetSheet As Worksheet ' VLOOKUP 대상 시트

' 폼 초기화
Private Sub UserForm_Initialize()

' 연번 입력 칸에 이전에 입력했던 값 넣기
SerialNumberStartInput = GetCustomDocumentProperty(SerialNumberStartInput.name, MIN_SERIAL_NUMBER)
SerialNumberEndInput = GetCustomDocumentProperty(SerialNumberEndInput.name, MIN_SERIAL_NUMBER)

' 현재 엑셀 파일 위치를 기본 저장 위치로 설정
DEFAULT_SAVE_PATH = ActiveWorkbook.Path & "\"
SavePathInput = DEFAULT_SAVE_PATH

' 시트 선택 목록 상자에 모든 시트명 등록
For Each sheet In ActiveWorkbook.Sheets
    SingleSelectSheetListBox.AddItem sheet.name
    MultiSelectSheetListBox.AddItem sheet.name
Next sheet

End Sub

' CustomDocumentProperties에 값 저장하기
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

' CustomDocumentProperties로부터 값 가져오기
Function GetCustomDocumentProperty(name, Optional defaultValue As Variant)

If Not IsMissing(arg) Then
    GetCustomDocumentProperty = defaultValue
End If

On Error Resume Next
GetCustomDocumentProperty = ActiveWorkbook.CustomDocumentProperties(name)

End Function

' "시작 연번" 값 변경 이벤트 핸들러
Private Sub SerialNumberStartInput_Change()

UpdateCustomDocumentProperty SerialNumberStartInput.name, SerialNumberStartInput.value, msoPropertyTypeString

End Sub

' "끝 연번" 값 변경 이벤트 핸들러
Private Sub SerialNumberEndInput_Change()

UpdateCustomDocumentProperty SerialNumberEndInput.name, SerialNumberEndInput.value, msoPropertyTypeString

End Sub

' "VLOOKUP 대상 시트" 목록 상자 클릭 이벤트 핸들러
Private Sub SingleSelectSheetListBox_Click()

' 목록 상자에서 선택된 시트를 vlookupTargetSheet에 저장
Dim i As Integer
For i = 0 To SingleSelectSheetListBox.ListCount - 1
    If SingleSelectSheetListBox.Selected(i) Then
        Set vlookupTargetSheet = Sheets(SingleSelectSheetListBox.List(i))
    End If
Next i

End Sub

' 저장 위치 변경을 위한 "변경" 버튼 클릭 이벤트 핸들러
Private Sub ChangeSavePathButton_Click()

Dim savePath As String: savePath = GetSavePath
If savePath <> "" Then
    SavePathInput.value = savePath & "\"
End If

End Sub

' 폴더 선택 다이얼로그로부터 저장 위치 받아오기
Function GetSavePath() As String

Dim fileDialog As fileDialog
Dim selectedItem As String

Set fileDialog = Application.fileDialog(msoFileDialogFolderPicker)
With fileDialog
    .Title = "저장 위치 선택"
    .AllowMultiSelect = False
    .InitialFileName = DEFAULT_SAVE_PATH
    If .Show = -1 Then
        selectedItem = .SelectedItems(1)
    End If
End With

' 값 반환
GetSavePath = selectedItem
Set fileDialog = Nothing

End Function

' "PDF로 저장" 버튼 클릭 이벤트 핸들러
Private Sub SaveAsPdfButton_Click()

PrintVLookupSheet True

End Sub

' "인쇄" 버튼 클릭 이벤트 핸들러
Private Sub PrintButton_Click()

PrintVLookupSheet False

End Sub

' 출력 함수, willExport가 True면 PDF로 저장, False면 인쇄
Function PrintVLookupSheet(willExport As Boolean)

' 연번으로 입력된 값의 유효성 확인
If CheckSerialNumbersToPrint() = False Then
    Exit Function
End If

If vlookupTargetSheet Is Nothing Then
    MsgBox "VLOOKUP의 대상 시트를 선택해주세요."
    Exit Function
End If

If Not IsInteger(EmployerNameColumnNumberInput) Then
    MsgBox "직원명이 담긴 열 번호를 입력해주세요."
    Exit Function
End If

' ArrayList에 VLOOKUP을 사용하는 시트들 담기
Set selectedSheetNames = CreateObject("System.Collections.ArrayList")
Dim i As Integer
For i = 0 To MultiSelectSheetListBox.ListCount - 1
    If MultiSelectSheetListBox.Selected(i) Then
        selectedSheetNames.Add MultiSelectSheetListBox.List(i)
    End If
Next i

If selectedSheetNames.Count = 0 Then
    MsgBox "출력할 시트를 선택해주세요."
    Exit Function
End If

Dim canceledSerialNumber As Integer: canceledSerialNumber = -1 ' 작업을 완료하지 못한 연번
Dim printed As Boolean: printed = False ' 출력 수행 여부 확인용 변수

' 목록 상자에서 선택된 각 시트에 대해 반복
For Each sheetName In selectedSheetNames
    Set vlookupSheet = Sheets(sheetName)
    
    Set vlookupCells = CreateObject("System.Collections.ArrayList") ' 해당 시트 내에서 VLOOKUP을 사용하는 모든 셀을 담는 ArrayList
    Set prevFormulas = CreateObject("System.Collections.ArrayList") ' 수식의 원상복구를 위해 기존 수식을 담아두는 ArrayList
    Set lookupValuePositions = CreateObject("System.Collections.ArrayList") ' 수식 내에서 변경할 위치를 담아두는 ArrayList
    
    Dim pos As Integer ' 수식 내에서 변경할 위치를 찾기 위한 변수 선언
    
    ' 수식을 사용하는 각 셀에 대해 반복
    For Each c In vlookupSheet.UsedRange.SpecialCells(xlCellTypeFormulas)
        pos = InStr(1, c.Formula, "VLOOKUP", vbTextCompare) ' 수식 내에서 "VLOOKUP"이 등장하는 위치 저장
        If pos Then ' 수식에서 "VLOOKUP"이 등장하면
            vlookupCells.Add c ' 해당 셀 저장
            prevFormulas.Add c.Formula ' 기존 수식 저장
            lookupValuePositions.Add pos + 7 ' Len("LOOKUP(") ' 수식 변경 위치 저장
        End If
    Next c
    
    Dim j As Integer ' 반복문을 위한 변수 선언
    
    ' 사용자가 입력한 시작 연번부터 끝 연번까지 반복
    For serialNumber = GetCustomDocumentProperty(SerialNumberStartInput.name) To GetCustomDocumentProperty(SerialNumberEndInput.name)
        ' 해당 시트 내의 모든 VLOOKUP 함수식 변경 (VLOOKUP이 2개 이상인 경우는 사용 불가)
        For j = 0 To vlookupCells.Count - 1
            Set c = vlookupCells(j)
            pos = lookupValuePositions(j)
            
            ' 예를 들어 c.Formula = "=VLOOKUP(AC4, 기초정보!A:BI, 2, FALSE)"인 경우 "AC4"를 현재 연번으로 변경
            ' Left(...) = "=VLOOKUP("
            ' Mid(...) = " 기초정보!A:BI, 2, FALSE)"
            c.Formula = Left(c.Formula, pos) & serialNumber & Mid(c.Formula, InStr(pos, c.Formula, ",", vbTextCompare))
        Next j
                    
        If willExport Then ' PDF로 저장하기
            ' 사용자가 입력한 직원명 열 번호와 VLOOKUP 함수를 이용하여 직원명 가져오기
            Dim employerName As Variant
            employerName = Application.VLookup( _
                serialNumber, _
                vlookupTargetSheet.Range("A:" & GetLastColumnLetter(vlookupTargetSheet)), _
                EmployerNameColumnNumberInput.value, _
                False _
            )
            
            ' 직원명이 존재하지 않는 경우 이후 연번에 대한 작업 취소
            If IsError(employerName) Then
                canceledSerialNumber = serialNumber
                Exit For
            End If
            
            ' 확장자를 제외한 파일명
            Dim stem As String: stem = GetStem(ActiveWorkbook.name) & "_" & sheetName & "_" & employerName
            
            ' 저장 위치, 시트명, stem 매개변수로 ExportAsPdf 호출
            ExportAsPdf SavePathInput.value, CStr(sheetName), stem
        Else: vlookupSheet.PrintOut Copies:=1, Collate:=True, IgnorePrintAreas:=False ' 인쇄
        End If
        
        printed = True
    Next serialNumber
    
    ' 수식 원상복구
    For j = 0 To vlookupCells.Count - 1
        Set c = vlookupCells(j)
        c.Formula = prevFormulas(j)
    Next j
Next sheetName

If canceledSerialNumber <> -1 Then
    MsgBox "연번 " & serialNumber & "부터는 작업이 취소되었습니다."
    Exit Function
End If

If printed Then
    MsgBox "작업을 완료하였습니다."
Else: MsgBox "작업을 완료하지 못했습니다."
End If

End Function

' 사용자가 입력한 연번이 유효한지 확인하기
Function CheckSerialNumbersToPrint() As Boolean

' 정수가 아니거나, 유효 최소 연번보다 작으면 False 반환
If Not IsInteger(SerialNumberStartInput.value) _
    Or Not IsInteger(SerialNumberEndInput.value) _
    Or SerialNumberStartInput.value < MIN_SERIAL_NUMBER _
    Or SerialNumberEndInput.value < MIN_SERIAL_NUMBER Then
    MsgBox "연번은 " & MIN_SERIAL_NUMBER & " 이상의 정수를 입력해주세요."
    CheckSerialNumbersToPrint = False
    Exit Function
End If

If Int(SerialNumberStartInput.value) > Int(SerialNumberEndInput.value) Then
    MsgBox "시작 연번은 끝 연번 이하여야 합니다."
    CheckSerialNumbersToPrint = False
ElseIf Int(SerialNumberEndInput.value) - Int(SerialNumberStartInput.value) > 50 Then
    MsgBox "한 번에 처리 가능한 최대 연번 개수는 50개입니다."
    CheckSerialNumbersToPrint = False
Else: CheckSerialNumbersToPrint = True
End If

End Function

' 주어진 값이 정수인지 확인하기
Function IsInteger(value) As Boolean

' numeric이 아니면 False 반환
If Not IsNumeric(value) Then
    IsInteger = False
    Exit Function
End If

value = CDbl(value)
If Not value = Int(value) Then ' Int(정수)가 아니면 False 반환
    IsInteger = False
Else: IsInteger = True ' True 반환
End If

End Function

' 주어진 시트 내에서 값이 들어있는 마지막 열 문자(예: "AC") 반환
Function GetLastColumnLetter(sheet As Worksheet)

' 마지막 셀
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

' 값 반환
GetLastColumnLetter = GetColumnLetter(lastCell)

End Function

' 주어진 셀의 열 문자(예: "AC") 반환
Function GetColumnLetter(cell As Range)

' 값 반환
GetColumnLetter = Split(cell.Address(True, False), "$")(0)

End Function

' 폴더명, 시트명, 파일명(확장자 제외)을 이용하여 PDF로 저장하기
Function ExportAsPdf(dirname As String, sheetName As String, stem As String)

On Error GoTo errHandler ' 에러가 발생하는 경우 errHandler로 이동

' 엑셀 파일명으로 폴더 생성
Dim workbookDirname As String
workbookDirname = dirname & GetStem(ActiveWorkbook.name) & "\"
GetReadyDir (workbookDirname)

' 위의 폴더 안에 시트명으로 폴더 생성
Dim sheetDirname As String
sheetDirname = workbookDirname & sheetName & "\"
GetReadyDir (sheetDirname)

' 위의 경로 및 파일명을 조합한 최종 경로 생성
Dim fullPath As String
fullPath = sheetDirname & GetAvailableFilename(sheetDirname, stem, ".pdf")

' 해당 시트를 PDF로 내보내기
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
    MsgBox "PDF 생성 실패: " & fullPath
    Resume exitHandler
    
End Function

' 주어진 폴더명의 폴더가 존재하지 않는 경우 생성
Function GetReadyDir(dirname As String)

If Dir(dirname, vbDirectory) = "" Then
    MkDir (dirname)
End If

End Function

' 폴더명, 파일명(확장자 제외), 확장자를 이용하여 저장 가능한 파일명 찾기
Function GetAvailableFilename(dirname As String, stem As String, ext As String) As String
' https://stackoverflow.com/a/27220620

Dim testFilename As String: testFilename = "" ' 사용 가능한 파일명인지 확인하기 위한 테스트 파일명
Dim fileCounter As Integer: fileCounter = 1 ' 동일한 파일명을 피하기 위한 파일 카운터. "abc.pdf", "abc (2).pdf", ...

' 사용 가능한 파일명을 찾을 때까지 반복
Do While True
    ' 파일 카운터를 이용하여 테스트 파일명 생성
    testFilename = stem & IIf(fileCounter > 1, " (" & Trim(Str(fileCounter)) & ")", "") & ext

    ' 파일명이 사용 가능한 경우
    If (Dir(dirname & testFilename) = "") Then
        Exit Do ' 반복문 탈출
    End If
    
    fileCounter = fileCounter + 1 ' 카운터 증가
Loop

' 값 반환
GetAvailableFilename = testFilename

End Function

' 확장자를 제외한 파일명 가져오기
Function GetStem(basename As String)

' 값 반환
GetStem = Left(basename, InStrRev(basename, ".") - 1)

End Function

' "닫기" 버튼 클릭 이벤트 핸들러
Private Sub CloseButton_Click()

' 폼 닫기
Unload Me

End Sub
