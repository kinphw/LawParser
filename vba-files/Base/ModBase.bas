Attribute VB_Name = "ModBase"
'namespace=vba-files\Base\
Option Explicit

'개행문자 2개를 1개로 치환
Sub ReplaceCRLF2to1(control As IRibbonControl)

    Debug.Print "연속개행문자를 1개로 치환합니다.";

    Selection.Replace What:="" & Chr(10) & "" & Chr(10) & "", Replacement:="" & Chr(10) & "", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2


    Debug.Print "공백+개행문자를 1개로 치환합니다.";

    Selection.Replace What:="" & Chr(32) & "" & Chr(10) & "", Replacement:="" & Chr(10) & "", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2

End Sub

' 각 행 높이 자동설정
Sub AllRowsAutoFit(control As IRibbonControl)
    'ActiveSheet.UsedRange.EntireRow.AutoFit
    Selection.EntireRow.AutoFit
End Sub

'각 행의 높이를 10 증가
Sub IncreaseRowHeightBy10(control As IRibbonControl)

    If TypeName(Selection) = "Range" Then

        Selection.Rows(1).RowHeight = WorksheetFunction.Min(Selection.Rows(1).RowHeight + 10, 409.5) '409.5가 최대임

    End If

End Sub


' 법령저장을 위해 열(복수열) 설정
Sub SetLawColumn(control As IRibbonControl)

    Selection.EntireColumn.Select

    '행 변경
    Selection.ColumnWidth = 80

    With Selection
        .HorizontalAlignment = xlLeft '수평왼쪽정렬
        .VerticalAlignment = xlTop '수직상단정렬
        .WrapText = True '텍스트겹치기
    End With

    ActiveWindow.Zoom = 80

End Sub

Public Sub ReplaceJoBracket(control As IRibbonControl)
    Dim rngUsed As Range
    Dim rngCell As Range
    
    ' 현재 시트의 사용 영역(UsedRange)을 가져옴
    Set rngUsed = ActiveSheet.UsedRange
    
    ' 사용 영역 내 각 셀을 순회
    For Each rngCell In rngUsed
        If Not IsEmpty(rngCell) Then
            ' 셀의 값이 문자열인 경우에만 Replace 수행
            If VarType(rngCell.Value) = vbString Then
                rngCell.Value = Replace(rngCell.Value, "조 (", "조(")
            End If
        End If
    Next rngCell
End Sub



Sub WhoAlert(control As IRibbonControl)
    
    MsgBox "제작 : 박병" & vbCrLf & _
    "목적 : 법령정보시스템 엑셀파싱" & vbCrLf & _
    "버전 : 0.1.0 (250313)" & vbCrLf & _
    "문의 : github.com/kinphw/LawParser" & vbCrLf

End Sub
