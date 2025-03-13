Attribute VB_Name = "MCompare"
'namespace=vba-files\Tracker\
Option Explicit

' 리본 메뉴 버튼 (전체 비교)
Public Sub ButtonCompareAll(control As IRibbonControl)
    Dim response As VbMsgBoxResult
    response = MsgBox("전체 A열/B열 범위를 대상으로 실행하시겠습니까?", _
                      vbQuestion + vbYesNo, "실행 확인")
    If response = vbYes Then
        Dim comparer As New CCompare
        comparer.CompareAll
    Else
        MsgBox "실행을 취소했습니다.", vbInformation, "취소"
    End If
End Sub

' 리본 메뉴 버튼 (현재 선택된 행만 비교)
Public Sub ButtonCompareSelection(control As IRibbonControl)
    Dim response As VbMsgBoxResult
    response = MsgBox("현재 선택된 셀의 행만 실행하시겠습니까?", _
                      vbQuestion + vbYesNo, "실행 확인")
    If response = vbYes Then
        Dim comparer As New CCompare
        comparer.CompareSelection
    Else
        MsgBox "실행을 취소했습니다.", vbInformation, "취소"
    End If
End Sub

