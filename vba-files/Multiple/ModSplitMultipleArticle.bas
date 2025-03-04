Attribute VB_Name = "ModSplitMultipleArticle"
'namespace=vba-files\Multiple\
Option Explicit

'namespace=vba-files\Multiple\

'===============================
' (1) 리본메뉴와 연동되는 Public Sub
'===============================
Public Sub ApplySplitMultipleArticle(control As IRibbonControl)

    Dim rng As Range
    Dim oneCell As Range

    ' 현재 선택 영역 가져오기
    Set rng = Selection
    If rng Is Nothing Then
        MsgBox "선택된 셀이 없습니다."
        Exit Sub
    End If

    ' 선택된 각 셀에 대해 순차 처리
    For Each oneCell In rng
        If Not IsEmpty(oneCell) Then
            ' ★ SplitMultipleArticlesCell 호출 (다른 모듈)
            SplitMultipleArticlesCell oneCell
        End If
    Next oneCell

    MsgBox "선택 범위 내 모든 셀에 대한 '조문 분리'가 완료되었습니다."

End Sub
