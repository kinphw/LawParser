Attribute VB_Name = "ModMultipleArticle_SplitLogic"
Option Explicit

'namespace=vba-files\Multiple\

'===============================
' (2) 셀의 텍스트를 쪼개는 핵심 로직
'===============================
Public Sub SplitMultipleArticlesCell(ByVal targetCell As Range)
    Dim originalText As String
    Dim cleansedText As String

    Dim re As Object  ' Late binding (CreateObject)
    Dim articleMatches As Object ' MatchCollection
    Dim pattern As String
    Dim currentRow As Long, currentCol As Long

    ' 유효성 검사
    If targetCell Is Nothing Then Exit Sub
    If IsEmpty(targetCell.Value) Then Exit Sub

    ' [A] 원본 텍스트
    originalText = targetCell.Value

    ' [B] 전처리
    cleansedText = CleanseAndRemoveLeadingSpaces(originalText)
    cleansedText = RemoveTargetLines(cleansedText)

    ' [C] 정규식 세팅
    ' (장, 절, 조를 줄 맨 앞에서만 찾되, 조는 괄호 필수 + 하이픈 처리)
    ' pattern = "((?:^|\n)제\d+(?:-\d+)?(?:장|절|조(?:의\d+)?\([^)]*\)))([\s\S]*?)(?=((?:^|\n)제\d+(?:-\d+)?(?:장|절|조(?:의\d+)?\([^)]*\))|$))"
    pattern = "((?:^|\n)제\d+(?:-\d+)?(?:장|절|조(?:의\d+)?(?:\([^)]*\)|\s*삭제<[^>]+>)))([\s\S]*?)(?=((?:^|\n)제\d+(?:-\d+)?(?:장|절|조(?:의\d+)?(?:\([^)]*\)|\s*삭제<[^>]+>))|$))"

    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = False
    re.MultiLine = False
    re.pattern = pattern

    ' [D] 매칭
    Set articleMatches = re.Execute(cleansedText)

    ' [E] 매칭이 없으면 → 전체 텍스트 그대로
    If articleMatches.Count = 0 Then
        targetCell.Value = cleansedText
        Exit Sub
    End If

    ' [F] 매칭 결과를 행 단위로 쪼개어 삽입
    currentRow = targetCell.Row
    currentCol = targetCell.Column

    Dim i As Long
    For i = 0 To articleMatches.Count - 1
        Dim oneArticle As String
        oneArticle = articleMatches(i).Value

        ' 전처리 (앞뒤 개행/공백 제거)
        oneArticle = CleanArticleText(oneArticle)

        If i = 0 Then
            ' 첫 번째 덩어리는 현재 셀에
            targetCell.Value = oneArticle
        Else
            ' 두 번째 이후 덩어리는 아래 행에 순차 삽입
            Rows(currentRow + 1).Insert Shift:=xlDown
            Cells(currentRow + 1, currentCol).Value = oneArticle
            currentRow = currentRow + 1
        End If
    Next i

End Sub
