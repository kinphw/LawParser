Attribute VB_Name = "ModSplitMultipleArticle"
'namespace=vba-files\Multiple\

Option Explicit

' 진입점
Public Sub ApplySplitMultipleArticle(control As IRibbonControl)

    Dim rng As Range
    Dim oneCell As Range

    Set rng = Selection
    If rng Is Nothing Then
        MsgBox "선택된 셀이 없습니다."
        Exit Sub
    End If

    ' 선택된 각 셀에 대해 순차 처리
    For Each oneCell In rng
        If Not IsEmpty(oneCell) Then
            SplitMultipleArticlesCell oneCell
        End If
    Next oneCell

    MsgBox "선택 범위 내 모든 셀에 대한 '조문 분리'가 완료되었습니다."
End Sub


' Level2

'--------------------------------------------------
' [2] 단일 셀에서 "제XX조(…)" 구간을
' 정규식으로 추출하여 행 단위로 분리
'--------------------------------------------------
Private Sub SplitMultipleArticlesCell(ByVal targetCell As Range)
    Dim originalText As String
    Dim cleansedText As String

    Dim re As Object ' Late binding (CreateObject)
    'Dim re As New RegExp ' Early binding (참조 설정 필요)

    Dim articleMatches As Object ' MatchCollection
    Dim oneMatch As Object ' Match

    Dim pattern As String
    Dim currentRow As Long
    Dim currentCol As Long

    ' 유효성 검사
    If targetCell Is Nothing Then Exit Sub
    If IsEmpty(targetCell.Value) Then Exit Sub

    ' [A] 원본 텍스트 가져오기
    originalText = targetCell.Value

    ' [B] 중복된 개행문자 제거 (클렌징)
    'cleansedText = ReplaceConsecutiveLineFeeds(originalText)
    cleansedText = CleanseAndRemoveLeadingSpaces(originalText)

    ' 로직 수기 추가
    cleansedText = RemoveTargetLines(cleansedText)

    ' [C] 정규식 세팅
    ' - "제\d+조\("로 시작해서 괄호 안 문구 닫고,
    ' 그 뒤 임의의 문자(개행 포함)를 최소로 반복(*?)
    ' 다음에 다시 "제\d+조\("가 나타나거나 문자열 종료($) 시까지 한 덩어리
    'pattern = "제\d+조\([^)]*\)([\s\S]*?)(?=제\d+조\(|$)"
    'pattern = "제\d+조(?:의\d+)?\([^)]*\)(]\s\S]*?)(?=제\d+조(?:의\d+)?\(|$)" '수정
    'pattern = "제\d+조(?:의\d+)?\([^)]*\)([\s\S]*?)(?=제\d+조(?:의\d+)?\(|$)" '기존 잘되던거
    'pattern = "제\d+(?:장|조(?:의\d+)?\([^)]*\))([\s\S]*?)(?=제\d+(?:장|조(?:의\d+)?\([^)]*\))|$)" '장추가
    'pattern = "제\d+(?:장|절|조(?:의\d+)?(?:\([^)]*\))?)([\s\S]*?)(?=제\d+(?:장|절|조(?:의\d+)?(?:\([^)]*\))?)|$)" '장+쩔추가
    'pattern = "제\d+(?:장|절|조(?:의\d+)?\([^)]*\))([\s\S]*?)(?=제\d+(?:장|절|조(?:의\d+)?\([^)]*\))|$)" '직전대비 괄호보완
    'pattern = "((?:^|\n)제\d+(?:장|절)\s|제\d+조(?:의\d+)?\([^)]*\))([\s\S]*?)(?=((?:^|\n)제\d+(?:장|절)\s|제\d+조(?:의\d+)?\([^)]*\)|$))" '장분리 보완
    'pattern = "((?:^|\n)제\d+(?:-\d+)?(?:장|절)\s|제\d+(?:-\d+)?조(?:의\d+)?\([^)]*\))([\s\S]*?)(?=((?:^|\n)제\d+(?:-\d+)?(?:장|절)\s|제\d+(?:-\d+)?조(?:의\d+)?\([^)]*\)|$))" '외국환거래규정케이스 보완
    pattern = "((?:^|\n)제\d+(?:-\d+)?(?:장|절|조(?:의\d+)?\([^)]*\)))([\s\S]*?)(?=((?:^|\n)제\d+(?:-\d+)?(?:장|절|조(?:의\d+)?\([^)]*\))|$))" '한번 더 보완

    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = False
    're.MultiLine = True 'Debug
    re.MultiLine = False
    re.pattern = pattern

    ' [D] 매칭 수행
    Set articleMatches = re.Execute(cleansedText)

    ' [E] 만약 매칭이 없다면(=조문 패턴이 안 잡히면) 그냥 텍스트만 정리해서 종료
    If articleMatches.Count = 0 Then
        targetCell.Value = cleansedText
        Exit Sub
    End If

    ' [F] 첫 번째 매칭은 "현재 셀"에 기록,
    ' 나머지는 "아래 행"에 삽입
    currentRow = targetCell.Row
    currentCol = targetCell.Column

    Dim i As Long
    For i = 0 To articleMatches.Count - 1
        Dim oneArticle As String
        oneArticle = articleMatches(i).Value

        ' ★★★ 전처리 함수 호출 (앞뒤 개행/공백 제거) ★★★
        oneArticle = CleanArticleText(oneArticle)        

        If i = 0 Then
            ' 첫 번째 Article은 현재 셀에 저장
            targetCell.Value = oneArticle
        Else
            ' 두 번째부터는 행 삽입 후 그 셀에 기록
            Rows(currentRow + 1).Insert Shift:=xlDown
            Cells(currentRow + 1, currentCol).Value = oneArticle
            currentRow = currentRow + 1
        End If
    Next i

    'MsgBox "조문 분리 작업이 완료되었습니다."

End Sub

Private Function CleanseAndRemoveLeadingSpaces(ByVal txt As String) As String
    Dim temp As String
    Dim arrLines As Variant
    Dim i As Long

    ' 1) CRLF -> LF 통일
    temp = Replace(txt, vbCrLf, vbLf)

    ' 2) 연속된 LF 2개 이상을 1개로
    Do While InStr(temp, vbLf & vbLf) > 0
        temp = Replace(temp, vbLf & vbLf, vbLf)
    Loop

    ' 3) 각 줄로 분할
    arrLines = Split(temp, vbLf)

    ' 4) 각 줄의 맨 앞 공백 제거 (LTrim)
    For i = LBound(arrLines) To UBound(arrLines)
        arrLines(i) = LTrim(arrLines(i))
    Next i

    ' 5) 다시 LF로 합침
    temp = Join(arrLines, vbLf)

    CleanseAndRemoveLeadingSpaces = temp
End Function
'[3] 조문체계도버튼, 연혁 등 삭제

Private Function RemoveTargetLines(ByVal inputText As String) As String

    Dim temp As String
    Dim arrLines As Variant
    Dim resultLines As Collection
    Dim i As Long

    ' 제외할 단어들을 배열로 선언
    Dim excludeWords As Variant
    excludeWords = Array("조문체계도버튼", "연혁", "관련규제버튼", "위임행정규칙버튼", "생활법령버튼", "위임행정규칙")

    ' (1) 우선 CRLF -> LF 로 통일
    temp = Replace(inputText, vbCrLf, vbLf)

    ' (2) 여러 줄로 Split
    arrLines = Split(temp, vbLf)

    ' (3) 결과를 담을 Collection 객체 준비
    Set resultLines = New Collection

    ' (4) 각 줄 순회
    For i = LBound(arrLines) To UBound(arrLines)
        Dim oneLine As String
        oneLine = Trim(arrLines(i))
        
        ' 빈 줄이 아니고 제외대상 단어가 없는 경우만 추가
        If oneLine <> "" And Not ContainsAnyWord(oneLine, excludeWords) Then
            resultLines.Add oneLine
        End If
    Next i

    
    Dim outputText As String
    ' Collection을 문자열로 변환
    If resultLines.Count = 0 Then
        outputText = ""
    Else
        Dim lineVal As Variant
        For Each lineVal In resultLines
            If outputText = "" Then
                outputText = lineVal
            Else
                outputText = outputText & vbLf & lineVal
            End If
        Next lineVal
    End If
    
    RemoveTargetLines = outputText
    
End Function


' 문자열에 배열의 단어가 포함되어 있는지 확인하는 함수
Private Function ContainsAnyWord(str As String, words As Variant) As Boolean
    Dim word As Variant
    
    For Each word In words
        If InStr(str, CStr(word)) > 0 Then
            ContainsAnyWord = True
            Exit Function
        End If
    Next word
    
    ContainsAnyWord = False
End Function


Private Function CleanArticleText(ByVal txt As String) As String
    Dim temp As String
    temp = txt

    ' 1) CRLF -> LF로 통일
    temp = Replace(temp, vbCrLf, vbLf)

    ' 2) 전체 Trim(좌우 공백 제거)
    temp = Trim(temp)

    ' 3) 앞쪽에 남아있는 연속된 vbLf 제거
    Do While Left(temp, 1) = vbLf
        temp = Mid(temp, 2)
    Loop

    ' 4) 뒤쪽에 남아있는 연속된 vbLf 제거
    Do While Right(temp, 1) = vbLf
        temp = Left(temp, Len(temp) - 1)
    Loop

    CleanArticleText = temp
End Function    