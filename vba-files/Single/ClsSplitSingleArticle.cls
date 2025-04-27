VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsSplitSingleArticle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'namespace=vba-files\Single\

Option Explicit

'=========================================================
' 클래스: clsLawTextSplitter
'
' - bHang, bHo, bMok: 분할 기준을 설정 (True/False)
' * 항(①②…), 호(n. / n의m.), 목(가. 나. 다. …)
'
' - SplitCell: 주어진 셀의 텍스트를 줄단위로 분리하여,
' (bHang/bHo/bMok)에 해당하는 라인이 나오면
' 새 덩어리를 시작.
'
' - 분리된 각 줄은 "IndentLine"을 통해
' 항→공백2칸, 호→공백4칸, 목→공백6칸 으로 들여쓰기.
'=========================================================

'-----------------------------
' [Public 속성] 분기 조건
'-----------------------------
Public bHang As Boolean ' 항(원문자) 분할할지 여부
Public bHo As Boolean ' 호 분할할지 여부
Public bMok As Boolean ' 목(가.,나.,다. 등) 분할할지 여부

'-------------------------------------------------------------
' [Public 메서드] SplitCell
' - 지정된 셀의 텍스트를 (bHang/bHo/bMok)에 따라 분리
' - 분리 시, 각 줄은 IndentLine()으로 들여쓰기 적용
'-------------------------------------------------------------
Public Sub SplitCell(ByVal targetCell As Range)
    Dim originalText As String
    Dim cleansedText As String
    Dim lines As Variant
    Dim i As Long

    ' (1) 유효성 검사
    If targetCell Is Nothing Then Exit Sub
    If IsEmpty(targetCell.Value) Then Exit Sub

    ' (2) 원본문자 → 클렌징
    originalText = CStr(targetCell.Value)
    cleansedText = CleanUpLineFeed(originalText)

    ' (3) 줄 단위 분할
    lines = Split(cleansedText, vbLf)

    ' (4) 결과 덩어리를 담을 Collection
    Dim chunkList As New Collection

    ' (5) 현재 누적 중인 덩어리 (문자열)
    Dim currentChunk As String
    currentChunk = ""

    ' (6) "위에서 아래" 순서로 순회
    For i = LBound(lines) To UBound(lines)
        Dim oneLine As String
        oneLine = lines(i)

        ' -- 들여쓰기 적용 (항=2칸, 호=4칸, 목=6칸) --
        Dim processedLine As String
        processedLine = IndentLine(oneLine)

        ' "분리 기준점"인지 확인
        If ShouldSplitLine(oneLine) Then
            ' 만약 현재 덩어리에 어떤 내용이 이미 있으면
            ' 먼저 chunkList에 추가(확정) 후 새로 시작
            If currentChunk <> "" Then
                chunkList.Add currentChunk
                currentChunk = ""
            End If
        ' 새 덩어리는 현재 processedLine으로 시작
            currentChunk = processedLine
        Else
            ' 분리점이 아니면 현재 덩어리에 누적
            If currentChunk = "" Then
                currentChunk = processedLine
            Else
                currentChunk = currentChunk & vbLf & processedLine
            End If
        End If
    Next i

    ' (7) 마지막에 남은 덩어리도 추가
    If currentChunk <> "" Then
        chunkList.Add currentChunk
    End If

    ' (8) chunkList를 순서대로 시트에 기록
    Dim r As Long, c As Long
    r = targetCell.Row
    c = targetCell.Column

    Dim idx As Long
    For idx = 1 To chunkList.Count
        Dim chunkText As String
        chunkText = chunkList(idx)

        ' LF → CRLF 로 바꿔서 셀에서 멀티라인 표시
        chunkText = Replace(chunkText, vbLf, vbCrLf)

        If idx = 1 Then
            ' 첫 덩어리는 대상 셀에
            targetCell.Value = chunkText
        Else
            ' 두 번째 이후는 행 삽입 후 그 셀에
            Rows(r + 1).Insert Shift:=xlDown
            Cells(r + 1, c).Value = chunkText
            r = r + 1
        End If
    Next idx
End Sub


'-----------------------------------------------------
' [Private] 개행문자 클렌징
' - CRLF→LF 통일
' - 연속된 LF(빈줄) 2개 이상 → 1개로 축소
'-----------------------------------------------------
Private Function CleanUpLineFeed(ByVal txt As String) As String
    Dim temp As String
    temp = Replace(txt, vbCrLf, vbLf)

    Do While InStr(temp, vbLf & vbLf) > 0
        temp = Replace(temp, vbLf & vbLf, vbLf)
    Loop

    CleanUpLineFeed = temp
End Function

'-----------------------------------------------------
' [Private] 분리 기준점인지 판별
'
' - bHang = True & IsHangLine => 분리
' - bHo = True & IsHoLine => 분리
' - bMok = True & IsMokLine => 분리
'-----------------------------------------------------
Private Function ShouldSplitLine(ByVal oneLine As String) As Boolean
    Dim t As String
    t = Trim(oneLine)

    ' 항
    If bHang And IsHangLine(t) Then
        ShouldSplitLine = True
        Exit Function
    End If

    ' 호
    If bHo And IsHoLine(t) Then
        ShouldSplitLine = True
        Exit Function
    End If

    ' 목
    If bMok And IsMokLine(t) Then
        ShouldSplitLine = True
        Exit Function
    End If

    ShouldSplitLine = False
End Function


'-----------------------------------------------------
' [Private] 항(원문자) 판별
' - 간단히 첫 글자가 ①~⑨ (확장 시 [①-?] 등)
'-----------------------------------------------------
Private Function IsHangLine(ByVal lineText As String) As Boolean
    If lineText Like "[①-⑮]*" Then
        IsHangLine = True
    Else
        IsHangLine = False
    End If
End Function

'-----------------------------------------------------
' [Private] 호 판별
' - 정규식: "^\d+(의\d+)*\."
' - "1.", "1의2." 등으로 시작하면 True
' - 뒤에 무언가가 더 있어도 OK
'-----------------------------------------------------
Private Function IsHoLine(ByVal lineText As String) As Boolean
    Static reHo As Object
    If reHo Is Nothing Then
        Set reHo = CreateObject("VBScript.RegExp")
        reHo.pattern = "^\d+(의\d+)*\."
        reHo.IgnoreCase = False
        reHo.Global = False
    End If

    IsHoLine = reHo.Test(lineText)
End Function

'-----------------------------------------------------
' [Private] 목 판별
' - "가.", "나.", "다.", "라.", "마.", "바.",
' "사.", "아.", "자.", "차.", "카.", "타.",
' "파.", "하."
' - 정규식: "^(가|나|다|라|마|바|사|아|자|차|카|타|파|하)\."
' - 뒤에 공백/문자 가능
'-----------------------------------------------------
Private Function IsMokLine(ByVal lineText As String) As Boolean
    Static reMok As Object
    If reMok Is Nothing Then
        Set reMok = CreateObject("VBScript.RegExp")
        reMok.pattern = "^(가|나|다|라|마|바|사|아|자|차|카|타|파|하)\."
        reMok.IgnoreCase = False
        reMok.Global = False
    End If

    IsMokLine = reMok.Test(lineText)
End Function

'-----------------------------------------------------
' [Private] 들여쓰기(Indent) 부여
' - 항 => 공백 2칸
' - 호 => 공백 4칸
' - 목 => 공백 6칸
'
' * 여기서는 "어느 하나"라도 매칭되면 끝.
' (항인지? 아니면 호인지? 아니면 목인지?)
'-----------------------------------------------------
Private Function IndentLine(ByVal lineText As String) As String
    Dim t As String
    t = Trim(lineText)

    ' 1) 항인 줄?
    If IsHangLine(t) Then
        IndentLine = lineText ' 2칸 '0칸
        Exit Function
    End If

    ' 2) 호인 줄?
    If IsHoLine(t) Then
        IndentLine = " " & lineText ' 4칸 '2칸
        Exit Function
    End If

    ' 3) 목인 줄?
    If IsMokLine(t) Then
        IndentLine = " " & lineText ' 6칸 '4칸
        Exit Function
    End If

    ' 4) 그 외
    IndentLine = lineText
End Function
