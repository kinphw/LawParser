Attribute VB_Name = "ModMultipleArticle_Utils"
Option Explicit

'namespace=vba-files\Multiple\

'===============================
' (3) 전처리 및 유틸 함수 모음
'===============================

Public Function CleanseAndRemoveLeadingSpaces(ByVal txt As String) As String
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

    ' 4) 각 줄의 맨 앞 공백 제거
    For i = LBound(arrLines) To UBound(arrLines)
        arrLines(i) = LTrim(arrLines(i))
    Next i

    ' 5) 다시 LF로 합침
    temp = Join(arrLines, vbLf)
    CleanseAndRemoveLeadingSpaces = temp
End Function


Public Function RemoveTargetLines(ByVal inputText As String) As String

    Dim temp As String
    Dim arrLines As Variant
    Dim resultLines As Collection
    Dim i As Long

    ' 제외할 단어들을 배열로 선언
    Dim excludeWords As Variant
    excludeWords = Array("조문체계도버튼", "연혁", "관련규제버튼", "위임행정규칙버튼", "생활법령버튼", "위임행정규칙")

    ' (1) CRLF -> LF 통일
    temp = Replace(inputText, vbCrLf, vbLf)

    ' (2) 여러 줄로 Split
    arrLines = Split(temp, vbLf)

    ' (3) 결과를 담을 Collection 객체
    Set resultLines = New Collection

    ' (4) 각 줄 순회
    For i = LBound(arrLines) To UBound(arrLines)
        Dim oneLine As String
        oneLine = Trim(arrLines(i))
        
        ' 빈 줄이 아니고 제외대상 단어가 없는 경우만
        If oneLine <> "" And Not ContainsAnyWord(oneLine, excludeWords) Then
            resultLines.Add oneLine
        End If
    Next i

    ' (5) 결과를 문자열로
    Dim outputText As String
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


Public Function ContainsAnyWord(str As String, words As Variant) As Boolean
    Dim word As Variant
    
    For Each word In words
        If InStr(str, CStr(word)) > 0 Then
            ContainsAnyWord = True
            Exit Function
        End If
    Next word
    
    ContainsAnyWord = False
End Function


Public Function CleanArticleText(ByVal txt As String) As String
    Dim temp As String
    temp = txt

    ' 1) CRLF -> LF
    temp = Replace(temp, vbCrLf, vbLf)

    ' 2) 양쪽 공백 제거
    temp = Trim(temp)

    ' 3) 앞쪽의 연속된 vbLf 제거
    Do While Left(temp, 1) = vbLf
        temp = Mid(temp, 2)
    Loop

    ' 4) 뒤쪽의 연속된 vbLf 제거
    Do While Right(temp, 1) = vbLf
        temp = Left(temp, Len(temp) - 1)
    Loop

    CleanArticleText = temp
End Function
