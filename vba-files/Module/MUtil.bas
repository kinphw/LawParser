Attribute VB_Name = "MUtil"

Option Explicit

Public Function RemoveLeadingPatterns( _
    ByVal s As String, _
    Optional ByVal removeNumberDot As Boolean = True, _
    Optional ByVal removeNumberParen As Boolean = True, _
    Optional ByVal removeKoreanParen As Boolean = True, _
    Optional ByVal removeCircledNumbers As Boolean = True) As String
    
    Dim arrPatterns() As String
    Dim n As Long: n = 0
    
    ' 1) [공백]숫자.[공백] => ^\s*\d+\.\s*
    If removeNumberDot Then
        ReDim Preserve arrPatterns(n)
        arrPatterns(n) = "\d+\.\s*"
        n = n + 1
    End If
    
    ' 2) [공백]숫자)[공백] => ^\s*\d+\)\s*
    If removeNumberParen Then
        ReDim Preserve arrPatterns(n)
        arrPatterns(n) = "\d+\)\s*"
        n = n + 1
    End If
    
    ' 3) [공백][가나다라마바사아자차카타파하])[공백] => ^\s*[가나다라마바사아자차카타파하]\)\s*
    If removeKoreanParen Then
        ReDim Preserve arrPatterns(n)
        arrPatterns(n) = "[가나다라마바사아자차카타파하]\)\s*"
        n = n + 1
    End If

    ' 4) ①~⑨ (유니코드 원형 숫자) 제거 => ^\s*[\u2460-\u2468]\s*
    If removeCircledNumbers Then
        ReDim Preserve arrPatterns(n)
        arrPatterns(n) = "[①-⑨]\s*"
        n = n + 1
    End If
    
    ' 제거할 패턴이 전혀 선택되지 않은 경우, 원본 그대로 리턴
    If n = 0 Then
        RemoveLeadingPatterns = s
        Exit Function
    End If
    
    ' 정규표현식 패턴 만들기
    ' ^\s* (?: 패턴1 | 패턴2 | 패턴3 | 패턴4 )
    Dim finalPattern As String
    finalPattern = "^\s*(?:" & Join(arrPatterns, "|") & ")"
    
    ' 정규표현식 객체 생성
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    With re
        .pattern = finalPattern
        .Global = False      ' 시작 부분 1회만 치환
        .IgnoreCase = True
    End With
    
    ' 치환 실행
    RemoveLeadingPatterns = re.Replace(s, "")
End Function

'선택영역 Remove
Sub RemoveSelection(control As IRibbonControl)

    Dim cell As Range
    
    ' 선택된 모든 셀을 순환하면서 LTrim 적용
    For Each cell In Selection
        ' 셀이 비어있지 않을 경우만 처리
        If Not IsEmpty(cell.Value) Then
            cell.Value = RemoveLeadingPatterns(cell.Value)
        End If
    Next cell
    
    MsgBox "항호처리 제거 완료!", vbInformation, "항호처리 제거 완료!"
    
End Sub


' 선택영역 LTrim
Sub TrimLeftSelection(control As IRibbonControl)
    Dim cell As Range
    
    ' 선택된 모든 셀을 순환하면서 LTrim 적용
    For Each cell In Selection
        ' 셀이 비어있지 않을 경우만 처리
        If Not IsEmpty(cell.Value) Then
            cell.Value = LTrim(cell.Value)
        End If
    Next cell
    
    MsgBox "좌측 공백 제거 완료!", vbInformation, "LTrim 적용 완료"
End Sub

