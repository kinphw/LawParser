Attribute VB_Name = "MCompare"
'namespace=vba-files\Tracker\
Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 공용변수부 (전역변수) - 이름만 영어로 변경, 접두어 적용
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private sDelimiter As String          ' Delimiter

Private sOrig As String               ' 문자열 변경 전
Private sNew As String                ' 문자열 변경 후

Private lWordCountOrig As Long        ' 어절수 (변경 전)
Private lWordCountNew As Long         ' 어절수 (변경 후)

Private lRowPos As Long               ' 위치(행)
Private lColPos As Long               ' 위치(열)
Private lMaxRow As Long               ' 범위(행)
Private lMaxCol As Long               ' 범위(열)

Private lStartRow As Long             ' 작업 시작 행

' Range 참조
Private rngOrig As Range
Private rngNew As Range
Private rngOutput As Range

' 요청사항: 특별출력 (현재 사용 안 함)
Private rngOutputSpecial As Range

' 최종 문자열 결과
Private sResult As String

' 삭제/추가 문구 인덱스
Private arrDelWordStart() As Variant
Private arrDelWordLen() As Variant
Private arrAddWordStart() As Variant
Private arrAddWordLen() As Variant

' 변경 전/후 "어절" 배열
Private arrWordsOrig As Variant
Private arrWordsNew As Variant
Private arrWordsResult() As Variant

' 존재 여부
Private bHasOrig As Boolean
Private bHasNew As Boolean
Private bHasDelWords As Boolean
Private bHasAddWords As Boolean
Private bHasResult As Boolean

' 인덱스 제어
Private lWordIdxOrig As Long
Private lWordIdxNew As Long
Private lWordIdxResult As Long

Private lDelWordIdx As Long
Private lAddWordIdx As Long

Private vMatchPos As Variant

' 결과 문자열에서 몇 글자까지 썼는지 누적
Private lResultCharCount As Long


'''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 리본 메뉴에서 호출되는 두 버튼
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ButtonCompareAll(control As IRibbonControl)
    Dim lResponse As VbMsgBoxResult
    lResponse = MsgBox("전체 A열/B열 범위를 대상으로 실행하시겠습니까?", vbQuestion + vbYesNo, "실행 확인")
    If lResponse = vbYes Then
        CompareMain True
    Else
        MsgBox "실행을 취소했습니다.", vbInformation, "취소"
    End If
End Sub

Public Sub ButtonCompareSelection(control As IRibbonControl)
    Dim lResponse As VbMsgBoxResult
    lResponse = MsgBox("현재 선택된 셀의 행만 실행하시겠습니까?", vbQuestion + vbYesNo, "실행 확인")
    If lResponse = vbYes Then
        CompareMain False
    Else
        MsgBox "실행을 취소했습니다.", vbInformation, "취소"
    End If
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 메인 루틴
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CompareMain(Optional ByVal doAll As Boolean = True)
    sDelimiter = " "  ' 어절 구분자(공백)

    If doAll Then
        ' A열/B열 전체 범위에서 마지막 행 찾기
        Dim lLastRowA As Long, lLastRowB As Long
        lLastRowA = Cells(Rows.Count, 1).End(xlUp).Row  ' A열
        lLastRowB = Cells(Rows.Count, 2).End(xlUp).Row  ' B열
        lMaxRow = Application.WorksheetFunction.Max(lLastRowA, lLastRowB)
        
        lMaxCol = 1       ' A/B 열만 비교
        lStartRow = 1     ' 일반적으로 1행부터 시작

        Dim lRow As Long
        For lRow = lStartRow To lMaxRow
            CompareOneRow lRow
        Next lRow
    Else
        ' 선택된 셀의 행만 실행
        lMaxCol = 1
        Dim lSelRow As Long
        lSelRow = Selection.Row
        
        CompareOneRow lSelRow
    End If
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
' "특정 한 행" A열/B열 비교 → C열에 결과 출력
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CompareOneRow(ByVal lRowNum As Long)
    Dim lCol As Long
    
    ' lMaxCol = 1 이므로, lCol=1만 반복 (A열 vs B열)
    For lCol = 1 To lMaxCol
        
        ' (1) A열 vs B열 문자열
        sOrig = Cells(lRowNum, lCol).Value
        sNew = Cells(lRowNum, lCol + lMaxCol).Value  ' lCol=1이면 +1 → B열

        ' (2) 결과 출력할 셀 (C열)
        Set rngOrig = Cells(lRowNum, lCol)
        Set rngNew = Cells(lRowNum, lCol + lMaxCol)
        Set rngOutput = Cells(lRowNum, lCol + 2 * lMaxCol)  ' C열

        rngOutput.Clear

        ' (3) 분기 처리
        If Len(sOrig) = 0 And Len(sNew) = 0 Then
            ' 둘 다 비어있으면 회색 배경
            rngOutput.Interior.ColorIndex = 15

        ElseIf Len(sOrig) = 0 Then
            ' 변경 전만 비었음 → 신규(밑줄)
            rngOutput.Value = sNew
            With rngOutput.Font
                .ColorIndex = 14      ' 연두색
                .Underline = True
            End With

        ElseIf Len(sNew) = 0 Then
            ' 변경 후만 비었음 → 삭제(취소선)
            rngOutput.Value = sOrig
            With rngOutput.Font
                .ColorIndex = 3       ' 빨간색
                .Strikethrough = True
            End With

        Else
            ' 둘 다 값이 있음 → 어절 단위 비교
            PrepareWords
            CompareText
            DisplayText
        End If
    Next lCol
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
' (보조) 문자열을 Split → 전역 배열에 세팅
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PrepareWords()
    Dim arrTmp As Variant
    Dim lCount As Long
    
    ' --- 변경 전 ---
    If Len(sOrig) = 0 Then
        bHasOrig = False
        lWordCountOrig = 0
    Else
        arrTmp = Split(sOrig, sDelimiter)  ' 0-based
        lCount = UBound(arrTmp) - LBound(arrTmp) + 1
        If lCount <= 0 Then
            bHasOrig = False
            lWordCountOrig = 0
        Else
            bHasOrig = True
            ReDim arrWordsOrig(1 To lCount)
            
            Dim i As Long
            For i = 1 To lCount
                arrWordsOrig(i) = arrTmp(i - 1)
            Next i
            lWordCountOrig = lCount
        End If
    End If
    
    ' --- 변경 후 ---
    If Len(sNew) = 0 Then
        bHasNew = False
        lWordCountNew = 0
    Else
        arrTmp = Split(sNew, sDelimiter)
        lCount = UBound(arrTmp) - LBound(arrTmp) + 1
        If lCount <= 0 Then
            bHasNew = False
            lWordCountNew = 0
        Else
            bHasNew = True
            ReDim arrWordsNew(1 To lCount)
            
            Dim j As Long
            For j = 1 To lCount
                arrWordsNew(j) = arrTmp(j - 1)
            Next j
            lWordCountNew = lCount
        End If
    End If
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
' (핵심1) "어절" 단위 변경 전/후 비교 → 전역배열 기록
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CompareText()

    ' (1) 결과 배열 미리 최댓길이로 ReDim
    ReDim arrWordsResult(1 To (lWordCountOrig + lWordCountNew))
    bHasResult = True

    ' (2) 삭제/추가 관련 배열 초기화
    bHasDelWords = False
    bHasAddWords = False
    lDelWordIdx = 0
    lAddWordIdx = 0

    lResultCharCount = 1
    
    lWordIdxOrig = 1
    lWordIdxNew = 1
    lWordIdxResult = 0

    ' (3) 1차 루프
    Do
        Dim sWordOrig As String
        Dim sWordNew As String
        
        ' 변경 전 어절
        If (Not bHasOrig) Or (lWordIdxOrig > lWordCountOrig) Then
            sWordOrig = ""
        Else
            sWordOrig = arrWordsOrig(lWordIdxOrig)
        End If
        
        ' 변경 후 어절
        If (Not bHasNew) Or (lWordIdxNew > lWordCountNew) Then
            sWordNew = ""
        Else
            sWordNew = arrWordsNew(lWordIdxNew)
        End If

        ' (3-1) 동일 어절이면 결과에 그대로 추가
        If LCase(sWordOrig) = LCase(sWordNew) And (sWordOrig <> "") Then
            
            lWordIdxResult = lWordIdxResult + 1
            ReDim Preserve arrWordsResult(1 To lWordIdxResult)
            arrWordsResult(lWordIdxResult) = sWordOrig
            
            lResultCharCount = lResultCharCount + 1 + Len(sWordOrig)
            
            ' 후 배열에서 해당 어절 사용 소진
            arrWordsNew(lWordIdxNew) = vbNullString
            
            lWordIdxOrig = lWordIdxOrig + 1
            lWordIdxNew = lWordIdxNew + 1

        Else
            ' (3-2) 다른 어절이면, 변경 후에 있는지 Match
            Dim sTmpWord As String
            sTmpWord = sWordOrig
            
            If sTmpWord = "" Then
                ' 변경 전 소진된 상태
                Exit Do
            End If
            
            vMatchPos = Application.Match(sTmpWord, arrWordsNew, 0)
            
            If IsError(vMatchPos) Then
                ' 변경 후에 없는 단어 → "삭제된" 단어
                lWordIdxResult = lWordIdxResult + 1
                ReDim Preserve arrWordsResult(1 To lWordIdxResult)
                arrWordsResult(lWordIdxResult) = sTmpWord
                
                If Not bHasDelWords Then
                    bHasDelWords = True
                    ReDim arrDelWordStart(1 To 1)
                    ReDim arrDelWordLen(1 To 1)
                End If
                
                lDelWordIdx = lDelWordIdx + 1
                ReDim Preserve arrDelWordStart(1 To lDelWordIdx)
                ReDim Preserve arrDelWordLen(1 To lDelWordIdx)
                
                arrDelWordStart(lDelWordIdx) = lResultCharCount
                arrDelWordLen(lDelWordIdx) = Len(sTmpWord)
                
                lResultCharCount = lResultCharCount + 1 + Len(sTmpWord)
                
                lWordIdxOrig = lWordIdxOrig + 1
            
            Else
                ' 변경 후 배열에 같은 단어가 존재
                ' 그 앞에 추가된 단어들을 처리
                Do
                    If lWordIdxNew > lWordCountNew Then
                        Exit Do
                    End If
                    
                    If LCase(arrWordsNew(lWordIdxNew)) = LCase(sTmpWord) Then
                        Exit Do
                    End If
                    
                    ' 추가된 단어
                    lWordIdxResult = lWordIdxResult + 1
                    ReDim Preserve arrWordsResult(1 To lWordIdxResult)
                    
                    Dim sAddWord As String
                    sAddWord = arrWordsNew(lWordIdxNew)
                    arrWordsResult(lWordIdxResult) = sAddWord

                    If Not bHasAddWords Then
                        bHasAddWords = True
                        ReDim arrAddWordStart(1 To 1)
                        ReDim arrAddWordLen(1 To 1)
                    End If

                    lAddWordIdx = lAddWordIdx + 1
                    ReDim Preserve arrAddWordStart(1 To lAddWordIdx)
                    ReDim Preserve arrAddWordLen(1 To lAddWordIdx)

                    arrAddWordStart(lAddWordIdx) = lResultCharCount
                    arrAddWordLen(lAddWordIdx) = Len(sAddWord)

                    lResultCharCount = lResultCharCount + 1 + Len(sAddWord)

                    ' 중복 매칭 방지
                    arrWordsNew(lWordIdxNew) = vbNullString
                    lWordIdxNew = lWordIdxNew + 1
                    
                Loop
            End If
        End If

        ' (3-3) 둘 중 하나라도 소진되면 종료
        If (lWordIdxOrig > lWordCountOrig) Or (lWordIdxNew > lWordCountNew) Then
            Exit Do
        End If
    Loop

    ' (4) 남은 "추가된" 어절 처리
    Do While lWordIdxNew <= lWordCountNew
        Dim sRemB As String
        sRemB = arrWordsNew(lWordIdxNew)
        If sRemB <> "" Then
            lWordIdxResult = lWordIdxResult + 1
            ReDim Preserve arrWordsResult(1 To lWordIdxResult)
            arrWordsResult(lWordIdxResult) = sRemB

            If Not bHasAddWords Then
                bHasAddWords = True
                ReDim arrAddWordStart(1 To 1)
                ReDim arrAddWordLen(1 To 1)
            End If
            
            lAddWordIdx = lAddWordIdx + 1
            ReDim Preserve arrAddWordStart(1 To lAddWordIdx)
            ReDim Preserve arrAddWordLen(1 To lAddWordIdx)
            
            arrAddWordStart(lAddWordIdx) = lResultCharCount
            arrAddWordLen(lAddWordIdx) = Len(sRemB)
            
            lResultCharCount = lResultCharCount + 1 + Len(sRemB)
        End If
        lWordIdxNew = lWordIdxNew + 1
    Loop

    ' (5) 남은 "삭제된" 어절 처리
    Do While lWordIdxOrig <= lWordCountOrig
        Dim sRemA As String
        sRemA = arrWordsOrig(lWordIdxOrig)
        If sRemA <> "" Then
            lWordIdxResult = lWordIdxResult + 1
            ReDim Preserve arrWordsResult(1 To lWordIdxResult)
            arrWordsResult(lWordIdxResult) = sRemA

            If Not bHasDelWords Then
                bHasDelWords = True
                ReDim arrDelWordStart(1 To 1)
                ReDim arrDelWordLen(1 To 1)
            End If
            
            lDelWordIdx = lDelWordIdx + 1
            ReDim Preserve arrDelWordStart(1 To lDelWordIdx)
            ReDim Preserve arrDelWordLen(1 To lDelWordIdx)

            arrDelWordStart(lDelWordIdx) = lResultCharCount
            arrDelWordLen(lDelWordIdx) = Len(sRemA)

            lResultCharCount = lResultCharCount + 1 + Len(sRemA)
        End If
        lWordIdxOrig = lWordIdxOrig + 1
    Loop

    ' (6) 최종 결과 문자열
    sResult = Join(arrWordsResult, sDelimiter)
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
' (핵심2) CompareText 결과를 셀에 표시 (밑줄/취소선)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub DisplayText()
    Dim i As Long
    
    rngOutput.Clear
    If bHasResult Then
        rngOutput.Value = sResult
    End If

    ' 삭제된 어절(취소선)
    If bHasDelWords Then
        For i = 1 To UBound(arrDelWordStart)
            With rngOutput.Characters(Start:=arrDelWordStart(i), Length:=arrDelWordLen(i)).Font
                .ColorIndex = 3           ' 빨간색
                .Strikethrough = True
            End With
        Next i
    End If
    
    ' 추가된 어절(밑줄)
    If bHasAddWords Then
        For i = 1 To UBound(arrAddWordStart)
            With rngOutput.Characters(Start:=arrAddWordStart(i), Length:=arrAddWordLen(i)).Font
                .ColorIndex = 14         ' 연두색
                .Underline = True
            End With
        Next i
    End If

    rngOutput.WrapText = True
End Sub

