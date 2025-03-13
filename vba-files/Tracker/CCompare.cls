VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCompare"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'namespace=vba-files\Tracker\

Option Explicit

'-----------------------------------------------------------------------------------------
' 클래스: CCompare
'   - 기존에 전역 변수/함수로 관리하던 Compare 로직을 모두 캡슐화
'   - 전역변수 대신 Private 멤버 변수를 사용
'   - CompareAll / CompareSelection 메서드로 외부에서 호출
'-----------------------------------------------------------------------------------------

' (1) 기존 "전역 변수"를 여기 클래스의 Private 멤버로 옮김
Private sDelimiter As String

Private sOrig As String         ' 변경 전 문자열
Private sNew As String          ' 변경 후 문자열

Private lWordCountOrig As Long  ' 어절수 (변경 전)
Private lWordCountNew As Long   ' 어절수 (변경 후)

Private lMaxRow As Long
Private lMaxCol As Long
Private lStartRow As Long

Private sResult As String

' 배열들
Private arrDelWordStart() As Variant
Private arrDelWordLen() As Variant
Private arrAddWordStart() As Variant
Private arrAddWordLen() As Variant

Private arrWordsOrig As Variant
Private arrWordsNew As Variant
Private arrWordsResult() As Variant

Private bHasOrig As Boolean
Private bHasNew As Boolean
Private bHasDelWords As Boolean
Private bHasAddWords As Boolean
Private bHasResult As Boolean

Private lWordIdxOrig As Long
Private lWordIdxNew As Long
Private lWordIdxResult As Long

Private lDelWordIdx As Long
Private lAddWordIdx As Long

Private vMatchPos As Variant
Private lResultCharCount As Long

' 결과 표시할 Range
Private rngOutput As Range


'-----------------------------------------------------------------------------------------
' (2) 전체 비교 (A열/B열 마지막 행까지)
'-----------------------------------------------------------------------------------------
Public Sub CompareAll()
    ' 어절 구분자는 공백
    sDelimiter = " "

    ' (a) A열/B열 마지막 행 찾기
    Dim lLastRowA As Long, lLastRowB As Long
    lLastRowA = Cells(Rows.Count, 1).End(xlUp).Row
    lLastRowB = Cells(Rows.Count, 2).End(xlUp).Row
    lMaxRow = Application.WorksheetFunction.Max(lLastRowA, lLastRowB)
    
    ' (b) A/B 열만 비교
    lMaxCol = 1
    lStartRow = 1

    ' (c) 실제 행 루프
    Dim r As Long
    For r = lStartRow To lMaxRow
        CompareOneRow r
    Next r
End Sub

'-----------------------------------------------------------------------------------------
' (3) 선택된 셀 영역의 모든 행을 비교 (중복 없이)
'-----------------------------------------------------------------------------------------
Public Sub CompareSelection()
    sDelimiter = " "
    lMaxCol = 1

    ' Dictionary를 이용해 중복 없는 행 목록 생성
    Dim dictRows As Object
    Set dictRows = CreateObject("Scripting.Dictionary")
    
    Dim rngSel As Range, c As Range
    Set rngSel = Selection
    
    Dim lRow As Long
    For Each c In rngSel
        lRow = c.Row
        If Not dictRows.Exists(lRow) Then
            dictRows.Add lRow, lRow
        End If
    Next c
    
    ' Dictionary에 모인 고유 행 번호들 각각에 대해 CompareOneRow 실행
    Dim key As Variant
    For Each key In dictRows.Keys
        CompareOneRow CLng(key)
    Next key
End Sub


'-----------------------------------------------------------------------------------------
' (4) 한 행에 대한 A열/B열 비교 → C열에 결과 표시
'-----------------------------------------------------------------------------------------
Private Sub CompareOneRow(ByVal lRowNum As Long)
    Dim lCol As Long
    
    ' lMaxCol=1 → A열 vs B열
    For lCol = 1 To lMaxCol
        
        ' (1) A열/B열 문자열
        sOrig = Cells(lRowNum, lCol).Value
        sNew = Cells(lRowNum, lCol + lMaxCol).Value  ' B열
        
        ' (2) 결과 출력할 셀 (C열)
        Dim rngOrig As Range, rngNew As Range
        Set rngOrig = Cells(lRowNum, lCol)
        Set rngNew = Cells(lRowNum, lCol + lMaxCol)
        Set rngOutput = Cells(lRowNum, lCol + 2 * lMaxCol)
        
        rngOutput.Clear

        ' (3) 분기 처리
        If Len(sOrig) = 0 And Len(sNew) = 0 Then
            ' 둘 다 비어있으면 회색 배경
            rngOutput.Interior.ColorIndex = 15

        ElseIf Len(sOrig) = 0 Then
            ' 변경 전만 비었음 → 신규(밑줄)
            rngOutput.Value = sNew
            With rngOutput.Font
                .ColorIndex = 14   ' 연두색
                .Underline = True
            End With

        ElseIf Len(sNew) = 0 Then
            ' 변경 후만 비었음 → 삭제(취소선)
            rngOutput.Value = sOrig
            With rngOutput.Font
                .ColorIndex = 3    ' 빨간색
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

'-----------------------------------------------------------------------------------------
' (5) 보조: 문자열 Split → 전역(=클래스 Private) 배열에 담기
'-----------------------------------------------------------------------------------------
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

'-----------------------------------------------------------------------------------------
' (6) 핵심: "어절" 단위로 변경 전/후 비교 → 배열 기록
'-----------------------------------------------------------------------------------------
Private Sub CompareText()

    ' (1) 결과 배열 최대길이로 할당
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

        ' (3-1) 동일 어절
        If LCase(sWordOrig) = LCase(sWordNew) And (sWordOrig <> "") Then
            
            lWordIdxResult = lWordIdxResult + 1
            ReDim Preserve arrWordsResult(1 To lWordIdxResult)
            arrWordsResult(lWordIdxResult) = sWordOrig
            
            lResultCharCount = lResultCharCount + 1 + Len(sWordOrig)
            
            ' 후 배열에서 해당 어절 소진
            arrWordsNew(lWordIdxNew) = vbNullString
            
            lWordIdxOrig = lWordIdxOrig + 1
            lWordIdxNew = lWordIdxNew + 1

        Else
            ' (3-2) 다른 어절이면, 변경 후에 존재하는지 Match
            Dim sTmpWord As String
            sTmpWord = sWordOrig
            
            If sTmpWord = "" Then
                ' 변경 전이 이미 소진된 상태
                Exit Do
            End If
            
            vMatchPos = Application.Match(sTmpWord, arrWordsNew, 0)
            
            If IsError(vMatchPos) Then
                ' 변경 후에 없는 단어 → 삭제된 단어
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
                ' 변경 후 배열에 같은 단어 존재
                ' 그 앞쪽 단어들은 '추가된' 것으로 간주
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

    ' (6) 결과 문자열
    sResult = Join(arrWordsResult, sDelimiter)
End Sub

'-----------------------------------------------------------------------------------------
' (7) DisplayText: CompareText 결과를 셀에 표시 (밑줄/취소선)
'-----------------------------------------------------------------------------------------
Private Sub DisplayText()
    rngOutput.Clear
    If bHasResult Then
        rngOutput.Value = sResult
    End If

    ' 삭제된 어절(취소선)
    If bHasDelWords Then
        Dim i As Long
        For i = 1 To UBound(arrDelWordStart)
            With rngOutput.Characters( _
                Start:=arrDelWordStart(i), _
                Length:=arrDelWordLen(i) _
            ).Font
                .ColorIndex = 3         ' 빨간색
                .Strikethrough = True
            End With
        Next i
    End If
    
    ' 추가된 어절(밑줄)
    If bHasAddWords Then
        Dim j As Long
        For j = 1 To UBound(arrAddWordStart)
            With rngOutput.Characters( _
                Start:=arrAddWordStart(j), _
                Length:=arrAddWordLen(j) _
            ).Font
                .ColorIndex = 14       ' 연두색
                .Underline = True
            End With
        Next j
    End If

    rngOutput.WrapText = True
End Sub


