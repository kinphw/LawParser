Attribute VB_Name = "MCompare"
'namespace=vba-files\Tracker\
Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 공용변수부 (전역변수)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private 구분자 As String 'Delimiter (기본은 " ")

Private 문자열_변경전 As String
Private 문자열_변경후 As String

Private 어절수_문자열_변경전 As Long
Private 어절수_문자열_변경후 As Long

Private 위치_행 As Long
Private 위치_열 As Long
Private 범위_행 As Long
Private 범위_열 As Long

Private 작업시작행 As Long

' Range 참조
Private 위치_변경전 As Range
Private 위치_변경후 As Range
Private 위치_출력 As Range

' 요청사항: 행별로 신규/삭제/변경/NA 표시 목적 (현재 사용 안 함)
Private 위치_출력_특별 As Range

' 최종 문자열 결과
Private 문자열_작업결과 As String

' 삭제/추가 문구용 인덱스 배열 (동적)
Private 배열_삭제된문자열_시작위치() As Variant
Private 배열_삭제된문자열_길이() As Variant
Private 배열_추가된문자열_시작위치() As Variant
Private 배열_추가된문자열_길이() As Variant

' 변경 전/후 “어절” 배열
Private 배열_문자열_변경전 As Variant
Private 배열_문자열_변경후 As Variant
Private 배열_문자열_작업결과() As Variant

' 존재 여부에 대한 Boolean
Private 존재여부_문자열_변경전 As Boolean
Private 존재여부_문자열_변경후 As Boolean
Private 존재여부_삭제된문자열 As Boolean
Private 존재여부_추가된문자열 As Boolean
Private 존재여부_문자열_작업결과 As Boolean

' 인덱스 제어용
Private 어절순번_문자열_변경전 As Long
Private 어절순번_문자열_변경후 As Long
Private 어절순번_문자열_작업결과 As Long

Private 어절순번_삭제된문자열 As Long
Private 어절순번_추가된문자열 As Long

Private 일치하는위치 As Variant

' 결과 문자열에서 몇 글자까지 썼는지 누적
Private 글자수_작업결과 As Long

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 리본 메뉴에서 호출되는 두 버튼
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ButtonCompareAll(control As IRibbonControl)
    Dim response As VbMsgBoxResult
    response = MsgBox("전체 A열/B열 범위를 대상으로 실행하시겠습니까?", vbQuestion + vbYesNo, "실행 확인")
    If response = vbYes Then
        CompareMain True
    Else
        MsgBox "실행을 취소했습니다.", vbInformation, "취소"
    End If
End Sub

Public Sub ButtonCompareSelection(control As IRibbonControl)
    Dim response As VbMsgBoxResult
    response = MsgBox("현재 선택된 셀의 행만 실행하시겠습니까?", vbQuestion + vbYesNo, "실행 확인")
    If response = vbYes Then
        CompareMain False
    Else
        MsgBox "실행을 취소했습니다.", vbInformation, "취소"
    End If
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 메인 루틴
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CompareMain(Optional ByVal doAll As Boolean = True)
    구분자 = " "        ' 어절 구분자는 공백으로

    If doAll Then
        ' 전체 A열/B열에서 마지막 행까지만
        Dim lastRowA As Long, lastRowB As Long
        lastRowA = Cells(Rows.Count, 1).End(xlUp).Row  ' A열
        lastRowB = Cells(Rows.Count, 2).End(xlUp).Row  ' B열
        범위_행 = Application.WorksheetFunction.Max(lastRowA, lastRowB)
        
        범위_열 = 1      ' 하드코딩: A/B 열만 비교
        작업시작행 = 1   ' 보통 1행부터

        Dim r As Long
        For r = 작업시작행 To 범위_행
            CompareOneRow r
        Next r
    Else
        ' 현재 Selection이 있는 행만
        범위_열 = 1
        Dim selRow As Long
        selRow = Selection.Row
        
        CompareOneRow selRow
    End If
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
' "특정 한 행"에 대해 A열/B열 비교 → C열에 결과 출력
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CompareOneRow(ByVal rowNum As Long)
    Dim col As Long
    
    ' 범위_열=1 이므로, col=1 만 1회 반복 (A열, B열)
    For col = 1 To 범위_열
        
        ' (1) A열 vs B열 문자열 가져오기
        문자열_변경전 = Cells(rowNum, col).Value
        문자열_변경후 = Cells(rowNum, col + 범위_열).Value  ' col=1이면 +1 → B열

        ' (2) 결과 출력할 셀 (C열)
        Set 위치_변경전 = Cells(rowNum, col)
        Set 위치_변경후 = Cells(rowNum, col + 범위_열)
        Set 위치_출력 = Cells(rowNum, col + 2 * 범위_열)  ' col=1이면 +2 → C열

        위치_출력.Clear

        ' (3) 양쪽 문자열이 모두 비어 있는지, 한 쪽만 비었는지, 둘 다 있는지 분기
        If Len(문자열_변경전) = 0 And Len(문자열_변경후) = 0 Then
            ' 둘 다 비어 있으면 결과 셀을 회색 배경
            위치_출력.Interior.ColorIndex = 15

        ElseIf Len(문자열_변경전) = 0 Then
            ' 변경 전만 비어 있음 → 신규(밑줄)
            위치_출력.Value = 문자열_변경후
            With 위치_출력.Font
                .ColorIndex = 14        ' 연두색
                .Underline = True
            End With

        ElseIf Len(문자열_변경후) = 0 Then
            ' 변경 후만 비어 있음 → 삭제(취소선)
            위치_출력.Value = 문자열_변경전
            With 위치_출력.Font
                .ColorIndex = 3         ' 빨간색
                .Strikethrough = True
            End With

        Else
            ' 둘 다 값이 있음 → 어절 단위 비교 로직
            ' → CompareText & DisplayText

            ' 1) 먼저 "배열_문자열_변경전 / 배열_문자열_변경후" + "어절수" 세팅
            Call PrepareWords

            ' 2) CompareText → 실제 "어절 단위" 비교 로직
            CompareText

            ' 3) DisplayText → 밑줄/취소선 표시
            DisplayText
        End If
    Next col
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
' (보조 Sub) 문자열을 Split → 전역 배열에 담기
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PrepareWords()
    Dim tmpArr As Variant
    Dim cnt As Long
    
    ' --- 변경전 ---
    If Len(문자열_변경전) = 0 Then
        존재여부_문자열_변경전 = False
        어절수_문자열_변경전 = 0
    Else
        tmpArr = Split(문자열_변경전, 구분자)  ' 0-based
        
        cnt = UBound(tmpArr) - LBound(tmpArr) + 1
        If cnt <= 0 Then
            존재여부_문자열_변경전 = False
            어절수_문자열_변경전 = 0
        Else
            존재여부_문자열_변경전 = True
            ReDim 배열_문자열_변경전(1 To cnt)
            Dim i As Long
            For i = 1 To cnt
                배열_문자열_변경전(i) = tmpArr(i - 1)
            Next i
            어절수_문자열_변경전 = cnt
        End If
    End If
    
    ' --- 변경후 ---
    If Len(문자열_변경후) = 0 Then
        존재여부_문자열_변경후 = False
        어절수_문자열_변경후 = 0
    Else
        tmpArr = Split(문자열_변경후, 구분자)  ' 0-based
        
        cnt = UBound(tmpArr) - LBound(tmpArr) + 1
        If cnt <= 0 Then
            존재여부_문자열_변경후 = False
            어절수_문자열_변경후 = 0
        Else
            존재여부_문자열_변경후 = True
            ReDim 배열_문자열_변경후(1 To cnt)
            Dim j As Long
            For j = 1 To cnt
                배열_문자열_변경후(j) = tmpArr(j - 1)
            Next j
            어절수_문자열_변경후 = cnt
        End If
    End If
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
' (핵심1) "어절" 단위로 변경 전/후 비교 → 전역변수에 기록
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CompareText()

    ' (1) 결과 배열(문자열_작업결과)을 미리 최대 길이로 ReDim
    '     = (어절수_문자열_변경전 + 어절수_문자열_변경후)
    ReDim 배열_문자열_작업결과(1 To (어절수_문자열_변경전 + 어절수_문자열_변경후))
    존재여부_문자열_작업결과 = True

    ' (2) 삭제/추가 관련 전역배열 초기화
    존재여부_삭제된문자열 = False
    존재여부_추가된문자열 = False
    어절순번_삭제된문자열 = 0
    어절순번_추가된문자열 = 0

    글자수_작업결과 = 1  ' 결과 문자열에서 "현재까지 몇 글자?" (1부터 시작)
    
    어절순번_문자열_변경전 = 1
    어절순번_문자열_변경후 = 1
    어절순번_문자열_작업결과 = 0

    ' (3) 1차 무한루프: 서로 같거나 다른 어절을 비교
    Do
        Dim 어절_문자열_변경전 As String
        Dim 어절_문자열_변경후 As String
        
        ' 변경전 어절 가져오기
        If (존재여부_문자열_변경전 = False) Or _
           (어절순번_문자열_변경전 > 어절수_문자열_변경전) Then
            어절_문자열_변경전 = ""
        Else
            어절_문자열_변경전 = 배열_문자열_변경전(어절순번_문자열_변경전)
        End If
        
        ' 변경후 어절 가져오기
        If (존재여부_문자열_변경후 = False) Or _
           (어절순번_문자열_변경후 > 어절수_문자열_변경후) Then
            어절_문자열_변경후 = ""
        Else
            어절_문자열_변경후 = 배열_문자열_변경후(어절순번_문자열_변경후)
        End If

        ' (3-1) 동일 어절이면 그대로 결과에 추가
        If LCase(어절_문자열_변경전) = LCase(어절_문자열_변경후) And _
           (어절_문자열_변경전 <> "") Then
            
            어절순번_문자열_작업결과 = 어절순번_문자열_작업결과 + 1
            ReDim Preserve 배열_문자열_작업결과(1 To 어절순번_문자열_작업결과)
            배열_문자열_작업결과(어절순번_문자열_작업결과) = 어절_문자열_변경전
            
            글자수_작업결과 = 글자수_작업결과 + 1 + Len(어절_문자열_변경전)
            
            ' 동일 어절은 후배열에서 비워서 중복 매칭 방지
            배열_문자열_변경후(어절순번_문자열_변경후) = vbNullString
            
            어절순번_문자열_변경전 = 어절순번_문자열_변경전 + 1
            어절순번_문자열_변경후 = 어절순번_문자열_변경후 + 1

        Else
            ' (3-2) 다른 어절이면, "어절_문자열_변경전"이
            '       변경후 배열에 있는지 (MATCH) 확인
            Dim tmpWord As String
            tmpWord = 어절_문자열_변경전
            
            If tmpWord = "" Then
                ' 빈 문자열이라면 그냥 넘어가기
                ' (즉, 변경전이 소진된 상태)
                Exit Do
            End If
            
            일치하는위치 = Application.Match(tmpWord, 배열_문자열_변경후, 0)
            
            If IsError(일치하는위치) Then
                ' 변경후에 없는 단어 -> "삭제된 단어"
                
                어절순번_문자열_작업결과 = 어절순번_문자열_작업결과 + 1
                ReDim Preserve 배열_문자열_작업결과(1 To 어절순번_문자열_작업결과)
                배열_문자열_작업결과(어절순번_문자열_작업결과) = tmpWord
                
                ' 삭제된문자열 배열 ReDim
                If 존재여부_삭제된문자열 = False Then
                    존재여부_삭제된문자열 = True
                    ReDim 배열_삭제된문자열_시작위치(1 To 1)
                    ReDim 배열_삭제된문자열_길이(1 To 1)
                End If
                
                어절순번_삭제된문자열 = 어절순번_삭제된문자열 + 1
                ReDim Preserve 배열_삭제된문자열_시작위치(1 To 어절순번_삭제된문자열)
                ReDim Preserve 배열_삭제된문자열_길이(1 To 어절순번_삭제된문자열)
                
                배열_삭제된문자열_시작위치(어절순번_삭제된문자열) = 글자수_작업결과
                배열_삭제된문자열_길이(어절순번_삭제된문자열) = Len(tmpWord)
                
                글자수_작업결과 = 글자수_작업결과 + 1 + Len(tmpWord)
                
                어절순번_문자열_변경전 = 어절순번_문자열_변경전 + 1
            
            Else
                ' 변경후 배열에 같은 단어가 있다
                ' → 그 앞에 있는 "추가된 단어"들을 모두 추가로 처리
                Do
                    If 어절순번_문자열_변경후 > 어절수_문자열_변경후 Then
                        Exit Do
                    End If
                    
                    If LCase(배열_문자열_변경후(어절순번_문자열_변경후)) = LCase(tmpWord) Then
                        Exit Do
                    End If
                    
                    ' "추가된" 단어 하나를 결과에 넣어줌
                    어절순번_문자열_작업결과 = 어절순번_문자열_작업결과 + 1
                    ReDim Preserve 배열_문자열_작업결과(1 To 어절순번_문자열_작업결과)
                    
                    Dim addWord As String
                    addWord = 배열_문자열_변경후(어절순번_문자열_변경후)
                    배열_문자열_작업결과(어절순번_문자열_작업결과) = addWord

                    ' 추가된문자열 배열 ReDim
                    If 존재여부_추가된문자열 = False Then
                        존재여부_추가된문자열 = True
                        ReDim 배열_추가된문자열_시작위치(1 To 1)
                        ReDim 배열_추가된문자열_길이(1 To 1)
                    End If

                    어절순번_추가된문자열 = 어절순번_추가된문자열 + 1
                    ReDim Preserve 배열_추가된문자열_시작위치(1 To 어절순번_추가된문자열)
                    ReDim Preserve 배열_추가된문자열_길이(1 To 어절순번_추가된문자열)

                    배열_추가된문자열_시작위치(어절순번_추가된문자열) = 글자수_작업결과
                    배열_추가된문자열_길이(어절순번_추가된문자열) = Len(addWord)

                    글자수_작업결과 = 글자수_작업결과 + 1 + Len(addWord)

                    ' 중복 매칭 방지
                    배열_문자열_변경후(어절순번_문자열_변경후) = vbNullString
                    어절순번_문자열_변경후 = 어절순번_문자열_변경후 + 1
                    
                Loop
                
                ' ↑ 루프를 빠져나오면 "현재 bWord == aWord"
                '   실제 동일 단어를 처리할 수도 있으나,
                '   여기서는 "다음 번 CompareText 루프"에서 처리.
                
            End If
        End If

        ' (3-3) 둘 중 하나라도 모든 어절을 소진하면 1차 루프 종료
        If (어절순번_문자열_변경전 > 어절수_문자열_변경전) Or _
           (어절순번_문자열_변경후 > 어절수_문자열_변경후) Then
            Exit Do
        End If
    Loop

    ' (4) 남은 "추가된" 어절, "삭제된" 어절 처리
    '   - 변경후가 남은 경우 (추가)
    Do While 어절순번_문자열_변경후 <= 어절수_문자열_변경후
        Dim remB As String
        remB = 배열_문자열_변경후(어절순번_문자열_변경후)
        If remB <> "" Then
            어절순번_문자열_작업결과 = 어절순번_문자열_작업결과 + 1
            ReDim Preserve 배열_문자열_작업결과(1 To 어절순번_문자열_작업결과)
            배열_문자열_작업결과(어절순번_문자열_작업결과) = remB

            If 존재여부_추가된문자열 = False Then
                존재여부_추가된문자열 = True
                ReDim 배열_추가된문자열_시작위치(1 To 1)
                ReDim 배열_추가된문자열_길이(1 To 1)
            End If
            
            어절순번_추가된문자열 = 어절순번_추가된문자열 + 1
            ReDim Preserve 배열_추가된문자열_시작위치(1 To 어절순번_추가된문자열)
            ReDim Preserve 배열_추가된문자열_길이(1 To 어절순번_추가된문자열)
            
            배열_추가된문자열_시작위치(어절순번_추가된문자열) = 글자수_작업결과
            배열_추가된문자열_길이(어절순번_추가된문자열) = Len(remB)
            
            글자수_작업결과 = 글자수_작업결과 + 1 + Len(remB)
        End If
        어절순번_문자열_변경후 = 어절순번_문자열_변경후 + 1
    Loop

    '   - 변경전이 남은 경우 (삭제)
    Do While 어절순번_문자열_변경전 <= 어절수_문자열_변경전
        Dim remA As String
        remA = 배열_문자열_변경전(어절순번_문자열_변경전)
        If remA <> "" Then
            어절순번_문자열_작업결과 = 어절순번_문자열_작업결과 + 1
            ReDim Preserve 배열_문자열_작업결과(1 To 어절순번_문자열_작업결과)
            배열_문자열_작업결과(어절순번_문자열_작업결과) = remA

            If 존재여부_삭제된문자열 = False Then
                존재여부_삭제된문자열 = True
                ReDim 배열_삭제된문자열_시작위치(1 To 1)
                ReDim 배열_삭제된문자열_길이(1 To 1)
            End If
            
            어절순번_삭제된문자열 = 어절순번_삭제된문자열 + 1
            ReDim Preserve 배열_삭제된문자열_시작위치(1 To 어절순번_삭제된문자열)
            ReDim Preserve 배열_삭제된문자열_길이(1 To 어절순번_삭제된문자열)

            배열_삭제된문자열_시작위치(어절순번_삭제된문자열) = 글자수_작업결과
            배열_삭제된문자열_길이(어절순번_삭제된문자열) = Len(remA)

            글자수_작업결과 = 글자수_작업결과 + 1 + Len(remA)
        End If
        어절순번_문자열_변경전 = 어절순번_문자열_변경전 + 1
    Loop

    ' (5) 최종 결과 문자열 Join
    문자열_작업결과 = Join(배열_문자열_작업결과, 구분자)
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
' (핵심2) CompareText 결과를 셀에 표시 (밑줄/취소선)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub DisplayText()
    Dim i As Long
    
    ' (1) 결과 문자열
    위치_출력.Clear
    If 존재여부_문자열_작업결과 Then
        위치_출력.Value = 문자열_작업결과
    End If

    ' (2) 삭제된 어절(취소선) 표시
    If 존재여부_삭제된문자열 Then
        For i = 1 To UBound(배열_삭제된문자열_시작위치)
            With 위치_출력.Characters( _
                Start:=배열_삭제된문자열_시작위치(i), _
                Length:=배열_삭제된문자열_길이(i) _
            ).Font
                .ColorIndex = 3           ' 빨간색
                .Strikethrough = True
            End With
        Next i
    End If
    
    ' (3) 추가된 어절(밑줄) 표시
    If 존재여부_추가된문자열 Then
        For i = 1 To UBound(배열_추가된문자열_시작위치)
            With 위치_출력.Characters( _
                Start:=배열_추가된문자열_시작위치(i), _
                Length:=배열_추가된문자열_길이(i) _
            ).Font
                .ColorIndex = 14         ' 연두색
                .Underline = True
            End With
        Next i
    End If

    ' (4) 셀 줄바꿈
    위치_출력.WrapText = True
End Sub




