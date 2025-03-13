Attribute VB_Name = "MCompare"
'namespace=vba-files\Tracker\
Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ���뺯���� (��������) - �̸��� ����� ����, ���ξ� ����
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private sDelimiter As String          ' Delimiter

Private sOrig As String               ' ���ڿ� ���� ��
Private sNew As String                ' ���ڿ� ���� ��

Private lWordCountOrig As Long        ' ������ (���� ��)
Private lWordCountNew As Long         ' ������ (���� ��)

Private lRowPos As Long               ' ��ġ(��)
Private lColPos As Long               ' ��ġ(��)
Private lMaxRow As Long               ' ����(��)
Private lMaxCol As Long               ' ����(��)

Private lStartRow As Long             ' �۾� ���� ��

' Range ����
Private rngOrig As Range
Private rngNew As Range
Private rngOutput As Range

' ��û����: Ư����� (���� ��� �� ��)
Private rngOutputSpecial As Range

' ���� ���ڿ� ���
Private sResult As String

' ����/�߰� ���� �ε���
Private arrDelWordStart() As Variant
Private arrDelWordLen() As Variant
Private arrAddWordStart() As Variant
Private arrAddWordLen() As Variant

' ���� ��/�� "����" �迭
Private arrWordsOrig As Variant
Private arrWordsNew As Variant
Private arrWordsResult() As Variant

' ���� ����
Private bHasOrig As Boolean
Private bHasNew As Boolean
Private bHasDelWords As Boolean
Private bHasAddWords As Boolean
Private bHasResult As Boolean

' �ε��� ����
Private lWordIdxOrig As Long
Private lWordIdxNew As Long
Private lWordIdxResult As Long

Private lDelWordIdx As Long
Private lAddWordIdx As Long

Private vMatchPos As Variant

' ��� ���ڿ����� �� ���ڱ��� ����� ����
Private lResultCharCount As Long


'''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ���� �޴����� ȣ��Ǵ� �� ��ư
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ButtonCompareAll(control As IRibbonControl)
    Dim lResponse As VbMsgBoxResult
    lResponse = MsgBox("��ü A��/B�� ������ ������� �����Ͻðڽ��ϱ�?", vbQuestion + vbYesNo, "���� Ȯ��")
    If lResponse = vbYes Then
        CompareMain True
    Else
        MsgBox "������ ����߽��ϴ�.", vbInformation, "���"
    End If
End Sub

Public Sub ButtonCompareSelection(control As IRibbonControl)
    Dim lResponse As VbMsgBoxResult
    lResponse = MsgBox("���� ���õ� ���� �ุ �����Ͻðڽ��ϱ�?", vbQuestion + vbYesNo, "���� Ȯ��")
    If lResponse = vbYes Then
        CompareMain False
    Else
        MsgBox "������ ����߽��ϴ�.", vbInformation, "���"
    End If
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ���� ��ƾ
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CompareMain(Optional ByVal doAll As Boolean = True)
    sDelimiter = " "  ' ���� ������(����)

    If doAll Then
        ' A��/B�� ��ü �������� ������ �� ã��
        Dim lLastRowA As Long, lLastRowB As Long
        lLastRowA = Cells(Rows.Count, 1).End(xlUp).Row  ' A��
        lLastRowB = Cells(Rows.Count, 2).End(xlUp).Row  ' B��
        lMaxRow = Application.WorksheetFunction.Max(lLastRowA, lLastRowB)
        
        lMaxCol = 1       ' A/B ���� ��
        lStartRow = 1     ' �Ϲ������� 1����� ����

        Dim lRow As Long
        For lRow = lStartRow To lMaxRow
            CompareOneRow lRow
        Next lRow
    Else
        ' ���õ� ���� �ุ ����
        lMaxCol = 1
        Dim lSelRow As Long
        lSelRow = Selection.Row
        
        CompareOneRow lSelRow
    End If
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
' "Ư�� �� ��" A��/B�� �� �� C���� ��� ���
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CompareOneRow(ByVal lRowNum As Long)
    Dim lCol As Long
    
    ' lMaxCol = 1 �̹Ƿ�, lCol=1�� �ݺ� (A�� vs B��)
    For lCol = 1 To lMaxCol
        
        ' (1) A�� vs B�� ���ڿ�
        sOrig = Cells(lRowNum, lCol).Value
        sNew = Cells(lRowNum, lCol + lMaxCol).Value  ' lCol=1�̸� +1 �� B��

        ' (2) ��� ����� �� (C��)
        Set rngOrig = Cells(lRowNum, lCol)
        Set rngNew = Cells(lRowNum, lCol + lMaxCol)
        Set rngOutput = Cells(lRowNum, lCol + 2 * lMaxCol)  ' C��

        rngOutput.Clear

        ' (3) �б� ó��
        If Len(sOrig) = 0 And Len(sNew) = 0 Then
            ' �� �� ��������� ȸ�� ���
            rngOutput.Interior.ColorIndex = 15

        ElseIf Len(sOrig) = 0 Then
            ' ���� ���� ����� �� �ű�(����)
            rngOutput.Value = sNew
            With rngOutput.Font
                .ColorIndex = 14      ' ���λ�
                .Underline = True
            End With

        ElseIf Len(sNew) = 0 Then
            ' ���� �ĸ� ����� �� ����(��Ҽ�)
            rngOutput.Value = sOrig
            With rngOutput.Font
                .ColorIndex = 3       ' ������
                .Strikethrough = True
            End With

        Else
            ' �� �� ���� ���� �� ���� ���� ��
            PrepareWords
            CompareText
            DisplayText
        End If
    Next lCol
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
' (����) ���ڿ��� Split �� ���� �迭�� ����
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PrepareWords()
    Dim arrTmp As Variant
    Dim lCount As Long
    
    ' --- ���� �� ---
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
    
    ' --- ���� �� ---
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
' (�ٽ�1) "����" ���� ���� ��/�� �� �� �����迭 ���
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CompareText()

    ' (1) ��� �迭 �̸� �ִ���̷� ReDim
    ReDim arrWordsResult(1 To (lWordCountOrig + lWordCountNew))
    bHasResult = True

    ' (2) ����/�߰� ���� �迭 �ʱ�ȭ
    bHasDelWords = False
    bHasAddWords = False
    lDelWordIdx = 0
    lAddWordIdx = 0

    lResultCharCount = 1
    
    lWordIdxOrig = 1
    lWordIdxNew = 1
    lWordIdxResult = 0

    ' (3) 1�� ����
    Do
        Dim sWordOrig As String
        Dim sWordNew As String
        
        ' ���� �� ����
        If (Not bHasOrig) Or (lWordIdxOrig > lWordCountOrig) Then
            sWordOrig = ""
        Else
            sWordOrig = arrWordsOrig(lWordIdxOrig)
        End If
        
        ' ���� �� ����
        If (Not bHasNew) Or (lWordIdxNew > lWordCountNew) Then
            sWordNew = ""
        Else
            sWordNew = arrWordsNew(lWordIdxNew)
        End If

        ' (3-1) ���� �����̸� ����� �״�� �߰�
        If LCase(sWordOrig) = LCase(sWordNew) And (sWordOrig <> "") Then
            
            lWordIdxResult = lWordIdxResult + 1
            ReDim Preserve arrWordsResult(1 To lWordIdxResult)
            arrWordsResult(lWordIdxResult) = sWordOrig
            
            lResultCharCount = lResultCharCount + 1 + Len(sWordOrig)
            
            ' �� �迭���� �ش� ���� ��� ����
            arrWordsNew(lWordIdxNew) = vbNullString
            
            lWordIdxOrig = lWordIdxOrig + 1
            lWordIdxNew = lWordIdxNew + 1

        Else
            ' (3-2) �ٸ� �����̸�, ���� �Ŀ� �ִ��� Match
            Dim sTmpWord As String
            sTmpWord = sWordOrig
            
            If sTmpWord = "" Then
                ' ���� �� ������ ����
                Exit Do
            End If
            
            vMatchPos = Application.Match(sTmpWord, arrWordsNew, 0)
            
            If IsError(vMatchPos) Then
                ' ���� �Ŀ� ���� �ܾ� �� "������" �ܾ�
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
                ' ���� �� �迭�� ���� �ܾ ����
                ' �� �տ� �߰��� �ܾ���� ó��
                Do
                    If lWordIdxNew > lWordCountNew Then
                        Exit Do
                    End If
                    
                    If LCase(arrWordsNew(lWordIdxNew)) = LCase(sTmpWord) Then
                        Exit Do
                    End If
                    
                    ' �߰��� �ܾ�
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

                    ' �ߺ� ��Ī ����
                    arrWordsNew(lWordIdxNew) = vbNullString
                    lWordIdxNew = lWordIdxNew + 1
                    
                Loop
            End If
        End If

        ' (3-3) �� �� �ϳ��� �����Ǹ� ����
        If (lWordIdxOrig > lWordCountOrig) Or (lWordIdxNew > lWordCountNew) Then
            Exit Do
        End If
    Loop

    ' (4) ���� "�߰���" ���� ó��
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

    ' (5) ���� "������" ���� ó��
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

    ' (6) ���� ��� ���ڿ�
    sResult = Join(arrWordsResult, sDelimiter)
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
' (�ٽ�2) CompareText ����� ���� ǥ�� (����/��Ҽ�)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub DisplayText()
    Dim i As Long
    
    rngOutput.Clear
    If bHasResult Then
        rngOutput.Value = sResult
    End If

    ' ������ ����(��Ҽ�)
    If bHasDelWords Then
        For i = 1 To UBound(arrDelWordStart)
            With rngOutput.Characters(Start:=arrDelWordStart(i), Length:=arrDelWordLen(i)).Font
                .ColorIndex = 3           ' ������
                .Strikethrough = True
            End With
        Next i
    End If
    
    ' �߰��� ����(����)
    If bHasAddWords Then
        For i = 1 To UBound(arrAddWordStart)
            With rngOutput.Characters(Start:=arrAddWordStart(i), Length:=arrAddWordLen(i)).Font
                .ColorIndex = 14         ' ���λ�
                .Underline = True
            End With
        Next i
    End If

    rngOutput.WrapText = True
End Sub

