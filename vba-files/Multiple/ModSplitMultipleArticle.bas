Attribute VB_Name = "ModSplitMultipleArticle"
'namespace=vba-files\Multiple\

Option Explicit

' ������
Public Sub ApplySplitMultipleArticle(control As IRibbonControl)

    Dim rng As Range
    Dim oneCell As Range

    Set rng = Selection
    If rng Is Nothing Then
        MsgBox "���õ� ���� �����ϴ�."
        Exit Sub
    End If

    ' ���õ� �� ���� ���� ���� ó��
    For Each oneCell In rng
        If Not IsEmpty(oneCell) Then
            SplitMultipleArticlesCell oneCell
        End If
    Next oneCell

    MsgBox "���� ���� �� ��� ���� ���� '���� �и�'�� �Ϸ�Ǿ����ϴ�."
End Sub


' Level2

'--------------------------------------------------
' [2] ���� ������ "��XX��(��)" ������
' ���Խ����� �����Ͽ� �� ������ �и�
'--------------------------------------------------
Private Sub SplitMultipleArticlesCell(ByVal targetCell As Range)
    Dim originalText As String
    Dim cleansedText As String

    Dim re As Object ' Late binding (CreateObject)
    'Dim re As New RegExp ' Early binding (���� ���� �ʿ�)

    Dim articleMatches As Object ' MatchCollection
    Dim oneMatch As Object ' Match

    Dim pattern As String
    Dim currentRow As Long
    Dim currentCol As Long

    ' ��ȿ�� �˻�
    If targetCell Is Nothing Then Exit Sub
    If IsEmpty(targetCell.Value) Then Exit Sub

    ' [A] ���� �ؽ�Ʈ ��������
    originalText = targetCell.Value

    ' [B] �ߺ��� ���๮�� ���� (Ŭ��¡)
    'cleansedText = ReplaceConsecutiveLineFeeds(originalText)
    cleansedText = CleanseAndRemoveLeadingSpaces(originalText)

    ' ���� ���� �߰�
    cleansedText = RemoveTargetLines(cleansedText)

    ' [C] ���Խ� ����
    ' - "��\d+��\("�� �����ؼ� ��ȣ �� ���� �ݰ�,
    ' �� �� ������ ����(���� ����)�� �ּҷ� �ݺ�(*?)
    ' ������ �ٽ� "��\d+��\("�� ��Ÿ���ų� ���ڿ� ����($) �ñ��� �� ���
    'pattern = "��\d+��\([^)]*\)([\s\S]*?)(?=��\d+��\(|$)"
    'pattern = "��\d+��(?:��\d+)?\([^)]*\)(]\s\S]*?)(?=��\d+��(?:��\d+)?\(|$)" '����
    pattern = "��\d+��(?:��\d+)?\([^)]*\)([\s\S]*?)(?=��\d+��(?:��\d+)?\(|$)"

    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = False
    're.MultiLine = True 'Debug
    re.MultiLine = False
    re.pattern = pattern

    ' [D] ��Ī ����
    Set articleMatches = re.Execute(cleansedText)

    ' [E] ���� ��Ī�� ���ٸ�(=���� ������ �� ������) �׳� �ؽ�Ʈ�� �����ؼ� ����
    If articleMatches.Count = 0 Then
        targetCell.Value = cleansedText
        Exit Sub
    End If

    ' [F] ù ��° ��Ī�� "���� ��"�� ���,
    ' �������� "�Ʒ� ��"�� ����
    currentRow = targetCell.Row
    currentCol = targetCell.Column

    Dim i As Long
    For i = 0 To articleMatches.Count - 1
        Dim oneArticle As String
        oneArticle = articleMatches(i).Value

        If i = 0 Then
            ' ù ��° Article�� ���� ���� ����
            targetCell.Value = oneArticle
        Else
            ' �� ��°���ʹ� �� ���� �� �� ���� ���
            Rows(currentRow + 1).Insert Shift:=xlDown
            Cells(currentRow + 1, currentCol).Value = oneArticle
            currentRow = currentRow + 1
        End If
    Next i

    'MsgBox "���� �и� �۾��� �Ϸ�Ǿ����ϴ�."

End Sub


'--------------------------------------------------
' [1] ReplaceConsecutiveLineFeeds
' (�ߺ��� �ٹٲ��� �ϳ��� ���)
'--------------------------------------------------
'Private Function ReplaceConsecutiveLineFeeds(ByVal txt As String) As String
'Dim temp As String
'temp = txt
'
'' vbLf���� �������� ó�� (ȯ�濡 ���� vbCrLf�� ���)
'Do While InStr(temp, vbLf & vbLf) > 0
'temp = Replace(temp, vbLf & vbLf, vbLf)
'Loop
'
'ReplaceConsecutiveLineFeeds = temp
'End Function
'
'
'Option Explicit

Private Function CleanseAndRemoveLeadingSpaces(ByVal txt As String) As String
    Dim temp As String
    Dim arrLines As Variant
    Dim i As Long

    ' 1) CRLF -> LF ����
    temp = Replace(txt, vbCrLf, vbLf)

    ' 2) ���ӵ� LF 2�� �̻��� 1����
    Do While InStr(temp, vbLf & vbLf) > 0
        temp = Replace(temp, vbLf & vbLf, vbLf)
    Loop

    ' 3) �� �ٷ� ����
    arrLines = Split(temp, vbLf)

    ' 4) �� ���� �� �� ���� ���� (LTrim)
    For i = LBound(arrLines) To UBound(arrLines)
        arrLines(i) = LTrim(arrLines(i))
    Next i

    ' 5) �ٽ� LF�� ��ħ
    temp = Join(arrLines, vbLf)

    CleanseAndRemoveLeadingSpaces = temp
End Function
'[3] ����ü�赵��ư, ���� �� ����

Private Function RemoveTargetLines(ByVal inputText As String) As String

    Dim temp As String
    Dim arrLines As Variant
    Dim resultLines As Collection
    Dim i As Long

    ' ������ �ܾ���� �迭�� ����
    Dim excludeWords As Variant
    excludeWords = Array("����ü�赵��ư", "����", "���ñ�����ư", "����������Ģ��ư", "��Ȱ���ɹ�ư", "����������Ģ")

    ' (1) �켱 CRLF -> LF �� ����
    temp = Replace(inputText, vbCrLf, vbLf)

    ' (2) ���� �ٷ� Split
    arrLines = Split(temp, vbLf)

    ' (3) ����� ���� Collection ��ü �غ�
    Set resultLines = New Collection

    ' (4) �� �� ��ȸ
    For i = LBound(arrLines) To UBound(arrLines)
        Dim oneLine As String
        oneLine = Trim(arrLines(i))
        
        ' �� ���� �ƴϰ� ���ܴ�� �ܾ ���� ��츸 �߰�
        If oneLine <> "" And Not ContainsAnyWord(oneLine, excludeWords) Then
            resultLines.Add oneLine
        End If
    Next i

    
    Dim outputText As String
    ' Collection�� ���ڿ��� ��ȯ
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

    ' ' (4) �� ���� ��ȸ�ϸ鼭, "���� ���"�� �ƴϸ� resultLines�� �߰�
    ' For i = LBound(arrLines) To UBound(arrLines)
    '     Dim oneLine As String
    '     oneLine = arrLines(i)

    '     ' ����1
    '     ' ���� ���� �� �ϳ���� �� ���� (�߰�X)
    '     ' �������� ���� 1ĭ ����
    '     ' - ����ü�赵��ư
    '     ' - ����
    '     ' - ���ñ�����ư
    '     ' - ��������������ư
    '     ' - ��Ȱ���ɹ�ư

    '     If (oneLine = "����ü�赵��ư ") Or (oneLine = "���� ") Or (oneLine = "���ñ�����ư ") Or (oneLine = "����������Ģ��ư ")  Or (oneLine = "��Ȱ���ɹ�ư ") Then
    '     ' Skip

    '     ' ����2 : "��¼���ư"�� ���ԵǾ� ������ ��ŵ
    '     ElseIf InStr(oneLine, "����ü�赵��ư") > 0 Or InStr(oneLine, "���ñ�����ư") > 0 Or InStr(oneLine, "����������Ģ��ư") > 0 Or InStr(oneLine, "��Ȱ���ɹ�ư") > 0 Then
    '     ' Skip

    '     ' ����3 : �׳� "����"�̶�� ��ŵ
    '     ElseIf (Trim(oneLine) = "����") Then

    '     ' ����4 : ���ʿ� �����̶�� ��ŵ
    '     ElseIf (oneLine = "") Then

    '     Else
    '     ' �� �� ������ resultLines�� �߰�
    '         resultLines.Add oneLine
    '     End If
    ' Next i

    ' ' (5) resultLines�� �ٽ� LF�� �̾����
    ' Dim outputText As String

    ' If resultLines.Count = 0 Then
    '     ' ��� ���ŵǾ� ������� �� ���ڿ�
    '     outputText = ""
    ' Else

    '     Dim lineVal As Variant
    '     For Each lineVal In resultLines

    '         If outputText = "" Then
    '             outputText = lineVal
    '         Else
    '             outputText = outputText & vbLf & lineVal
    '         End If

    '     Next lineVal
    ' End If

    ' ' (6) ��� ����
    ' RemoveTargetLines = outputText

' End Function

' ���ڿ��� �迭�� �ܾ ���ԵǾ� �ִ��� Ȯ���ϴ� �Լ�
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
