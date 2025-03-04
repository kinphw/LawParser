Attribute VB_Name = "ModMultipleArticle_Utils"
Option Explicit

'namespace=vba-files\Multiple\

'===============================
' (3) ��ó�� �� ��ƿ �Լ� ����
'===============================

Public Function CleanseAndRemoveLeadingSpaces(ByVal txt As String) As String
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

    ' 4) �� ���� �� �� ���� ����
    For i = LBound(arrLines) To UBound(arrLines)
        arrLines(i) = LTrim(arrLines(i))
    Next i

    ' 5) �ٽ� LF�� ��ħ
    temp = Join(arrLines, vbLf)
    CleanseAndRemoveLeadingSpaces = temp
End Function


Public Function RemoveTargetLines(ByVal inputText As String) As String

    Dim temp As String
    Dim arrLines As Variant
    Dim resultLines As Collection
    Dim i As Long

    ' ������ �ܾ���� �迭�� ����
    Dim excludeWords As Variant
    excludeWords = Array("����ü�赵��ư", "����", "���ñ�����ư", "����������Ģ��ư", "��Ȱ���ɹ�ư", "����������Ģ")

    ' (1) CRLF -> LF ����
    temp = Replace(inputText, vbCrLf, vbLf)

    ' (2) ���� �ٷ� Split
    arrLines = Split(temp, vbLf)

    ' (3) ����� ���� Collection ��ü
    Set resultLines = New Collection

    ' (4) �� �� ��ȸ
    For i = LBound(arrLines) To UBound(arrLines)
        Dim oneLine As String
        oneLine = Trim(arrLines(i))
        
        ' �� ���� �ƴϰ� ���ܴ�� �ܾ ���� ��츸
        If oneLine <> "" And Not ContainsAnyWord(oneLine, excludeWords) Then
            resultLines.Add oneLine
        End If
    Next i

    ' (5) ����� ���ڿ���
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

    ' 2) ���� ���� ����
    temp = Trim(temp)

    ' 3) ������ ���ӵ� vbLf ����
    Do While Left(temp, 1) = vbLf
        temp = Mid(temp, 2)
    Loop

    ' 4) ������ ���ӵ� vbLf ����
    Do While Right(temp, 1) = vbLf
        temp = Left(temp, Len(temp) - 1)
    Loop

    CleanArticleText = temp
End Function
