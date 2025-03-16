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
    
    ' 1) [����]����.[����] => ^\s*\d+\.\s*
    If removeNumberDot Then
        ReDim Preserve arrPatterns(n)
        arrPatterns(n) = "\d+\.\s*"
        n = n + 1
    End If
    
    ' 2) [����]����)[����] => ^\s*\d+\)\s*
    If removeNumberParen Then
        ReDim Preserve arrPatterns(n)
        arrPatterns(n) = "\d+\)\s*"
        n = n + 1
    End If
    
    ' 3) [����][�����ٶ󸶹ٻ������īŸ����])[����] => ^\s*[�����ٶ󸶹ٻ������īŸ����]\)\s*
    If removeKoreanParen Then
        ReDim Preserve arrPatterns(n)
        arrPatterns(n) = "[�����ٶ󸶹ٻ������īŸ����]\)\s*"
        n = n + 1
    End If

    ' 4) ��~�� (�����ڵ� ���� ����) ���� => ^\s*[\u2460-\u2468]\s*
    If removeCircledNumbers Then
        ReDim Preserve arrPatterns(n)
        arrPatterns(n) = "[��-��]\s*"
        n = n + 1
    End If
    
    ' ������ ������ ���� ���õ��� ���� ���, ���� �״�� ����
    If n = 0 Then
        RemoveLeadingPatterns = s
        Exit Function
    End If
    
    ' ����ǥ���� ���� �����
    ' ^\s* (?: ����1 | ����2 | ����3 | ����4 )
    Dim finalPattern As String
    finalPattern = "^\s*(?:" & Join(arrPatterns, "|") & ")"
    
    ' ����ǥ���� ��ü ����
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    With re
        .pattern = finalPattern
        .Global = False      ' ���� �κ� 1ȸ�� ġȯ
        .IgnoreCase = True
    End With
    
    ' ġȯ ����
    RemoveLeadingPatterns = re.Replace(s, "")
End Function

'���ÿ��� Remove
Sub RemoveSelection(control As IRibbonControl)

    Dim cell As Range
    
    ' ���õ� ��� ���� ��ȯ�ϸ鼭 LTrim ����
    For Each cell In Selection
        ' ���� ������� ���� ��츸 ó��
        If Not IsEmpty(cell.Value) Then
            cell.Value = RemoveLeadingPatterns(cell.Value)
        End If
    Next cell
    
    MsgBox "��ȣó�� ���� �Ϸ�!", vbInformation, "��ȣó�� ���� �Ϸ�!"
    
End Sub


' ���ÿ��� LTrim
Sub TrimLeftSelection(control As IRibbonControl)
    Dim cell As Range
    
    ' ���õ� ��� ���� ��ȯ�ϸ鼭 LTrim ����
    For Each cell In Selection
        ' ���� ������� ���� ��츸 ó��
        If Not IsEmpty(cell.Value) Then
            cell.Value = LTrim(cell.Value)
        End If
    Next cell
    
    MsgBox "���� ���� ���� �Ϸ�!", vbInformation, "LTrim ���� �Ϸ�"
End Sub

