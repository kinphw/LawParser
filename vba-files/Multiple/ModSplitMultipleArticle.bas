Attribute VB_Name = "ModSplitMultipleArticle"
'namespace=vba-files\Multiple\
Option Explicit

'namespace=vba-files\Multiple\

'===============================
' (1) �����޴��� �����Ǵ� Public Sub
'===============================
Public Sub ApplySplitMultipleArticle(control As IRibbonControl)

    Dim rng As Range
    Dim oneCell As Range

    ' ���� ���� ���� ��������
    Set rng = Selection
    If rng Is Nothing Then
        MsgBox "���õ� ���� �����ϴ�."
        Exit Sub
    End If

    ' ���õ� �� ���� ���� ���� ó��
    For Each oneCell In rng
        If Not IsEmpty(oneCell) Then
            ' �� SplitMultipleArticlesCell ȣ�� (�ٸ� ���)
            SplitMultipleArticlesCell oneCell
        End If
    Next oneCell

    MsgBox "���� ���� �� ��� ���� ���� '���� �и�'�� �Ϸ�Ǿ����ϴ�."

End Sub
