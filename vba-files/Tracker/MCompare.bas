Attribute VB_Name = "MCompare"
'namespace=vba-files\Tracker\
Option Explicit

' ���� �޴� ��ư (��ü ��)
Public Sub ButtonCompareAll(control As IRibbonControl)
    Dim response As VbMsgBoxResult
    response = MsgBox("��ü A��/B�� ������ ������� �����Ͻðڽ��ϱ�?", _
                      vbQuestion + vbYesNo, "���� Ȯ��")
    If response = vbYes Then
        Dim comparer As New CCompare
        comparer.CompareAll
    Else
        MsgBox "������ ����߽��ϴ�.", vbInformation, "���"
    End If
End Sub

' ���� �޴� ��ư (���� ���õ� �ุ ��)
Public Sub ButtonCompareSelection(control As IRibbonControl)
    Dim response As VbMsgBoxResult
    response = MsgBox("���� ���õ� ���� �ุ �����Ͻðڽ��ϱ�?", _
                      vbQuestion + vbYesNo, "���� Ȯ��")
    If response = vbYes Then
        Dim comparer As New CCompare
        comparer.CompareSelection
    Else
        MsgBox "������ ����߽��ϴ�.", vbInformation, "���"
    End If
End Sub

