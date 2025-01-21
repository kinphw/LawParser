'namespace=vba-files\Base\
Attribute VB_Name = "ModBase"
Option Explicit

'���๮�� 2���� 1���� ġȯ
Sub ReplaceCRLF2to1(control As IRibbonControl)

    Debug.Print "���Ӱ��๮�ڸ� 1���� ġȯ�մϴ�.";

    Selection.Replace What:="" & Chr(10) & "" & Chr(10) & "", Replacement:="" & Chr(10) & "", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2


    Debug.Print "����+���๮�ڸ� 1���� ġȯ�մϴ�.";

    Selection.Replace What:="" & Chr(32) & "" & Chr(10) & "", Replacement:="" & Chr(10) & "", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2

End Sub

' �� �� ���� �ڵ�����
Sub AllRowsAutoFit(control As IRibbonControl)
    'ActiveSheet.UsedRange.EntireRow.AutoFit
    Selection.EntireRow.AutoFit
End Sub

'�� ���� ���̸� 10 ����
Sub IncreaseRowHeightBy10(control As IRibbonControl)

    If TypeName(Selection) = "Range" Then

        Selection.Rows(1).RowHeight = WorksheetFunction.Min(Selection.Rows(1).RowHeight + 10, 409.5) '409.5�� �ִ���

    End If

End Sub


' ���������� ���� ��(������) ����
Sub SetLawColumn(control As IRibbonControl)

    Selection.EntireColumn.Select

    '�� ����
    Selection.ColumnWidth = 80

    With Selection
        .HorizontalAlignment = xlLeft '�����������
        .VerticalAlignment = xlTop '�����������
        .WrapText = True '�ؽ�Ʈ��ġ��
    End With

    ActiveWindow.Zoom = 80

End Sub

Sub WhoAlert(control As IRibbonControl)
    
    MsgBox "���� : �ں�" & vbCrLf & _
    "���� : ���������ý��� �����Ľ�" & vbCrLf & _        
    "���� : 0.0.1 (250121)" & vbCrLf & _        
    "���� : kinphw@naver.com" & vbCrLf

End Sub