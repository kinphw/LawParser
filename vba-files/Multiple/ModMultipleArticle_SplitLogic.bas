Attribute VB_Name = "ModMultipleArticle_SplitLogic"
Option Explicit

'namespace=vba-files\Multiple\

'===============================
' (2) ���� �ؽ�Ʈ�� �ɰ��� �ٽ� ����
'===============================
Public Sub SplitMultipleArticlesCell(ByVal targetCell As Range)
    Dim originalText As String
    Dim cleansedText As String

    Dim re As Object  ' Late binding (CreateObject)
    Dim articleMatches As Object ' MatchCollection
    Dim pattern As String
    Dim currentRow As Long, currentCol As Long

    ' ��ȿ�� �˻�
    If targetCell Is Nothing Then Exit Sub
    If IsEmpty(targetCell.Value) Then Exit Sub

    ' [A] ���� �ؽ�Ʈ
    originalText = targetCell.Value

    ' [B] ��ó��
    cleansedText = CleanseAndRemoveLeadingSpaces(originalText)
    cleansedText = RemoveTargetLines(cleansedText)

    ' [C] ���Խ� ����
    ' (��, ��, ���� �� �� �տ����� ã��, ���� ��ȣ �ʼ� + ������ ó��)
    ' pattern = "((?:^|\n)��\d+(?:-\d+)?(?:��|��|��(?:��\d+)?\([^)]*\)))([\s\S]*?)(?=((?:^|\n)��\d+(?:-\d+)?(?:��|��|��(?:��\d+)?\([^)]*\))|$))"
    pattern = "((?:^|\n)��\d+(?:-\d+)?(?:��|��|��(?:��\d+)?(?:\([^)]*\)|\s*����<[^>]+>)))([\s\S]*?)(?=((?:^|\n)��\d+(?:-\d+)?(?:��|��|��(?:��\d+)?(?:\([^)]*\)|\s*����<[^>]+>))|$))"

    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = False
    re.MultiLine = False
    re.pattern = pattern

    ' [D] ��Ī
    Set articleMatches = re.Execute(cleansedText)

    ' [E] ��Ī�� ������ �� ��ü �ؽ�Ʈ �״��
    If articleMatches.Count = 0 Then
        targetCell.Value = cleansedText
        Exit Sub
    End If

    ' [F] ��Ī ����� �� ������ �ɰ��� ����
    currentRow = targetCell.Row
    currentCol = targetCell.Column

    Dim i As Long
    For i = 0 To articleMatches.Count - 1
        Dim oneArticle As String
        oneArticle = articleMatches(i).Value

        ' ��ó�� (�յ� ����/���� ����)
        oneArticle = CleanArticleText(oneArticle)

        If i = 0 Then
            ' ù ��° ����� ���� ����
            targetCell.Value = oneArticle
        Else
            ' �� ��° ���� ����� �Ʒ� �࿡ ���� ����
            Rows(currentRow + 1).Insert Shift:=xlDown
            Cells(currentRow + 1, currentCol).Value = oneArticle
            currentRow = currentRow + 1
        End If
    Next i

End Sub
