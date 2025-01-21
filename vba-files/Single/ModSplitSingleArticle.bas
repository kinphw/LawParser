'namespace=vba-files\Single\
Attribute VB_Name = "ModSplitSingleArticle"

Option Explicit

'���� �� �� "��"�� �ɰ���
Public Sub SplitSingleArticleHang(control As IRibbonControl)

    Dim splitter As ClsSplitSingleArticle
    Set splitter = New ClsSplitSingleArticle

    splitter.bHang = True
    splitter.bHo = False
    splitter.bMok = False

    splitter.SplitCell Selection

End Sub

'���� �� �� "��", "ȣ" �� �ɰ���
Public Sub SplitSingleArticleHangHo(control As IRibbonControl)

    Dim splitter As ClsSplitSingleArticle
    Set splitter = New ClsSplitSingleArticle

    splitter.bHang = True
    splitter.bHo = True
    splitter.bMok = False

    splitter.SplitCell Selection

End Sub

'���� �� �� "��", "ȣ", "��" ���� �ɰ���
Public Sub SplitSingleArticleHangHoMok(control As IRibbonControl)

    Dim splitter As ClsSplitSingleArticle
    Set splitter = New ClsSplitSingleArticle

    splitter.bHang = True
    splitter.bHo = True
    splitter.bMok = True

    splitter.SplitCell Selection

End Sub