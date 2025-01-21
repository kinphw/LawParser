'namespace=vba-files\Single\
Attribute VB_Name = "ModSplitSingleArticle"

Option Explicit

'개별 셀 중 "항"만 쪼개기
Public Sub SplitSingleArticleHang(control As IRibbonControl)

    Dim splitter As ClsSplitSingleArticle
    Set splitter = New ClsSplitSingleArticle

    splitter.bHang = True
    splitter.bHo = False
    splitter.bMok = False

    splitter.SplitCell Selection

End Sub

'개별 셀 중 "항", "호" 만 쪼개기
Public Sub SplitSingleArticleHangHo(control As IRibbonControl)

    Dim splitter As ClsSplitSingleArticle
    Set splitter = New ClsSplitSingleArticle

    splitter.bHang = True
    splitter.bHo = True
    splitter.bMok = False

    splitter.SplitCell Selection

End Sub

'개별 셀 중 "항", "호", "목" 까지 쪼개기
Public Sub SplitSingleArticleHangHoMok(control As IRibbonControl)

    Dim splitter As ClsSplitSingleArticle
    Set splitter = New ClsSplitSingleArticle

    splitter.bHang = True
    splitter.bHo = True
    splitter.bMok = True

    splitter.SplitCell Selection

End Sub