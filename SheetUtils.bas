Attribute VB_Name = "SheetUtils"
Option Explicit

'/**
' * 全シートの初期化を行う.
' *
' * @param sp シート設定値を保持するクラス
' */
Public Sub initAllSheets(ByRef sp As SheetProperties)
    Dim shtNum As Long
    For shtNum = Sheets.Count To 1 Step -1
        With Sheets(shtNum)
            '非表示シートは処理しない
            If .Visible Then
                .Select
                .Cells(1, 1).Select
                '改ページ罫線
                .displayPageBreaks = sp.displayPageBreaks
                'ウィンドウサイズ
                With ActiveWindow
                    'スクロール位置
                    .scrollColumn = sp.scrollColumn
                    .scrollRow = sp.scrollRow
                    '枠線
                    .DisplayGridlines = sp.displayGridlines
                    '指定した時のみ拡大率を変更する
                    If Not (sp.zoomRate = 0) Then
                        .Zoom = sp.zoomRate
                    End If
                End With
            End If
        End With
    Next shtNum
End Sub

'/** 非表示になっている名前定義を表示する. */
Public Sub dispNameDefinition()
    Dim n As Object
    For Each n In Names
        n.Visible = True
    Next
End Sub
