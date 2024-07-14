Attribute VB_Name = "CmnUtils"
Option Explicit

'/** 拡大率100%で全シートの初期化を行う. */
Public Sub initAllSheets()
    Call initAllSheetsWithSpecifiedZoomRate(100)
End Sub

'/**
' * 任意の拡大率で全シートの初期化を行う.<br>
' * デフォルト値: 130%
' */
Public Sub initAllSheetsWithSpecified()
    Dim zoomRate As String: zoomRate = InputBox("拡大率を設定してください。", "拡大率の設定", 130)
    If IsNumeric(zoomRate) Then
        Call initAllSheetsWithSpecifiedZoomRate(zoomRate)
    End If
End Sub

'/**
' * 全シートに対してA1を選択.<br>
' * 任意の拡大率に設定する.
' *
' * @param zoomRate 拡大率
' */
Private Sub initAllSheetsWithSpecifiedZoomRate(ByVal zoomRate As Long)
    Dim shtNum As Long
    For shtNum = Sheets.Count To 1 Step -1
        With Sheets(shtNum)
            If .Visible Then
                .Select
                .Cells(1, 1).Select
                With ActiveWindow
                    .ScrollColumn = 1
                    .ScrollRow = 1
                    .Zoom = zoomRate
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
