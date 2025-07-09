VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SheetProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'拡大率(デフォルト値: 0)
Private zoomRate_          As Long
'列のスクロール位置(デフォルト値: 1)
Private scrollColumn_      As Long
'行のスクロール位置(デフォルト値: 1)
Private scrollRow_         As Long
'改ページ罫線の表示可否(デフォルト値: False)
Private displayPageBreaks_ As Boolean
'罫線の表示可否(デフォルト値: False)
Private displayGridlines_  As Boolean

'/** コンストラクタ */
Private Sub Class_Initialize()
    zoomRate_ = 0
    scrollColumn_ = 1
    scrollRow_ = 1
    displayPageBreaks_ = False
    displayGridlines_ = False
End Sub

'/** Property */
Property Let zoomRate(zoomRate As Long)
    zoomRate_ = zoomRate
End Property

Property Let scrollColumn(scrollColumn As Long)
    scrollColumn_ = scrollColumn
End Property

Property Let scrollRow(scrollRow As Long)
    scrollRow_ = scrollRow
End Property

Property Let displayPageBreaks(displayPageBreaks As Boolean)
    displayPageBreaks_ = displayPageBreaks
End Property

Property Let displayGridlines(displayGridlines As Boolean)
    displayGridlines_ = displayGridlines
End Property

Property Get zoomRate() As Long
    zoomRate = zoomRate_
End Property

Property Get scrollColumn() As Long
    scrollColumn = scrollColumn_
End Property

Property Get scrollRow() As Long
    scrollRow = scrollRow_
End Property

Property Get displayPageBreaks() As Boolean
    displayPageBreaks = displayPageBreaks_
End Property

Property Get displayGridlines() As Boolean
    displayGridlines = displayGridlines_
End Property
