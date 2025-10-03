Attribute VB_Name = "LombokStyle"
Option Explicit

'/**
' * 概要:
' *     Lombokの@Data的なものを生成します
' *
' * 使い方:
' *     A列に変数名
' *     B列に型
' *     を羅列して実行すると、outClm列(デフォルト:E列)に
' *     そのままクラスモジュールにコピペできる文字列を生成します
' */
Sub Data()
    Dim inRw  As Long
    Dim outRw As Long: outRw = 1
    Const outClm As Long = 5
    
    Dim valName As String
    Dim valType As String
    
    'フィールド宣言
    For inRw = 1 To Cells(Rows.Count, 1).End(xlUp).Row
        valName = Cells(inRw, 1).Value
        valType = Cells(inRw, 2).Value
        Cells(outRw, outClm).Value = "Private " & valName & "_ As " & valType
        outRw = outRw + 1
    Next inRw
    outRw = outRw + 1
    
    'Getter宣言
    Cells(outRw, outClm).Value = "''===Getter========================================"
    outRw = outRw + 1
    For inRw = 1 To Cells(Rows.Count, 1).End(xlUp).Row
        valName = Cells(inRw, 1).Value
        valType = Cells(inRw, 2).Value
        Cells(outRw, outClm).Value = "Property Get " & valName & " As " & valType
        outRw = outRw + 1
        Cells(outRw, outClm).Value = "    " & valName & " = " & valName & "_"
        outRw = outRw + 1
        Cells(outRw, outClm).Value = "End Property"
        outRw = outRw + 2
    Next inRw

    'Setter宣言
    Cells(outRw, outClm).Value = "''===Setter========================================"
    outRw = outRw + 1
    For inRw = 1 To Cells(Rows.Count, 1).End(xlUp).Row
        valName = Cells(inRw, 1).Value
        valType = Cells(inRw, 2).Value
        Cells(outRw, outClm).Value = "Property Let " & valName & "(ByVal arg As " & valType & ")"
        outRw = outRw + 1
        Cells(outRw, outClm).Value = "    " & valName & "_ = " & "arg"
        outRw = outRw + 1
        Cells(outRw, outClm).Value = "End Property"
        outRw = outRw + 2
    Next inRw
End Sub

'===出力イメージここから========================================
'Private hoge_ As String
'Private fuga_ As Long
'Private piyo_ As Product
'
''===Getter========================================
'Property Get hoge As String
'    hoge = hoge_
'End Property
'
'Property Get fuga As Long
'    fuga = fuga_
'End Property
'
'Property Get piyo As Product
'    piyo = piyo_
'End Property
'
''===Setter========================================
'Property Let hoge(ByVal arg As String)
'    hoge_ = arg
'End Property
'
'Property Let fuga(ByVal arg As Long)
'    fuga_ = arg
'End Property
'
'Property Let piyo(ByVal arg As Product)
'    piyo_ = arg
'End Property
'===出力イメージここまで========================================
