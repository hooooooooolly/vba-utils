Attribute VB_Name = "LombokStyle"
Option Explicit

'/**
' * �T�v:
' *     Lombok��@Data�I�Ȃ��̂𐶐����܂�
' *
' * �g����:
' *     A��ɕϐ���
' *     B��Ɍ^
' *     �𗅗񂵂Ď��s����ƁAoutClm��(�f�t�H���g:E��)��
' *     ���̂܂܃N���X���W���[���ɃR�s�y�ł��镶����𐶐����܂�
' */
Sub Data()
    Dim inRw  As Long
    Dim outRw As Long: outRw = 1
    Const outClm As Long = 5
    
    Dim valName As String
    Dim valType As String
    
    '�t�B�[���h�錾
    For inRw = 1 To Cells(Rows.Count, 1).End(xlUp).Row
        valName = Cells(inRw, 1).Value
        valType = Cells(inRw, 2).Value
        Cells(outRw, outClm).Value = "Private " & valName & "_ As " & valType
        outRw = outRw + 1
    Next inRw
    outRw = outRw + 1
    
    'Getter�錾
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

    'Setter�錾
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

'===�o�̓C���[�W��������========================================
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
'===�o�̓C���[�W�����܂�========================================
