Attribute VB_Name = "CmnUtils"
Option Explicit

'/** �g�嗦100%�őS�V�[�g�̏��������s��. */
Public Sub initAllSheets()
    Call initAllSheetsWithSpecifiedZoomRate(100)
End Sub

'/**
' * �C�ӂ̊g�嗦�őS�V�[�g�̏��������s��.<br>
' * �f�t�H���g�l: 130%
' */
Public Sub initAllSheetsWithSpecified()
    Dim zoomRate As String: zoomRate = InputBox("�g�嗦��ݒ肵�Ă��������B", "�g�嗦�̐ݒ�", 130)
    If IsNumeric(zoomRate) Then
        Call initAllSheetsWithSpecifiedZoomRate(zoomRate)
    End If
End Sub

'/**
' * �S�V�[�g�ɑ΂���A1��I��.<br>
' * �C�ӂ̊g�嗦�ɐݒ肷��.
' *
' * @param zoomRate �g�嗦
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
