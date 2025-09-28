VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum WriteMode
    MODE_READ = 1
    MODE_WRITE = 2
    MODE_APPEND = 8
End Enum

Private fso      As Object
Private ts       As Object
Private filePath As String
Private isOpen   As Boolean

' 初期化（コンストラクタ代わり）
Public Sub init(ByVal path As String)
    Set fso = CreateObject("Scripting.FileSystemObject")
    filePath = path
    isOpen = False
End Sub

' ファイルを開く
Public Sub openFile(Optional ByVal mode As WriteMode = MODE_READ)
    If isOpen Then closeFile
    Set ts = fso.OpenTextFile(filePath, mode, True)
    isOpen = True
End Sub

' ファイルを閉じる
Public Sub closeFile()
    If isOpen Then
        ts.Close
        Set ts = Nothing
        isOpen = False
    End If
End Sub

' ファイル名取得
Public Function getFileName() As String
    getFileName = fso.getFileName(filePath)
End Function

' ファイル存在チェック
Public Function exists() As Boolean
    exists = fso.FileExists(filePath)
End Function

' ファイル削除
Public Sub delete()
    If exists Then
        If isOpen Then closeFile
        fso.DeleteFile filePath, True
    End If
End Sub

' ファイル全体を読み込む
Public Function readAll() As String
    If Not ts Is Nothing Then
        readAll = ts.readAll
    Else
        Err.Raise vbObjectError + 1000, "FileHandler", "ファイルが開かれていません。"
    End If
End Function

' 1行ずつ読み込む
Public Function readLine() As String
    If Not ts Is Nothing Then
        If Not ts.AtEndOfStream Then
            readLine = ts.readLine
        Else
            readLine = vbNullString
        End If
    Else
        Err.Raise vbObjectError + 1003, "FileHandler", "ファイルが開かれていません。"
    End If
End Function

' 書き込み（改行なし）
Public Sub writeRaw(ByVal text As String)
    If Not ts Is Nothing Then
        ts.WRITE text
    Else
        Err.Raise vbObjectError + 1001, "FileHandler", "ファイルが開かれていません。"
    End If
End Sub

' 書き込み（改行あり）
Public Sub writeLine(ByVal text As String)
    If Not ts Is Nothing Then
        ts.writeLine text
    Else
        Err.Raise vbObjectError + 1002, "FileHandler", "ファイルが開かれていません。"
    End If
End Sub

' 追記モードで書き込み
Public Sub appendLine(ByVal text As String)
    Dim tsAppend As Object
    Set tsAppend = fso.OpenTextFile(filePath, WriteMode.MODE_APPEND, True)
    tsAppend.writeLine text
    tsAppend.Close
End Sub
