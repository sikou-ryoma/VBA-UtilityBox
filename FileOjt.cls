VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileOjt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'ファイル選択や保存関する操作を簡潔に行うためのユーティリティクラス

Private pFileNm As Variant 'select file path
Private pFolderPath As Variant 'select folder path

Public Property Let FileNm(ByVal fn As Variant)
    pFileNm = fn
End Property

Public Property Get FileNm() As Variant
    FileNm = pFileNm
End Property

Public Property Let folderPath(ByVal fp As Variant)
    pFolderPath = fp
End Property

Public Property Get folderPath() As Variant
    folderPath = pFolderPath
End Property

Public Function OpFile() As Variant
'-----------------------------------------------------
'   引数指定はなし
'   この関数は選択したファイルのフルパスを返してそのファイルを開き、キャンセル時はFalseを返す
'   キャンセル時に付いては呼び出し元で分岐を管理する必要がある
'-----------------------------------------------------
    Dim fileName As Variant
    MsgBox "ブックを選択してください。", vbInformation, "OpenFile"
    fileName = Application.GetOpenFilename(FileFilter:="Excelファイル,*.xlsx,Excel2003以前,*.xls,Excelマクロブック,*.xlsm")
    OpFile = fileName
    If fileName = False Then
        MsgBox "キャンセルが押されたため処理を中断します。", vbExclamation, "OpenFile"
        Exit Function
    Else
        Workbooks.Open fileName
    End If
    
End Function

Public Function SvFile(ByVal rcvFn As String) As Variant

    Dim fileName As Variant
    Dim FF As Long

    ' 拡張子によるフィルターインデックスの設定
    If rcvFn Like "*.xlsx" Then
        FF = 1
    ElseIf rcvFn Like "*.xls" Then
        FF = 2
    ElseIf rcvFn Like "*.xlsm" Then
        FF = 3
    Else
        FF = 1
    End If

    MsgBox "保存先を選択してください。", vbInformation, "SaveFile"

    On Error GoTo ErrHandler

    fileName = Application.GetSaveAsFilename( _
        InitialFileName:=rcvFn, _
        FileFilter:="Excelファイル,*.xlsx,Excel2003以前,*.xls,Excelマクロブック,*.xlsm", _
        FilterIndex:=FF)

    If fileName = False Then
        MsgBox "キャンセルが押されました。", vbExclamation, "SaveFile"
        SvFile = False
        Exit Function
    End If

    ' 現在のブックを保存（必要に応じてThisWorkbookかActiveWorkbookを使い分けてください）
    ActiveWorkbook.SaveAs fileName
    SvFile = fileName
    Exit Function

ErrHandler:
    SvFile = False
    MsgBox "保存に失敗しました。ファイル形式やファイル名を確認してください。" & vbCrLf & _
           "エラーコード：" & Err.Number & vbCrLf & "エラー内容：" & Err.Description, vbExclamation, "SaveFile"
    Err.Clear

End Function

Public Property Get GetFolderPath() As Variant
'-----------------------------------------------------
'   引数指定はなし
'   この関数は選択したフォルダのパスを返す、キャンセル時はFalseを返す
'   キャンセル時に付いては呼び出し元で分岐を管理する必要がある
'-----------------------------------------------------
    MsgBox "指定するフォルダを選択して下さい。", vbInformation, "GetFolderPath"
    With Application.FileDialog(msoFileDialogFolderPicker)
        If Not .Show Then
            GetFolderPath = False
            MsgBox "キャンセルが押されました。", vbExclamation, "GetFolderPath"
            Exit Property
        End If
        GetFolderPath = .SelectedItems(1)
    End With

End Property

'--- 最終的な \ を除去して返す
Public Function NormalizePath(ByVal path As String) As String

    If Right(path, 1) = "\" Then
        NormalizePath = Left(path, Len(path) - 1)
    Else
        NormalizePath = path
    End If
    
End Function

'--- 一つ上のフォルダのパスを取得
Public Function ParentPath(ByVal fullPath As String) As String

    fullPath = NormalizePath(fullPath)
    
    Dim lastSep As Long
    lastSep = InStrRev(fullPath, "\")
    
    If lastSep > 0 Then
        ParentPath = Left(fullPath, lastSep - 1)
    Else
        ParentPath = fullPath ' パス区切りがない（例外）
    End If
    
End Function

'--- 指定された階層数だけ上のフォルダパスを返す
Public Function UpPath(ByVal fullPath As String, Optional ByVal levelUp As Long = 1) As String

    Dim i As Long
    
    fullPath = NormalizePath(fullPath)
    
    For i = 1 To levelUp
        fullPath = ParentPath(fullPath)
    Next i
    
    UpPath = fullPath
End Function

