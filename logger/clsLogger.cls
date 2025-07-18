VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pLogLevel As String
Private pLogFolder As String
Private pFilePrefix As String
Private Const DEFAULT_LEVEL As String = "INFO"

'--- プロパティ：ログレベル
Public Property Let LogLevel(ByVal level As String)
    pLogLevel = UCase(level)
End Property

Public Property Get LogLevel() As String
    LogLevel = pLogLevel
End Property

'--- プロパティ：ログフォルダ
Public Property Let LogFolder(ByVal folder As String)
    pLogFolder = folder
End Property

Public Property Get LogFolder() As String
    LogFolder = pLogFolder
End Property

'--- プロパティ：ファイル名の接頭辞
Public Property Let FilePrefix(ByVal prefix As String)
    pFilePrefix = prefix
End Property

Public Property Get FilePrefix() As String
    FilePrefix = pFilePrefix
End Property

'--- 初期化（コンストラクタ）
Private Sub class_initialize()
    pLogFolder = ThisWorkbook.path
    pFilePrefix = "log"
    pLogLevel = "INFO"
End Sub

'--- ログ出力（INFO）
Public Sub Info(ByVal message As String)
    WriteLog "INFO", message
End Sub

'--- ログ出力（DEBUG）
Public Sub DebugMsg(ByVal message As String)
    WriteLog "DEBUG", message
End Sub

'--- ログ出力（WARN）
Public Sub WarnMsg(ByVal message As String)
    WriteLog "WARN", message
End Sub

'--- ログ出力（ERROR）
Public Sub ErrorMsg(ByVal message As String)
    WriteLog "ERROR", message
End Sub

'--- 共通ログ出力ロジック
Private Sub WriteLog(ByVal level As String, ByVal message As String)
    If Not ShouldWrite(level) Then Exit Sub
    
    Dim logPath As String
    Dim fileNum As Integer
    Dim timeStamp As String
    Dim logLine As String

    If Right(pLogFolder, 1) <> "\" Then
        pLogFolder = pLogFolder & "\"
    End If
    
    logPath = pLogFolder & pFilePrefix & "_" & Format(Date, "yyyymmdd") & ".txt"
    timeStamp = Format(Now, "yyyy/mm/dd hh:nn:ss")
    logLine = timeStamp & " [" & UCase(level) & "] " & message
    
    fileNum = FreeFile
    Open logPath For Append As #fileNum
        Print #fileNum, logLine
    Close #fileNum
End Sub

Private Function ShouldWrite(ByVal level As String) As Boolean
    Dim levelPriority As Object
    Set levelPriority = CreateObject("Scripting.Dictionary")
    
    levelPriority.Add "DEBUG", 1
    levelPriority.Add "INFO", 2
    levelPriority.Add "WARN", 3
    levelPriority.Add "ERROR", 4

    If levelPriority.Exists(UCase(level)) And levelPriority.Exists(UCase(pLogLevel)) Then
        ShouldWrite = levelPriority(UCase(level)) >= levelPriority(UCase(pLogLevel))
    Else
        ShouldWrite = False
    End If
End Function

