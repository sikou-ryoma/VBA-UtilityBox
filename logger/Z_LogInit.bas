Attribute VB_Name = "Z_LogInit"
Option Explicit

'----------------------------------------------------------------------
'   clsLoggerのイニシャライズ用パブリックモジュール
'   ロガーの設定をiniファイルから読み込む
'----------------------------------------------------------------------

Public logger As clsLogger

Public Sub InitializeLogger(ByVal folderPath As String)
    Dim iniPath As String
    
    iniPath = folderPath & "\config\log_config.ini"
    
    Set logger = New clsLogger
    With logger
        .LogLevel = ReadIniValue("Logger", "LogLevel", "INFO", iniPath)
        .LogFolder = folderPath & "\" & ReadIniValue("Logger", "LogFolder", "log", iniPath)
        .FilePrefix = ReadIniValue("Logger", "FilePrefix", "log", iniPath)
        
        If Dir(.LogFolder, vbDirectory) = "" Then
            On Error Resume Next
            MkDir .LogFolder
            If Err.Number <> 0 Then
                MsgBox "ログフォルダの作成に失敗しました：" & vbCrLf & .LogFolder & vbCrLf & Err.Description, vbCritical
                Exit Sub
            End If
            On Error GoTo 0
        End If
        
        .Info "[InitializeLogger] ログ開始"
        .Info "[InitializeLogger] LogLevel   : " & .LogLevel
        .Info "[InitializeLogger] LogFolder  : " & .LogFolder
        .Info "[InitializeLogger] FilePrefix : " & .FilePrefix
    End With
End Sub

