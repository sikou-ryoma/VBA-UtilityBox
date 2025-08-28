Attribute VB_Name = "Z_LogInit"
Option Explicit

'----------------------------------------------------------------------
'   clsLoggerのイニシャライズ用パブリックモジュール
'   ロガーの設定をiniファイルから読み込む
'----------------------------------------------------------------------

Public logger As clsLogger

Public Sub InitializeLogger(ByVal folderPath As String)

    Const PROC_NAME As String = "[InitializeLogger]"

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
        .Info PROC_NAME & " ログ開始"
        .Info PROC_NAME & " --------------------------------------------------------"
        .Info PROC_NAME & " LogLevel   : " & .LogLevel
        .Info PROC_NAME & " LogFolder  : " & .LogFolder
        .Info PROC_NAME & " FilePrefix : " & .FilePrefix
        .Info PROC_NAME & " --------------------------------------------------------"
    End With
End Sub

