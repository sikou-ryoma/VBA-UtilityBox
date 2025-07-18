Attribute VB_Name = "Z_LogInit"
Option Explicit

'----------------------------------------------------------------------
'   clsLogger�̃C�j�V�����C�Y�p�p�u���b�N���W���[��
'   ���K�[�̐ݒ��ini�t�@�C������ǂݍ���
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
                MsgBox "���O�t�H���_�̍쐬�Ɏ��s���܂����F" & vbCrLf & .LogFolder & vbCrLf & Err.Description, vbCritical
                Exit Sub
            End If
            On Error GoTo 0
        End If
        
        .Info "[InitializeLogger] ���O�J�n"
        .Info "[InitializeLogger] LogLevel   : " & .LogLevel
        .Info "[InitializeLogger] LogFolder  : " & .LogFolder
        .Info "[InitializeLogger] FilePrefix : " & .FilePrefix
    End With
End Sub

