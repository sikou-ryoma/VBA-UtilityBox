Attribute VB_Name = "Template_Macro"
Option Explicit


'モジュールレベル定数
'------------------------------------------------------------------------
Public Const VERSION As String = "v1.0.0-beta"
Public Const MACRO_NAME As String = "Template_Macro"


'モジュールレベル変数
'------------------------------------------------------------------------
Private FO As New FileOjt
Private bm As New BookManager
Private dUtil As New DateUtility

Private gProjectFolderPath As String
Private gTempFolderPath As String
Private gReportFolderPath As String
Private gCurrentFolderPath As String


Public Sub MainContriller()

    '設定
    '------------------------------------------------------------------------
    Const PROC_NAME As String = "[MainContriller]"
    
    Application.ScreenUpdating = False
    
    gProjectFolderPath = FO.UpPath(ThisWorkbook.path)
    Call InitializeLogger(gProjectFolderPath)
    gTempFolderPath = gProjectFolderPath & "\output\temp"
    gReportFolderPath = gProjectFolderPath & "\output\reports"
    'gCurrentFolderPath = Environ("USERPROFILE") & "\Documents"
    gCurrentFolderPath = gProjectFolderPath & "\test" 'デバッグ用パス
    
    logger.Info PROC_NAME & " macro name    : " & MACRO_NAME
    logger.Info PROC_NAME & " macro version : " & VERSION
    logger.DebugMsg PROC_NAME & " projectFolderPath : " & gProjectFolderPath
    logger.DebugMsg PROC_NAME & " tempFolderPath    : " & gTempFolderPath
    logger.DebugMsg PROC_NAME & " reportFolderPath  : " & gReportFolderPath
    
    logger.Info PROC_NAME & " 処理の開始"
    
    ChDrive Left(gCurrentFolderPath, 1)
    ChDir gCurrentFolderPath
    
    On Error GoTo ErrHandler
        
    
    '前処理
    '------------------------------------------------------------------------
    Call waitMsgShow
    
    
    '本処理
    '------------------------------------------------------------------------
    Call TempMsgBox("処理を中断しますか？")
    
    
    '後処理
    '------------------------------------------------------------------------
    Unload WaitMsg
    MsgBox "処理が完了しました。", vbInformation, MACRO_NAME
    logger.Info PROC_NAME & " 正常終了"
    Application.ScreenUpdating = True
    
    Exit Sub
    


ErrHandler:
    
    'エラー処理
    '------------------------------------------------------------------------
    logger.ErrorMsg PROC_NAME & " エラー発生 : " & Err.Description
    logger.WarnMsg PROC_NAME & " 処理を中断しました"
    Unload WaitMsg
    MsgBox "エラーが発生しました。" & vbCrLf & "エラーメッセージ : " & Err.Description, vbExclamation, MACRO_NAME
    MsgBox "処理を中断します。", vbExclamation, MACRO_NAME
    
End Sub

Private Sub TempMsgBox(ByVal message As String)
    
    Dim rc As Long
    
    rc = MsgBox(message, vbYesNo + vbQuestion, MACRO_NAME)
    If rc = vbYes Then
        MsgBox "処理を中断します。", vbExclamation, MACRO_NAME
    Else
        MsgBox "処理が完了しました。", vbInformation, MACRO_NAME
    End If

End Sub
