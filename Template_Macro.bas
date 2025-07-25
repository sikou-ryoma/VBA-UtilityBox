Attribute VB_Name = "Template_Macro"
Option Explicit

Public Const VERSION As String = "v1.0.0-beta"
Public Const MACRO_NAME As String = "Template_Macro"

Private FO As New FileOjt
Public dUtil As New DateUtility
Public currentFolderPath

Public Sub template_main()

    currentFolderPath = FO.UpPath(ThisWorkbook.path)
    Call InitializeLogger(currentFolderPath)

    On Error GoTo ErrHandler
    Call waitMsgShow
    
    Call template_MsgBox("処理を中断しますか？")
    
    Unload WaitMsg
    Exit Sub
    
ErrHandler:
    logger.ErrorMsg "[main] エラー発生 : " & Err.Description
    logger.WarnMsg "[main] 処理を中断しました"
    Unload WaitMsg
    MsgBox "エラーが発生しました。" & vbCrLf & "エラーメッセージ : " & Err.Description, vbExclamation, MACRO_NAME
    MsgBox "処理を中断します。", vbExclamation, MACRO_NAME
    
End Sub

Private Sub template_MsgBox(ByVal message As String)
    
    Dim rc As Long
    
    rc = MsgBox(message, vbYesNo + vbQuestion, MACRO_NAME)
    If rc = vbYes Then
        MsgBox "処理を中断します。", vbExclamation, MACRO_NAME
    Else
        MsgBox "処理が完了しました。", vbInformation, MACRO_NAME
    End If

End Sub
