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
    
    Call template_MsgBox("�����𒆒f���܂����H")
    
    Unload WaitMsg
    Exit Sub
    
ErrHandler:
    logger.ErrorMsg "[main] �G���[���� : " & Err.Description
    logger.WarnMsg "[main] �����𒆒f���܂���"
    Unload WaitMsg
    MsgBox "�G���[���������܂����B" & vbCrLf & "�G���[���b�Z�[�W : " & Err.Description, vbExclamation, MACRO_NAME
    MsgBox "�����𒆒f���܂��B", vbExclamation, MACRO_NAME
    
End Sub

Private Sub template_MsgBox(ByVal message As String)
    
    Dim rc As Long
    
    rc = MsgBox(message, vbYesNo + vbQuestion, MACRO_NAME)
    If rc = vbYes Then
        MsgBox "�����𒆒f���܂��B", vbExclamation, MACRO_NAME
    Else
        MsgBox "�������������܂����B", vbInformation, MACRO_NAME
    End If

End Sub
