Attribute VB_Name = "Template_Macro"
Option Explicit


'���W���[�����x���萔
'------------------------------------------------------------------------
Public Const VERSION As String = "v1.0.0-beta"
Public Const MACRO_NAME As String = "Template_Macro"


'���W���[�����x���ϐ�
'------------------------------------------------------------------------
Private FO As New FileOjt
Private bm As New BookManager
Private du As New DateUtility
Private Paths As New PathConfig



Public Sub MainContriller()

    '�ݒ�
    '------------------------------------------------------------------------
    Const PROC_NAME As String = "[MainContriller]"
    
    Application.ScreenUpdating = False
    
    Paths.Init FO.UpPath(ThisWorkbook.path)
    Call InitializeLogger(Paths.ProjectRoot)
    Paths.SetPath "test", Paths.ProjectRoot & "\test"
    Paths.SetPath "documents", Environ("USERPROFILE") & "\Documents"
    
    logger.Info PROC_NAME & " macro name    : " & MACRO_NAME
    logger.Info PROC_NAME & " macro version : " & VERSION
    logger.DebugMsg PROC_NAME & " ProjectRoot : " & Paths.ProjectRoot
    logger.DebugMsg PROC_NAME & " TempPath    : " & Paths.TempPath
    logger.DebugMsg PROC_NAME & " ReportsPath : " & Paths.ReportsPath
    
    logger.Info PROC_NAME & " �����̊J�n"
    
    ChDrive Left(Paths.GetPath("test"), 1)
    ChDir Paths.GetPath("test")
    
    On Error GoTo ErrHandler
        
    
    '�O����
    '------------------------------------------------------------------------
    Call waitMsgShow
    
    
    '�{����
    '------------------------------------------------------------------------
    Call TempMsgBox("�����𒆒f���܂����H")
    
    
    '�㏈��
    '------------------------------------------------------------------------
    Unload WaitMsg
    MsgBox "�������������܂����B", vbInformation, MACRO_NAME
    logger.Info PROC_NAME & " ����I��"
    Application.ScreenUpdating = True
    
    Exit Sub
    


ErrHandler:
    
    '�G���[����
    '------------------------------------------------------------------------
    logger.ErrorMsg PROC_NAME & " �G���[���� : " & Err.Description
    logger.WarnMsg PROC_NAME & " �����𒆒f���܂���"
    Unload WaitMsg
    MsgBox "�G���[���������܂����B" & vbCrLf & "�G���[���b�Z�[�W : " & Err.Description, vbExclamation, MACRO_NAME
    MsgBox "�����𒆒f���܂��B", vbExclamation, MACRO_NAME
    
End Sub

Private Sub TempMsgBox(ByVal message As String)
    
    Dim rc As Long
    
    rc = MsgBox(message, vbYesNo + vbQuestion, MACRO_NAME)
    If rc = vbYes Then
        MsgBox "�����𒆒f���܂��B", vbExclamation, MACRO_NAME
    Else
        MsgBox "�������������܂����B", vbInformation, MACRO_NAME
    End If

End Sub
