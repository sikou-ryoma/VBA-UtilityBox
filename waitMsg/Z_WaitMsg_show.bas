Attribute VB_Name = "Z_WaitMsg_show"
Option Explicit

Public Sub waitMsgShow(Optional ByVal seconds As Long = 2)
'----------------------------------------------------------------------------------------------------------
'   �t�H�[��"WaitMsg"��\��
'   Unload�͒��ڍs��
'----------------------------------------------------------------------------------------------------------
    Dim waitUntil As Date
    waitUntil = Now + TimeSerial(0, 0, seconds)
    
    With WaitMsg
        .StartUpPosition = 0
        .Left = 150
        .Top = 100
        .Show vbModeless
    End With
    
    Application.Wait waitUntil

End Sub

Public Sub test()
'   waitMsgShow�̎g�p��

    Call waitMsgShow
    
    '���炩�̏���
    
    Unload WaitMsg      '����Unload����

End Sub
