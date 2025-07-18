Attribute VB_Name = "Z_WaitMsg_show"
Option Explicit

Public Sub waitMsgShow(Optional ByVal seconds As Long = 2)
'----------------------------------------------------------------------------------------------------------
'   フォーム"WaitMsg"を表示
'   Unloadは直接行う
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
'   waitMsgShowの使用例

    Call waitMsgShow
    
    '何らかの処理
    
    Unload WaitMsg      '直接Unloadする

End Sub
