VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} WaitMsg 
   Caption         =   "マクロの実行"
   ClientHeight    =   1680
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6330
   OleObjectBlob   =   "WaitMsg.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "WaitMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub UserForm_Initialize()
    WaitMsg.Caption = MACRO_NAME & "_" & VERSION
End Sub
