Attribute VB_Name = "modConfig"
Option Explicit

'--- INIì«Ç›éÊÇËópAPIêÈåæ
Private Declare PtrSafe Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
    (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, _
    ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Public Function ReadIniValue(ByVal section As String, ByVal key As String, ByVal defaultVal As String, ByVal iniPath As String) As String
    Dim buffer As String * 255
    Dim ret As Long
    ret = GetPrivateProfileString(section, key, defaultVal, buffer, Len(buffer), iniPath)
    ReadIniValue = Trim(Left(buffer, ret))
End Function


