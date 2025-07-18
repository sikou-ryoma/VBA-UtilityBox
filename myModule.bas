Attribute VB_Name = "myModule"
Option Explicit

'---------------------------------------------
'   �V���[�g�J�b�g�L�[��ݒ肵�Ă���̂ł��D�݂ŔC�ӂ̃L�[��ݒ肵�Ă�������
'   �V���[�g�J�b�g�L�[��Ctrl + Shift + �C�ӂ̃L�[
'----------------------------------------------

Sub home()
Attribute home.VB_ProcData.VB_Invoke_Func = "E\n14"
'---------------------------------------------
'   �}�N���N�����ɃA�N�e�B�u�ȃu�b�N�̑S�V�[�g�̑I���Z����A1�ɖ߂�
'   �A�N�e�B�u�V�[�g�͍ŏI�V�[�g�ɂȂ�
'   E
'---------------------------------------------
    Dim wb As Workbook, wsLoop As Worksheet
    
    Set wb = ActiveWorkbook
    For Each wsLoop In wb.Worksheets
        wsLoop.Activate
        ActiveWindow.ScrollRow = 1
        ActiveWindow.ScrollColumn = 1
        Range("A1").Select
    Next wsLoop
    
End Sub

Sub sheetslist()
Attribute sheetslist.VB_ProcData.VB_Invoke_Func = "L\n14"
'---------------------------------------------
'   �N�����ɊJ���Ă���S�Ẵu�b�N�̃p�X�A�u�b�N���A�V�[�g�����擾���Ĉꗗ�ɂ���
'   �����N���ꏏ�ɍ쐬����A�A�N�Z�X�\
'   L
'---------------------------------------------
    Dim logBk As Workbook, logSh As New SheetInfo
    Dim wbCnt As Variant, wsCnt As Variant
    Dim i As Long, j As Long, SaveAsFilePath As String
    
    SaveAsFilePath = "C:\Path\BooksLog" ' �ۑ���̃p�X���w�肵�Ă�������
    Set logBk = Workbooks.Add
    Set logSh.sh = logBk.Sheets(1)
    With logSh.sh
        .Cells(1, A__).Value = "path"
        .Cells(1, B__).Value = "book"
        .Cells(1, C__).Value = "sheet"
        .Cells(1, D__).Value = "flag"
    End With
    
    For i = 1 To Workbooks.Count - 1
        If Workbooks(i).Name <> ThisWorkbook.Name Then
            For j = 1 To Workbooks(i).Worksheets.Count
                logSh.ER = logSh.EndRow()
                With logSh.sh
                    .Cells(logSh.ER + 1, A__).Value = Workbooks(i).Path
                    .Cells(logSh.ER + 1, B__).Value = Workbooks(i).Name
                    .Hyperlinks.Add anchor:=Cells(logSh.ER + 1, C__), _
                        Address:=Workbooks(i).Path & "\" & Workbooks(i).Name, _
                        SubAddress:="'" & Workbooks(i).Sheets(j).Name & "'!A1", _
                        TextToDisplay:=Workbooks(i).Sheets(j).Name
                End With
            Next j
        End If
    Next i
    
    logSh.sh.Cells.EntireColumn.AutoFit
    logBk.SaveAs FileName:=SaveAsFilePath & "\" _
        & Format(Now, "yyyymmdd") & "_" & Format(Time, "hhmmss") & "_" & "bookslog.xlsx"
    
End Sub

Sub SheetsCopy()
Attribute SheetsCopy.VB_ProcData.VB_Invoke_Func = "C\n14"
'---------------------------------------------
'   sheetslist�ɂĎ擾�����ꗗ��flag��Ɂu1�v�����͂���Ă���V�[�g�S�Ă�V�K�u�b�N�ɏW�񂷂�
'   �W�񂵂��u�b�N���\������邪�ۑ��͂���Ă��Ȃ��̂œK�X�ۑ�������K�v������
'   C
'---------------------------------------------
    Dim logBk As Workbook, logSh As New SheetInfo
    Dim rcvBk As Workbook, rcvSh As Worksheet
    Dim wbLoop As Workbook
    Dim wsName As String, wbPath As String
    Dim i As Long, flg As Boolean
    
    Application.ScreenUpdating = False
    For Each wbLoop In Workbooks
        If wbLoop.Name Like "*bookslog.xlsx" Then
            flg = True
            Set logBk = wbLoop
        End If
    Next wbLoop
    If flg = False Then
        MsgBox "�u�b�N���O���J����ĂȂ����ߏ������o���܂���B", vbExclamation, "SheetsCopy"
        Exit Sub
    End If
    Set logSh.sh = logBk.Sheets(1)
    Set rcvBk = Workbooks.Add
    Set rcvSh = rcvBk.ActiveSheet
    
    For i = 2 To logSh.EndRow()
        If logSh.sh.Cells(i, D__).Value = 1 Then
            wsName = logSh.sh.Cells(i, C__).Value
            Workbooks.Open (logSh.sh.Cells(i, A__).Value & "\" & logSh.sh.Cells(i, B__).Value)
            ActiveWorkbook.Sheets(wsName).Copy Before:=rcvSh
        End If
    Next i
    
    If rcvBk.Worksheets.Count > 1 Then
        Application.DisplayAlerts = False
        rcvBk.Sheets(Worksheets.Count).Delete
        Application.DisplayAlerts = True
    End If
    Application.ScreenUpdating = True
    
End Sub

Sub shProtect()
Attribute shProtect.VB_ProcData.VB_Invoke_Func = "P\n14"
'---------------------------------------------
'   �N������ƃu�b�N�̃A�N�e�B�u�V�[�g��ی삷��
'   ������ɂ����̂Ŋm�F�pMsgBox��ݒu
'   P
'---------------------------------------------
    If ActiveSheet.ProtectContents = True Then
        MsgBox "�V�[�g�̓��b�N����Ă��܂��B", vbExclamation, "shProtect"
    Else
        ActiveSheet.Protect
        MsgBox "�V�[�g�����b�N���܂����B", vbInformation, "shProtect"
    End If
    
End Sub

Sub shUnprotect()
Attribute shUnprotect.VB_ProcData.VB_Invoke_Func = "U\n14"
'---------------------------------------------
'   �N������ƃu�b�N�̃A�N�e�B�u�V�[�g�̕ی����������
'   �O�����ׂ̈�YesNo��ݒu
'   U
'---------------------------------------------
    Dim userResponse As Integer
    userResponse = MsgBox("���b�N���������܂����H", vbYesNo + vbQuestion, "shUnprotect")
    If userResponse = vbYes Then
        If ActiveSheet.ProtectContents = False Then
            MsgBox "�V�[�g�̓��b�N����Ă��܂���B", vbExclamation, "shUnprotect"
        Else
            ActiveSheet.Unprotect
            MsgBox "���b�N���������܂����B" & vbCrLf & "��ƌ�͕K���V�[�g��ی삵�Ă��������B", vbInformation, "shUnprotect"
        End If
    End If

End Sub

Sub ActiveSheetInfo()
Attribute ActiveSheetInfo.VB_ProcData.VB_Invoke_Func = "I\n14"
'---------------------------------------------
'   �N���������_�ł̃A�N�e�B�u�ȃu�b�N�̏ڍׂ̈ꗗ��\��
'   ��{�I�Ƀ}�N���̍쐬���̏��𓾂邽�߂ɗ��p
'   I
'---------------------------------------------
    Dim wb As Workbook
    Dim ws As New SheetInfo
    Dim colNum As Long
    
    Set wb = ActiveWorkbook
    Set ws.sh = wb.ActiveSheet
    colNum = ws.EndColumn(ActiveCell.Row)
    
    MsgBox "ActiveWorkbook : " & wb.Name & vbCrLf & _
        "ActiveSheet :  " & ws.sh.Name & vbCrLf & _
        "ActiveCellAddress : " & ActiveCell.Address & vbCrLf & _
        "value : " & ActiveCell.Value & vbCrLf & _
        "SelectionAddress : " & Selection.Address & vbCrLf & _
        "EndRow : " & ws.EndRow(ActiveCell.Column) & vbCrLf & _
        "EndColumn : " & colNum & "  (" & ColumnNumberToLetter(colNum) & ")" & vbCrLf & _
         "Color : " & ActiveCell.Interior.Color & vbCrLf & _
        "ThemeColor : " & ActiveCell.Interior.ThemeColor & vbCrLf & _
        "TintAndShade : " & ActiveCell.Interior.TintAndShade & vbCrLf & _
        "Font Color : " & ActiveCell.Font.Color, vbInformation, "ActiveSheetInfo"
    
End Sub


