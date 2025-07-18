Attribute VB_Name = "myModule"
Option Explicit

'---------------------------------------------
'   ショートカットキーを設定しているのでお好みで任意のキーを設定してください
'   ショートカットキーはCtrl + Shift + 任意のキー
'----------------------------------------------

Sub home()
Attribute home.VB_ProcData.VB_Invoke_Func = "E\n14"
'---------------------------------------------
'   マクロ起動時にアクティブなブックの全シートの選択セルをA1に戻す
'   アクティブシートは最終シートになる
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
'   起動時に開いている全てのブックのパス、ブック名、シート名を取得して一覧にする
'   リンクも一緒に作成され、アクセス可能
'   L
'---------------------------------------------
    Dim logBk As Workbook, logSh As New SheetInfo
    Dim wbCnt As Variant, wsCnt As Variant
    Dim i As Long, j As Long, SaveAsFilePath As String
    
    SaveAsFilePath = "C:\Path\BooksLog" ' 保存先のパスを指定してください
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
'   sheetslistにて取得した一覧のflag列に「1」が入力されているシート全てを新規ブックに集約する
'   集約したブックが表示されるが保存はされていないので適宜保存をする必要がある
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
        MsgBox "ブックログが開かれてないため処理が出来ません。", vbExclamation, "SheetsCopy"
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
'   起動するとブックのアクティブシートを保護する
'   分かりにくいので確認用MsgBoxを設置
'   P
'---------------------------------------------
    If ActiveSheet.ProtectContents = True Then
        MsgBox "シートはロックされています。", vbExclamation, "shProtect"
    Else
        ActiveSheet.Protect
        MsgBox "シートをロックしました。", vbInformation, "shProtect"
    End If
    
End Sub

Sub shUnprotect()
Attribute shUnprotect.VB_ProcData.VB_Invoke_Func = "U\n14"
'---------------------------------------------
'   起動するとブックのアクティブシートの保護を解除する
'   念押しの為にYesNoを設置
'   U
'---------------------------------------------
    Dim userResponse As Integer
    userResponse = MsgBox("ロックを解除しますか？", vbYesNo + vbQuestion, "shUnprotect")
    If userResponse = vbYes Then
        If ActiveSheet.ProtectContents = False Then
            MsgBox "シートはロックされていません。", vbExclamation, "shUnprotect"
        Else
            ActiveSheet.Unprotect
            MsgBox "ロックを解除しました。" & vbCrLf & "作業後は必ずシートを保護してください。", vbInformation, "shUnprotect"
        End If
    End If

End Sub

Sub ActiveSheetInfo()
Attribute ActiveSheetInfo.VB_ProcData.VB_Invoke_Func = "I\n14"
'---------------------------------------------
'   起動した時点でのアクティブなブックの詳細の一覧を表示
'   基本的にマクロの作成時の情報を得るために利用
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


