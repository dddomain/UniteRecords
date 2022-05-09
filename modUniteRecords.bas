
Option Explicit

Sub UniteRecords()

    Dim folder As String
    Dim file As String
    Dim book As Workbook
    Dim i As Integer
    
    With Application.FileDialog(msoFileDialogFolderPicker)
        .InitialFileName = "D:/user/"
        .AllowMultiSelect = False
    
        .Title = "フォルダの選択"
        If .Show = True Then
            folder = .SelectedItems(1)
        End If
    End With
    
    file = Dir(folder & "\*.xlsx")
    
    '実行を開始する行を選択する
    i = Application.InputBox _
        ( _
        prompt:="集計を開始する入力していください", _
        Title:="集計開始行選択", _
        Type:=2 _
        )
    
    Do While file <> ""
        Set book = Workbooks.Open(folder & "\" & file)
        ThisWorkbook.Worksheets(1).Range("A" & CStr(i) & ":V" & CStr(i)).Value = book.Worksheets(1).Range("A5:V5").Value
        file = Dir()
        i = i + 1
        book.Close
    Loop
    
End Sub

