Attribute VB_Name = "MyCommonModule"
Option Explicit
'************************************************
'MyCommonModule
' PERSONAL.XLS向きのVBAライブラリ
'************************************************

' 文字列の先頭と末尾に文字列を付加する
Sub AppendHeadTail()
    Dim head As String
    Dim tail As String
    Dim target As Range

    head = InputBox("先頭に付加する文字列を指定")
    tail = InputBox("末尾に付加する文字列を指定")

    For Each target In Selection
        target.Value = head & target.Value & tail
    Next

End Sub

' セルの結合/解除を行う
Sub ToggleMergeCells()
    On Error Resume Next
    With Selection
        .MergeCells = Not .MergeCells
    End With
End Sub

' ファイルの最終更新日時を取得する
Public Function GetLastSaveTime()
    Application.Volatile
    GetLastSaveTime = ActiveWorkbook.BuiltinDocumentProperties("Last save time").Value
End Function

' シートの並べ替えを行う
Sub SortSheets()
    Dim i As Integer
    Dim j As Integer
    
    Application.ScreenUpdating = False
    
    For i = 1 To Sheets.Count
        For j = 1 To Sheets.Count - 1
            If Sheets(j).Name > Sheets(j + 1).Name Then
                Sheets(j).Move after:=Sheets(j + 1)
            End If
        Next j
    Next i
    Application.ScreenUpdating = True
End Sub

' シート名の一覧を作成する
Sub CreateSheetNameList()
    Dim shtSheet As Worksheet
    Dim shtContentsList As Worksheet
    
    Set shtContentsList = ActiveWorkbook.Worksheets.Add(ActiveWorkbook.Sheets(1))
    shtContentsList.Name = "ContentsList"
    For Each shtSheet In ActiveWorkbook.Sheets
        With shtContentsList.Cells(shtSheet.Index, 1)
            .Value = shtSheet.Name
            Call .Hyperlinks.Add(shtContentsList.Cells(shtSheet.Index, 1), "", _
                "'" & shtSheet.Name & "'!A1")
        End With
    Next
    
End Sub

' 選択範囲内の重複値をハイライトする
Sub HighlightDuplication()
    Const DUPLICATE_COLOR_INDEX As Integer = 46
    Const DUPLICATE_PATTERN = xlSolid
    Dim cellA As Range
    Dim cellB As Range
    
    For Each cellA In Selection
        ' 重複判定済みセルと空セルはスキップ
        If cellA.Interior.ColorIndex <> 46 And cellA.Value <> "" Then
            For Each cellB In Selection
                If cellA.Row <> cellB.Row Or cellA.Column <> cellB.Column Then
                    '自身じゃないセルの値比較
                    If cellA.Value = cellB.Value Then
                        ' 重複していたらどっちも色付け
                        With cellA.Interior
                            .ColorIndex = DUPLICATE_COLOR_INDEX
                            .Pattern = xlSolid
                        End With
                        With cellB.Interior
                            .ColorIndex = DUPLICATE_COLOR_INDEX
                            .Pattern = xlSolid
                        End With
                    End If
                End If
            Next
        End If
    Next
End Sub
