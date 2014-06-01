Attribute VB_Name = "JiraModule"
Option Explicit
'************************************************
'JiraModule
' Issue管理ツールJIRAから出力したExcelシート用マクロ
'************************************************

Sub SetJIRAListStyle()
'
' SetJIRAListStyle Macro
' JIRAから出力した一覧にフィルタや条件付き書式を設定します
'
    Dim header As Range
    Dim dueDate As String
'
    Set header = Rows("4:4")
    
    header.Select
    Selection.AutoFilter
    
    ' format Priority
    Call SelectAllRow(header, "優先度")
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, _
        Formula1:="=""緊急"""
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Color = -16383844
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, _
        Formula1:="=""高"""
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Color = -16751204
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 10284031
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
    ' format Status
    Call SelectAllRow(header, "状況")
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, _
        Formula1:="=""解決済"""
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .ThemeColor = xlThemeColorLight2
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.249946592608417
    End With
    Selection.FormatConditions(1).StopIfTrue = False

    ' format Due Date
    Call SelectAllRow(header, "期限日")
    Call HighlightDueDate("緑", -16752384, 13561798)
    Call HighlightDueDate("黄", -16751204, 10284031)
    Call HighlightDueDate("赤", -16383844, 13551615)
    
    ' Print Option
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .FitToPagesWide = 1
        .FitToPagesTall = 1
    End With
    Application.PrintCommunication = True

End Sub

' 指定したヘッダー文字列の列を全選択する
Private Sub SelectAllRow(header As Range, headerText As String)
    Dim target As Range
    
    Set target = header.Find(headerText)
    target.Offset(1, 0).Select
    Range(Selection, Selection.End(xlDown)).Select
End Sub

' 指定日以前の期日をハイライトする
Private Sub HighlightDueDate(colorJa As String, fontColor As Long, interiorColor As Long)
    Dim dueDate As String
    
    dueDate = Date
    dueDate = InputBox(colorJa & "ハイライトする期限日を指定", "期限日", dueDate)
    If Trim$(dueDate) <> "" Then
        Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
            Formula1:="=" & CDbl(DateValue(dueDate))
        Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
        With Selection.FormatConditions(1).Font
            .Color = fontColor
            .TintAndShade = 0
        End With
        With Selection.FormatConditions(1).Interior
            .PatternColorIndex = xlAutomatic
            .Color = interiorColor
            .TintAndShade = 0
        End With
        Selection.FormatConditions(1).StopIfTrue = False
    End If
End Sub
