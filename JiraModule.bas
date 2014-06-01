Attribute VB_Name = "JiraModule"
Option Explicit
'************************************************
'JiraModule
' Issue�Ǘ��c�[��JIRA����o�͂���Excel�V�[�g�p�}�N��
'************************************************

Sub SetJIRAListStyle()
'
' SetJIRAListStyle Macro
' JIRA����o�͂����ꗗ�Ƀt�B���^������t��������ݒ肵�܂�
'
    Dim header As Range
    Dim dueDate As String
'
    Set header = Rows("4:4")
    
    header.Select
    Selection.AutoFilter
    
    ' format Priority
    Call SelectAllRow(header, "�D��x")
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, _
        Formula1:="=""�ً}"""
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
        Formula1:="=""��"""
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
    Call SelectAllRow(header, "��")
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, _
        Formula1:="=""������"""
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
    Call SelectAllRow(header, "������")
    Call HighlightDueDate("��", -16752384, 13561798)
    Call HighlightDueDate("��", -16751204, 10284031)
    Call HighlightDueDate("��", -16383844, 13551615)
    
    ' Print Option
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .FitToPagesWide = 1
        .FitToPagesTall = 1
    End With
    Application.PrintCommunication = True

End Sub

' �w�肵���w�b�_�[������̗��S�I������
Private Sub SelectAllRow(header As Range, headerText As String)
    Dim target As Range
    
    Set target = header.Find(headerText)
    target.Offset(1, 0).Select
    Range(Selection, Selection.End(xlDown)).Select
End Sub

' �w����ȑO�̊������n�C���C�g����
Private Sub HighlightDueDate(colorJa As String, fontColor As Long, interiorColor As Long)
    Dim dueDate As String
    
    dueDate = Date
    dueDate = InputBox(colorJa & "�n�C���C�g������������w��", "������", dueDate)
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
