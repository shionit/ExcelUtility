Attribute VB_Name = "MyCommonModule"
Option Explicit
'************************************************
'MyCommonModule
' PERSONAL.XLS������VBA���C�u����
'************************************************

' ������̐擪�Ɩ����ɕ������t������
Sub AppendHeadTail()
    Dim head As String
    Dim tail As String
    Dim target As Range

    head = InputBox("�擪�ɕt�����镶������w��")
    tail = InputBox("�����ɕt�����镶������w��")

    For Each target In Selection
        target.Value = head & target.Value & tail
    Next

End Sub

' �Z���̌���/�������s��
Sub ToggleMergeCells()
    On Error Resume Next
    With Selection
        .MergeCells = Not .MergeCells
    End With
End Sub

' �t�@�C���̍ŏI�X�V�������擾����
Public Function GetLastSaveTime()
    Application.Volatile
    GetLastSaveTime = ActiveWorkbook.BuiltinDocumentProperties("Last save time").Value
End Function

' �V�[�g�̕��בւ����s��
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

' �V�[�g���̈ꗗ���쐬����
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

' �I��͈͓��̏d���l���n�C���C�g����
Sub HighlightDuplication()
    Const DUPLICATE_COLOR_INDEX As Integer = 46
    Const DUPLICATE_PATTERN = xlSolid
    Dim cellA As Range
    Dim cellB As Range
    
    For Each cellA In Selection
        ' �d������ς݃Z���Ƌ�Z���̓X�L�b�v
        If cellA.Interior.ColorIndex <> 46 And cellA.Value <> "" Then
            For Each cellB In Selection
                If cellA.Row <> cellB.Row Or cellA.Column <> cellB.Column Then
                    '���g����Ȃ��Z���̒l��r
                    If cellA.Value = cellB.Value Then
                        ' �d�����Ă�����ǂ������F�t��
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
