Attribute VB_Name = "�悭�g���R�[�h"
Option Explicit
Sub �ŉ��s�ŉE��擾()
    Dim �ŉ��s As Long, �ŉE�� As Long
    With ActiveSheet
        �ŉ��s = .Cells(Rows.Count, 1).End(xlUp).Row
        �ŉE�� = .Cells(1, Columns.Count).End(xlToLeft).Column
    End With
End Sub
Sub �z��\�t()
    Dim �z��(1 To 3, 1 To 3)
    With ActiveSheet
        Range(.Cells(1, 1), .Cells(3, 3)) = �z��
    End With
End Sub
Sub �ی�ؑ�()
    With ActiveSheet
        Select Case .ProtectContents
            Case True: .Unprotect: MsgBox "�V�[�g�ی���������܂���"
            Case False: .Protect: MsgBox "�V�[�g��ی삵�܂���"
        End Select
    End With
End Sub
Sub �S�V�[�g�W�J(�V�[�g�� As String)
    Dim �V�[�g As Worksheet
    Application.ScreenUpdating = False
    For Each �V�[�g In Sheets
        �V�[�g.Visible = True
    Next
    Sheets(�V�[�g��).Activate
    Application.ScreenUpdating = True
End Sub
Sub �S�V�[�g��\��(�V�[�g�� As String)
    Dim �V�[�g As Worksheet
    Application.ScreenUpdating = False
    Sheets(�V�[�g��).Visible = True
    For Each �V�[�g In Sheets
        If �V�[�g.Name <> �V�[�g�� Then �V�[�g.Visible = False
    Next
    Application.ScreenUpdating = True
End Sub
Sub �������Ԍv��()
    Dim �n�� As Date
    �n�� = Timer
    ���s��.Show vbModeless
    ���s��.Repaint
    
    MsgBox "�������������܂���" & vbCrLf & vbCrLf & "�������ԁF" & Now - �n��
    Unload ���s��
End Sub
Sub ���֊�{�`�ƃJ�i�폜()
    Dim �ŉ��s, �ŉE�� As Long
    With ActiveSheet
        �ŉ��s = .Cells(Rows.Count, 1).End(xlUp).Row
        �ŉE�� = .Cells(1, Columns.Count).End(xlToLeft).Column
        .Range(Cells(1, 1), Cells(�ŉ��s, �ŉE��)).Characters.PhoneticCharacters = ""
        If �ŉ��s < 2 Then Exit Sub
        With .Sort
            With .SortFields
                .Clear
                .Add Key:=Range("A1"), Order:=xlAscending
                .Add Key:=Range("B1"), Order:=xlDescending
            End With
            .SetRange Range(Cells(1, 1), Cells(�ŉ��s, �ŉE��))
            .Header = xlYes
            .Apply
        End With
    End With
End Sub
Sub �f�[�^�ƌr���N���A(�V�[�g�� As String)
    With Sheets(�V�[�g��)
        Range(.Cells(1, 1), .Cells(Rows.Count, Columns.Count)).ClearContents
        Range(.Cells(1, 1), .Cells(Rows.Count, Columns.Count)).Borders.LineStyle = False
    End With
    MsgBox "�u" & �V�[�g�� & "�v�V�[�g�̓��e���N���A���܂���"
End Sub
Sub PDF�o��(�t�H���_�� As String, �t�@�C���� As String, �V�[�g�� As String)
    �t�H���_�� = ThisWorkbook.Path & "\" & �t�H���_��
    If Dir(�t�H���_��, vbDirectory) = "" Then MkDir �t�H���_��
    With Sheets(�V�[�g��)
        .ExportAsFixedFormat Type:=xlTypePDF, Filename:=�t�H���_�� & "\" & �t�@�C���� & ".pdf"
        MsgBox "�t�@�C�����F" & �t�@�C���� & ".pdf" & vbCrLf & vbCrLf & "PDF�o�͂��������܂����i�{�c�[�����K�w�E�u" & �t�H���_�� & "�v�t�H���_���j"
    End With
End Sub
Sub �����t�������ݒ��()
    Dim ���� As FormatCondition
    With ActiveSheet
        .Range("A1:D4").Borders.LineStyle = True
        .Cells.FormatConditions.Delete
        Set ���� = .Range("A1:D4").FormatConditions.Add(Type:=xlExpression, Formula1:="=A1=0")
        ����.Font.Color = RGB(255, 0, 0)
        Set ���� = .Range("A1:D4").FormatConditions.Add(Type:=xlExpression, Formula1:="=A1=1")
        ����.Interior.Color = RGB(252, 228, 214)
    End With
End Sub
Sub ���X�g�ɂ����͋K���ݒ�(���X�g)
'256���ȏ�̕�����`���Ń��X�g�`���̓��͋K���ݒ肷��ƃt�@�C�����j�����邱�Ƃɒ��ӁI
'���@�{��ł̓��X�g�ݒ�Z���Q�ƌ`���ɂ��邱�ƂŃt�@�C���j����������Ă��܂�
    Dim �͈͎� As String
    With Sheets("�v���_�E���ݒ�")
        .Cells(1, 1).Resize(Rows.Count, 1).ClearContents
        .Cells(1, 1).Resize(UBound(���X�g), 1) = ���X�g
        �͈͎� = "=�v���_�E���ݒ�!$A$1:$A$" & UBound(���X�g)
    End With
    With Sheets("���̓t�H�[��")
        With .Cells(1, 1).Validation
            .Delete
            .Add Type:=xlValidateList, Formula1:=�͈͎�
        End With
    End With
End Sub
Sub �o�b�N�A�b�v�e�L�X�g�o��(�f�[�^)
    Dim �t�@�C����
    �t�@�C���� = ThisWorkbook.Path & "\BU.txt"
    Open �t�@�C���� For Append As #1
    Print #1, �f�[�^
    Close #1
End Sub
Sub �����_�C�A���O�W�J()
    Application.CommandBars.FindControl(ID:=1849).Execute
End Sub
Sub Enter�����ؑ�()
    Application.MoveAfterReturn = True
    Select Case Application.MoveAfterReturnDirection
        Case xlToRight: Application.MoveAfterReturnDirection = xlDown
        Case xlDown: Application.MoveAfterReturnDirection = xlToRight
    End Select
End Sub
Sub �C�x���g����()
    Select Case Application.EnableEvents
        Case False: Application.EnableEvents = True: MsgBox "���������@�\��ON�ɐ؂�ւ��܂���"
        Case True: Application.EnableEvents = False: MsgBox "���������@�\��OFF�ɐ؂�ւ��܂���"
    End Select
End Sub
Sub �}�`�\���ؑ�(�V�[�g�� As String, �}�`�I�u�W�F�N�g�� As String)
    With Sheets(�V�[�g��)
        With .Shapes(�}�`�I�u�W�F�N�g��)
            Select Case .Visible
                Case True: .Visible = False
                Case False: .Visible = True
            End Select
        End With
    End With
End Sub

