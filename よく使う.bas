Attribute VB_Name = "よく使う"
Option Explicit
Sub 最下行最右列取得()
    Dim 最下行 As Long, 最右列 As Long
    With ActiveSheet
        最下行 = .Cells(Rows.Count, 1).End(xlUp).Row
        最右列 = .Cells(1, Columns.Count).End(xlToLeft).Column
    End With
End Sub
Sub 配列貼付()
    Dim 配列(1 To 3, 1 To 3)
    With ActiveSheet
        Range(.Cells(1, 1), .Cells(3, 3)) = 配列
    End With
End Sub
Sub 保護切替()
    With ActiveSheet
        Select Case .ProtectContents
            Case True: .Unprotect: MsgBox "シート保護を解除しました"
            Case False: .Protect: MsgBox "シートを保護しました"
        End Select
    End With
End Sub
Sub 全シート展開(シート名 As String)
    Dim シート As Worksheet
    Application.ScreenUpdating = False
    For Each シート In Sheets
        シート.Visible = True
    Next
    Sheets(シート名).Activate
    Application.ScreenUpdating = True
End Sub
Sub 全シート非表示(シート名 As String)
    Dim シート As Worksheet
    Application.ScreenUpdating = False
    Sheets(シート名).Visible = True
    For Each シート In Sheets
        If シート.Name <> シート名 Then シート.Visible = False
    Next
    Application.ScreenUpdating = True
End Sub
Sub 処理時間計測()
    Dim 始時 As Date, 終時 As Date
    始時 = Timer
    実行中.Show vbModeless
    実行中.Repaint
    
    終時 = Timer
    MsgBox "処理が完了しました" & vbCrLf & vbCrLf & "処理時間：" & 終時 - 始時
    Unload 実行中
End Sub
Sub 並替基本形とカナ削除()
    Dim 最下行, 最右列 As Long
    With ActiveSheet
        最下行 = .Cells(Rows.Count, 1).End(xlUp).Row
        最右列 = .Cells(1, Columns.Count).End(xlToLeft).Column
        .Range(Cells(1, 1), Cells(最下行, 最右列)).Characters.PhoneticCharacters = ""
        With .Sort
            With .SortFields
                .Clear
                .Add Key:=Range("A1"), Order:=xlAscending
                .Add Key:=Range("B1"), Order:=xlDescending
            End With
            .SetRange Range(Cells(1, 1), Cells(最下行, 最右列))
            .Header = xlYes
            .Apply
        End With
    End With
End Sub
Sub データと罫線クリア(シート名 As String)
    With Sheets(シート名)
        Range(.Cells(1, 1), .Cells(Rows.Count, Columns.Count)).ClearContents
        Range(.Cells(1, 1), .Cells(Rows.Count, Columns.Count)).Borders.LineStyle = False
    End With
    MsgBox "「" & シート名 & "」シートの内容をクリアしました"
End Sub
Sub PDF出力(フォルダ名 As String, ファイル名 As String, シート名 As String)
    フォルダ名 = ThisWorkbook.Path & "\" & フォルダ名
    If Dir(フォルダ名, vbDirectory) = "" Then MkDir フォルダ名
    With Sheets(シート名)
        .ExportAsFixedFormat Type:=xlTypePDF, Filename:=フォルダ名 & "\" & ファイル名 & ".pdf"
        MsgBox "ファイル名：" & ファイル名 & ".pdf" & vbCrLf & vbCrLf & "PDF出力が完了しました（本ツール同階層・「" & フォルダ名 & "」フォルダ内）"
    End With
End Sub
Sub 条件付書式等設定例()
    Dim 条件 As FormatCondition
    With ActiveSheet
        .Range("A1:D4").Borders.LineStyle = True
        .Cells.FormatConditions.Delete
        Set 条件 = .Range("A1:D4").FormatConditions.Add(Type:=xlExpression, Formula1:="=A1=0")
        条件.Font.Color = RGB(255, 0, 0)
        Set 条件 = .Range("A1:D4").FormatConditions.Add(Type:=xlExpression, Formula1:="=A1=1")
        条件.Interior.Color = RGB(252, 228, 214)
    End With
End Sub
Sub バックアップテキスト出力(データ)
    Dim ファイル名
    ファイル名 = ThisWorkbook.Path & "\BU.txt"
    Open ファイル名 For Append As #1
    Print #1, データ
    Close #1
End Sub
Sub 検索ダイアログ展開()
    Application.CommandBars.FindControl(ID:=1849).Execute
End Sub
Sub Enter方向切替()
    Application.MoveAfterReturn = True
    Select Case Application.MoveAfterReturnDirection
        Case xlToRight: Application.MoveAfterReturnDirection = xlDown
        Case xlDown: Application.MoveAfterReturnDirection = xlToRight
    End Select
End Sub
Sub イベント制御()
    Select Case Application.EnableEvents
        Case False: Application.EnableEvents = True: MsgBox "自動処理機能をONに切り替えました"
        Case True: Application.EnableEvents = False: MsgBox "自動処理機能をOFFに切り替えました"
    End Select
End Sub
