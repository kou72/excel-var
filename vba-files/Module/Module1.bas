Attribute VB_Name = "Module1"

Sub CreateAndModifySheetsFromVarList()
    ' スクリーン更新をオフにする
    Application.ScreenUpdating = False

    ' マスタ情報をセット
    Dim wsMaster As Worksheet
    Set wsMaster = ActiveSheet

    Dim templateName As String
    templateName = wsMaster.Range("template").Value

    Dim outputType As String
    outputType = wsMaster.Range("type").Value

    Dim filePath As String
    filePath = ThisWorkbook.Names("path").RefersToRange.Value
    If Not CheckFilePath(filePath, outputType) Then Exit Sub

    Dim tbl As ListObject
    Set tbl = wsMaster.ListObjects("varlist")
    
    ' varlistの各行をループ
    Dim i As Long
    For i = 2 To tbl.ListRows.Count
        ' テンプレートシート名と出力名を取得
        Dim outputName As String
        outputName = tbl.ListColumns(1).DataBodyRange.Cells(i).Value
        
        ' outputNameが空白またはNothingなら次の行へ
        If IsEmpty(outputName) Or outputName = "" Then
            GoTo NextRow
        End If
        
        ' テンプレートシートを指定
        Dim wsTemplate As Worksheet
        Set wsTemplate = ThisWorkbook.Sheets(templateName)
        
        If outputType = "textFile" Then
            ProcessAsTextFile wsTemplate, tbl, i, outputName
        Else
            ProcessAsWorksheet wsTemplate, tbl, i, outputName, filePath
        End If

    NextRow:
    Next i
    
    ' マスタシートをアクティブにする
    ' wsMaster.Activate

    ' スクリーン更新をオンに戻す
    Application.ScreenUpdating = True
End Sub

' filePathが有効なフォルダまたはExcelファイルを指しているか確認する関数
Function CheckFilePath(filePath As String, outputType As String) As Boolean
    If outputType = "sheet" Then ' Excelファイルのチェック
        On Error Resume Next ' エラーハンドラを有効にする
        Dim wb As Workbook
        Set wb = Workbooks.Open(filePath) ' filePathを開く
        On Error GoTo 0 ' エラーハンドラを無効にする
        If wb Is Nothing Then ' filePathが無効ならエラーメッセージを表示して終了
            MsgBox "選択されたファイルが無効です。有効なExcelファイルを選択してください。", vbCritical, "エラー"
            CheckFilePath = False
        Else ' filePathが有効ならTrueを返す
            wb.Close False
            CheckFilePath = True
        End If
    ElseIf outputType = "textFile" Then ' フォルダのチェック
        If Dir(filePath, vbDirectory) = "" Then
            MsgBox "選択されたパスが無効です。有効なフォルダを選択してください。", vbCritical, "エラー"
            CheckFilePath = False
        Else
            CheckFilePath = True
        End If
    End If
End Function

Sub ProcessAsTextFile(wsTemplate As Worksheet, tbl As ListObject, i As Long, outputName As String)
    Dim rng As Range
    Dim j As Long
    Dim replaceFrom As String
    Dim replaceTo As String
    Dim textOutput As String
    Dim fileName As String
    
    ' テンプレートシートの内容をテキストに変換
    For Each rng In wsTemplate.UsedRange
        textOutput = textOutput & rng.Value & vbTab
        If rng.Column = wsTemplate.UsedRange.Columns.Count Then
            textOutput = textOutput & vbCrLf
        End If
    Next rng
    
    ' 2列目以降の列をループ
    For j = 2 To tbl.ListColumns.Count
        ' 変換元と変換先の文字列を取得
        replaceFrom = tbl.DataBodyRange.Cells(1, j).Value
        replaceTo = tbl.DataBodyRange.Cells(i, j).Value
        
        ' 変換元と変換先が空でない場合のみ置換を行う
        If Not IsEmpty(replaceFrom) And Not IsEmpty(replaceTo) Then
            ' テキスト内のreplaceFromをreplaceToに置換
            textOutput = Replace(textOutput, replaceFrom, replaceTo)
        End If
    Next j
    
    ' テキストファイル名を設定
    fileName = ThisWorkbook.Path & "\" & outputName
    
    ' テキストファイルに出力
    Open fileName For Output As #1
    Print #1, textOutput
    Close #1
    
    ' テキスト出力変数をリセット
    textOutput = ""
End Sub

Sub ProcessAsWorksheet(wsTemplate As Worksheet, tbl As ListObject, i As Long, outputName As String, filePath As String)
    ' ' 指定されたパスのWorkbookを開く
    Dim wbTarget As Workbook
    Set wbTarget = Workbooks.Open(filePath)
    
    ' 重複しないシート名を見つけるためのループ
    Dim suffix As Integer
    suffix = 0
    Dim newSheetName As String
    newSheetName = outputName
    Do While WorksheetExists(wbTarget, newSheetName)
        suffix = suffix + 1
        newSheetName = outputName & " (" & suffix & ")"
    Loop
    
    ' シートをコピーして新しいシートを作成し、名前を変更
    wsTemplate.Copy After:=wbTarget.Sheets(wbTarget.Sheets.Count)
    wbTarget.Sheets(wbTarget.Sheets.Count).Name = newSheetName
    Dim wsNew As Worksheet
    Set wsNew = wbTarget.Sheets(newSheetName)
    
    ' 2列目以降の列をループ
    Dim j As Long
    For j = 2 To tbl.ListColumns.Count
        ' 変換元と変換先の文字列を取得
        Dim replaceFrom As String
        Dim replaceTo As String
        replaceFrom = tbl.DataBodyRange.Cells(1, j).Value
        replaceTo = tbl.DataBodyRange.Cells(i, j).Value
        
        ' 変換元と変換先が空でない場合のみ置換を行う
        If Not IsEmpty(replaceFrom) And Not IsEmpty(replaceTo) Then
            ' シート内の全てのセルを検索し、replaceFromをreplaceToに置換
            Dim rng As Range
            For Each rng In wsNew.UsedRange
                If rng.Value <> "" Then
                    rng.Value = Replace(rng.Value, replaceFrom, replaceTo)
                End If
            Next rng
        End If
    Next j
End Sub

' シートが存在するかどうかを確認するための関数
Function WorksheetExists(wb As Workbook, wsName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Sheets(wsName)
    On Error GoTo 0
    WorksheetExists = Not ws Is Nothing
End Function


Sub SelectFileOrFolderAndWritePath()
    ' "type"という名前のセルの内容を取得
    Dim selectType As Range
    Set selectType = ThisWorkbook.Names("type").RefersToRange

    ' "type"の内容に基づいてFileDialogのタイプを設定
    Dim fd As FileDialog
    If selectType.Value = "sheet" Then
        Set fd = Application.FileDialog(msoFileDialogFilePicker)
    ElseIf selectType.Value = "textFile" Then
        Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    Else
        MsgBox "名前付きセル 'type' の値が無効です"
        Exit Sub
    End If

    ' ダイアログを表示し、選択したパスを取得
    Dim selectedPath As String
    With fd
        .Title = "Select Path"
        .AllowMultiSelect = False
        
        If .Show = True Then
            selectedPath = .SelectedItems(1)
        End If
    End With

    ' 選択したパスを "path" という名前のセルに書き込む
    Dim rng As Range
    Set rng = ThisWorkbook.Names("path").RefersToRange
    rng.Value = selectedPath
End Sub