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
            ProcessAsWorksheet wsTemplate, tbl, i, outputName
        End If

    NextRow:
    Next i
    
    ' マスタシートをアクティブにする
    wsMaster.Activate

    ' スクリーン更新をオンに戻す
    Application.ScreenUpdating = True
End Sub

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

Sub ProcessAsWorksheet(wsTemplate As Worksheet, tbl As ListObject, i As Long, outputName As String)
    Dim wsNew As Worksheet
    Dim rng As Range
    Dim j As Long
    Dim replaceFrom As String
    Dim replaceTo As String
    
    ' シートをコピーして新しいシートを作成
    wsTemplate.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
    
    ' 新しいシートを選択して名前を変更
    ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count).Name = outputName
    
    ' 新しいシートを変数にセット
    Set wsNew = ThisWorkbook.Sheets(outputName)
    
    ' 2列目以降の列をループ
    For j = 2 To tbl.ListColumns.Count
        ' 変換元と変換先の文字列を取得
        replaceFrom = tbl.DataBodyRange.Cells(1, j).Value
        replaceTo = tbl.DataBodyRange.Cells(i, j).Value
        
        ' 変換元と変換先が空でない場合のみ置換を行う
        If Not IsEmpty(replaceFrom) And Not IsEmpty(replaceTo) Then
            ' シート内の全てのセルを検索し、replaceFromをreplaceToに置換
            For Each rng In wsNew.UsedRange
                If rng.Value <> "" Then
                    rng.Value = Replace(rng.Value, replaceFrom, replaceTo)
                End If
            Next rng
        End If
    Next j
End Sub

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
    With fd.Title = "Select Path".AllowMultiSelect = False
        If .Show = True Then
            selectedPath = .SelectedItems(1)
        End If
    End With

    ' 選択したパスを "path" という名前のセルに書き込む
    Dim rng As Range
    Set rng = ThisWorkbook.Names("path").RefersToRange
    rng.Value = selectedPath
End Sub