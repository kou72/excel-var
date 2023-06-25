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
    Dim tableDirection As String
    tableDirection = ThisWorkbook.Names("direction").RefersToRange.Value
    Dim tbl As ListObject
    Set tbl = wsMaster.ListObjects("varlist")

    ' 実行前に確認メッセージを表示
    Dim msg As String
    msg = "以下の内容で処理を実行します。" & vbCrLf & vbCrLf & _
        "テンプレートシート：" & templateName & vbCrLf & _
        "出力形式：" & outputType & vbCrLf & _
        "出力先：" & filePath & vbCrLf & _
        "出力数：" & GetOutputCount(tbl, tableDirection) & vbCrLf & _
        "変数数：" & GetVarCount(tbl, tableDirection) & vbCrLf & _
        "よろしいですか？"
    If MsgBox(msg, vbYesNo + vbQuestion, "確認") = vbNo Then Exit Sub

    ' 出力先が有効なフォルダまたはExcelファイルを指しているか確認
    If Not CheckFilePath(filePath, outputType) Then Exit Sub

    ' テーブルの方向で処理を分岐
    Dim i As Long
    Dim disableFlag As String
    Dim outputName As String
    Dim wsTemplate As Worksheet
    If tableDirection = "縦方向" Then
        ' varlistの各列をループ
        For i = 2 To tbl.ListColumns.Count
            ' "無効flag"が空白でない場合、この列の処理をスキップ
            disableFlag = tbl.Range.Cells(1, i).Offset(-1, 0).Value
            If disableFlag <> "" Then
                GoTo NextColumn
            End If

            ' テンプレートシート名と出力名を取得
            outputName = tbl.ListColumns(i).DataBodyRange.Cells(1).Value

            ' outputNameが空白またはNothingなら次の列へ
            If IsEmpty(outputName) Or outputName = "" Then
                GoTo NextColumn
            End If

            ' テンプレートシートを指定
            Set wsTemplate = ThisWorkbook.Sheets(templateName)
            
            If outputType = "textFile" Then
                ProcessAsTextFile wsTemplate, tbl, i, outputName, filePath, tableDirection
            Else
                ProcessAsWorksheet wsTemplate, tbl, i, outputName, filePath, tableDirection
            End If
        NextColumn:
        Next i
    Else
        ' varlistの各行をループ
        For i = 2 To tbl.ListRows.Count
            ' "無効flag"が空白でない場合、この行の処理をスキップ
            disableFlag = tbl.Range.Cells(i, 1).Offset(0, -1).Value
            If disableFlag <> "" Then
                GoTo NextRow
            End If

            ' テンプレートシート名と出力名を取得
            outputName = tbl.ListColumns(1).DataBodyRange.Cells(i).Value
            
            ' outputNameが空白またはNothingなら次の行へ
            If IsEmpty(outputName) Or outputName = "" Then
                GoTo NextRow
            End If

            ' テンプレートシートを指定
            Set wsTemplate = ThisWorkbook.Sheets(templateName)
            
            If outputType = "textFile" Then
                ProcessAsTextFile wsTemplate, tbl, i, outputName, filePath, tableDirection
            Else
                ProcessAsWorksheet wsTemplate, tbl, i, outputName, filePath, tableDirection
            End If

        NextRow:
        Next i

    End If

    ' スクリーン更新をオンに戻す
    Application.ScreenUpdating = True
End Sub

' 出力数を取得する関数
Function GetOutputCount(tbl As ListObject, tableDirection As String) As Long
    ' 出力数を初期化
    GetOutputCount = 0

    ' テーブルの方向で処理を分岐
    Dim i As Long
    Dim disableFlag As String
    Dim outputName As String
    If tableDirection = "縦方向" Then
        ' varlistの各行をループ
        For i = 2 To tbl.ListColumns.Count
            ' "無効flag"が空白でない場合、この列の処理をスキップ
            disableFlag = tbl.Range.Cells(1, i).Offset(-1, 0).Value
            If disableFlag <> "" Then
                GoTo NextColumn
            End If

            ' outputNameが空白またはNothingなら次の列へ
            outputName = tbl.ListColumns(i).DataBodyRange.Cells(1).Value
            If IsEmpty(outputName) Or outputName = "" Then
                GoTo NextColumn
            End If
            
            ' 出力数をカウント
            GetOutputCount = GetOutputCount + 1

        NextColumn:
        Next i
    Else
        ' varlistの各行をループ
        For i = 2 To tbl.ListRows.Count
            ' "無効flag"が空白でない場合、この行の処理をスキップ
            disableFlag = tbl.Range.Cells(i, 1).Offset(0, -1).Value
            If disableFlag <> "" Then
                GoTo NextRow
            End If

            ' outputNameが空白またはNothingなら次の行へ
            outputName = tbl.ListColumns(1).DataBodyRange.Cells(i).Value
            If IsEmpty(outputName) Or outputName = "" Then
                GoTo NextRow
            End If
            
            ' 出力数をカウント
            GetOutputCount = GetOutputCount + 1

        NextRow:
        Next i
    End If
End Function

' 変数数を取得する関数
Function GetVarCount(tbl As ListObject, tableDirection As String) As Long
    ' 変数数を初期化
    GetVarCount = 0

    ' テーブルの方向で処理を分岐
    Dim j As Long
    Dim disableFlag As String
    Dim replaceFrom As String
    If tableDirection = "縦方向" Then
        ' varlistの各行をループ
        For j = 2 To tbl.ListRows.Count
            ' "無効flag"が空白でない場合、この行の処理をスキップ
            disableFlag = tbl.Range.Cells(j, 1).Offset(0, -1).Value
            If disableFlag <> "" Then
                GoTo NextRow
            End If

            ' 変換元が空白またはNothingなら次の行へ
            replaceFrom = tbl.DataBodyRange.Cells(j, 1).Value
            If IsEmpty(replaceFrom) Or replaceFrom = "" Then
                GoTo NextRow
            End If

            ' 変数数をカウント
            GetVarCount = GetVarCount + 1

        NextRow:
        Next j
    Else
        ' varlistの各行をループ
        For j = 2 To tbl.ListColumns.Count
            ' "無効flag"が空白でない場合、この列の処理をスキップ
            disableFlag = tbl.DataBodyRange.Cells(1, j).Offset(-1, 0).Value
            If disableFlag <> "" Then
                GoTo NextColumn
            End If

            ' 変換元が空白またはNothingなら次の列へ
            replaceFrom = tbl.DataBodyRange.Cells(1, j).Value
            If IsEmpty(replaceFrom) Or replaceFrom = "" Then
                GoTo NextColumn
            End If

            ' 変数数をカウント
            GetVarCount = GetVarCount + 1

        NextColumn:
        Next j
    End If
End Function

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
            ' wb.Close False
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

Sub ProcessAsTextFile(wsTemplate As Worksheet, tbl As ListObject, i As Long, outputName As String, filePath As String, tableDirection As String)
    ' テンプレートシートの内容をテキストに変換
    Dim rng As Range
    Dim textOutput As String
    For Each rng In wsTemplate.UsedRange
        textOutput = textOutput & rng.Value & vbTab
        If rng.Column = wsTemplate.UsedRange.Columns.Count Then
            textOutput = textOutput & vbCrLf
        End If
    Next rng
    
    ' テーブルの方向で処理を分岐
    Dim j As Long
    Dim disableFlag As String
    Dim replaceFrom As String
    Dim replaceTo As String
    If tableDirection = "縦方向" Then
        ' 2列目以降の行をループ
        For j = 2 To tbl.ListRows.Count
            ' "無効flag"が空白でない場合、この行の処理をスキップ
            disableFlag = tbl.Range.Cells(j, 1).Offset(0, -1).Value
            If disableFlag <> "" Then
                GoTo NextRow
            End If
            
            ' 変換元と変換先の文字列を取得
            replaceFrom = tbl.DataBodyRange.Cells(j, 1).Value
            replaceTo = tbl.DataBodyRange.Cells(j, i).Value
            
            ' 変換元と変換先が空でない場合のみ置換を行う
            If Not IsEmpty(replaceFrom) And Not IsEmpty(replaceTo) Then
                ' テキスト内のreplaceFromをreplaceToに置換
                textOutput = Replace(textOutput, replaceFrom, replaceTo)
            End If
        NextRow:
        Next j
    Else
        ' 2列目以降の列をループ
        For j = 2 To tbl.ListColumns.Count
            ' "無効flag"が空白でない場合、この列の処理をスキップ
            disableFlag = tbl.DataBodyRange.Cells(1, j).Offset(-1, 0).Value
            If disableFlag <> "" Then
                GoTo NextColumn
            End If
            
            ' 変換元と変換先の文字列を取得
            replaceFrom = tbl.DataBodyRange.Cells(1, j).Value
            replaceTo = tbl.DataBodyRange.Cells(i, j).Value
            
            ' 変換元と変換先が空でない場合のみ置換を行う
            If Not IsEmpty(replaceFrom) And Not IsEmpty(replaceTo) Then
                ' テキスト内のreplaceFromをreplaceToに置換
                textOutput = Replace(textOutput, replaceFrom, replaceTo)
            End If
        NextColumn:
        Next j

    End If
        
    ' テキストファイル名を設定
    Dim fileName As String
    Dim uniqueNum As Integer
    uniqueNum = 0
    fileName = filePath & "\" & outputName
    
    ' ファイル名が重複している場合は通し番号を付与
    While Dir(fileName) <> ""
        uniqueNum = uniqueNum + 1
        Dim arr() As String
        arr = Split(outputName, ".")
        fileName = filePath & "\" & arr(0) & "(" & uniqueNum & ")." & arr(1)
    Wend
    
    ' テキストファイルに出力
    Open fileName For Output As #1
    Print #1, textOutput
    Close #1
    
    ' テキスト出力変数をリセット
    textOutput = ""
End Sub

Sub ProcessAsWorksheet(wsTemplate As Worksheet, tbl As ListObject, i As Long, outputName As String, filePath As String, tableDirection As String)
    ' 指定されたパスのWorkbookを開く
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
    
    
    ' テーブルの方向で処理を分岐
    Dim j As Long
    Dim disableFlag As String
    Dim replaceFrom As String
    Dim replaceTo As String
    Dim rng As Range
    If tableDirection = "縦方向" Then
        ' 2列目以降の行をループ
        For j = 2 To tbl.ListRows.Count
            ' "無効flag"が空白でない場合、この行の処理をスキップ
            disableFlag = tbl.Range.Cells(j, 1).Offset(0, -1).Value
            If disableFlag <> "" Then
                GoTo NextRow
            End If

            ' 変換元と変換先の文字列を取得
            replaceFrom = tbl.DataBodyRange.Cells(j, 1).Value
            replaceTo = tbl.DataBodyRange.Cells(j, i).Value

            ' 変換元と変換先が空でない場合のみ置換を行う
            If Not IsEmpty(replaceFrom) And Not IsEmpty(replaceTo) Then
                ' シート内の全てのセルを検索し、replaceFromをreplaceToに置換
                For Each rng In wsNew.UsedRange
                    If rng.Value <> "" Then
                        rng.Value = Replace(rng.Value, replaceFrom, replaceTo)
                    End If
                Next rng
            End If
        NextRow:
        Next j
    Else
        ' 2列目以降の列をループ
        For j = 2 To tbl.ListColumns.Count
            ' "無効flag"が空白でない場合、この列の処理をスキップ
            disableFlag = tbl.DataBodyRange.Cells(1, j).Offset(-1, 0).Value
            If disableFlag <> "" Then
                GoTo NextColumn
            End If

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
        NextColumn:
        Next j
    End If
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

Sub SetSheetNamesAsDropdownOptions()
    ' シートの数を取得
    Dim sheetCount As Integer
    sheetCount = ThisWorkbook.Sheets.Count
    
    ' シート名を保持するための配列を作成
    Dim sheetNames() As String
    ReDim sheetNames(1 To sheetCount - 1)  ' アクティブシートを除いた数で配列を初期化
    
    ' 各シートの名前を配列に格納
    Dim i As Integer
    Dim index As Integer
    index = 1
    For i = 1 To sheetCount
        If ThisWorkbook.Sheets(i).Name <> ActiveSheet.Name Then  ' アクティブシート以外の名前を追加
            sheetNames(index) = ThisWorkbook.Sheets(i).Name
            index = index + 1
        End If
    Next i
    
    ' データ検証リストのセルを指定
    Dim rng As Range
    Set rng = ActiveSheet.Range("template")  ' アクティブシートの"template"セルを指定
    
    ' すでにデータ検証が設定されている場合はそれを削除
    rng.Validation.Delete
    
    ' データ検証のリストを設定
    Dim strList As String
    strList = Join(sheetNames, ",")  ' 配列をカンマで連結した文字列に変換
    rng.Validation.Add Type:=xlValidateList, Formula1:=strList

    ' 処理結果をメッセージボックスで表示
    MsgBox "以下シート名をテンプレートリストに設定しました。" & vbCrLf & vbCrLf & Join(sheetNames, vbCrLf), vbInformation, "完了"
End Sub

' 行列を入れ替える関数
Sub TransposeTable()
    ' 実行するか確認
    If MsgBox("テーブルの行と列を入れ替えます。よろしいですか？", vbYesNo + vbQuestion, "確認") = vbNo Then Exit Sub

    ' 名前が付けられた範囲"varlist"を取得
    Dim wsMaster As Worksheet
    Set wsMaster = ActiveSheet
    Dim tbl As ListObject
    Set tbl = wsMaster.ListObjects("varlist")
    Dim rng As Range
    Set rng = tbl.DataBodyRange

    ' 元のデータを配列に読み込む
    Dim arrOriginal As Variant
    arrOriginal = rng.Value

    ' 行と列を入れ替えた新しい配列を作成する
    Dim arrTransposed As Variant
    ReDim arrTransposed(1 To UBound(arrOriginal, 2), 1 To UBound(arrOriginal, 1))
    Dim i As Long
    Dim j As Long
    For i = 1 To UBound(arrOriginal, 1)
        For j = 1 To UBound(arrOriginal, 2)
            arrTransposed(j, i) = arrOriginal(i, j)
        Next j
    Next i

    ' 新しい行数・列数に合わせてテーブルをリサイズ
    tbl.Resize tbl.Range.Resize(tbl.ListColumns.Count, tbl.ListRows.Count)

    ' 新しい配列を書き込む
    For i = 1 To UBound(arrTransposed, 1)
        For j = 1 To UBound(arrTransposed, 2)
            tbl.Range.Cells(i, j).Value = arrTransposed(i, j)
        Next j
    Next i

    ' directionセルの値を取得
    Dim rngDirection As Range
    Set rngDirection = wsMaster.Range("direction")
    
    ' directionセルの値に応じて値を書き換え
    If rngDirection.Value = "横方向" Then
        rngDirection.Value = "縦方向"
    ElseIf rngDirection.Value = "縦方向" Then
        rngDirection.Value = "横方向"
    Else
        ' directionセルの値が想定外の場合はエラーメッセージを表示
        MsgBox "directionセルの値が想定外です。値を確認してください。"
    End If

End Sub
