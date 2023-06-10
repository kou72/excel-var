Attribute VB_Name = "Module1"

Sub RunBothProcesses()
    ' スクリーン更新をオフにする
    Application.ScreenUpdating = False

    ' プロセスを実行する
    Call DeleteSheetsFromVarList
    Call CreateAndModifySheetsFromVarList

    ' スクリーン更新をオンに戻す
    Application.ScreenUpdating = True
End Sub

Sub CreateAndModifySheetsFromVarList()
    Dim wsMaster As Worksheet
    Dim wsTemplate As Worksheet
    Dim wsNew As Worksheet
    Dim rng As Range
    Dim i As Long
    Dim j As Long
    Dim templateName As String
    Dim outputName As String
    Dim outputType As String
    Dim replaceFrom As String
    Dim replaceTo As String
    Dim textOutput As String
    Dim fileName As String

    ' マスタシートを指定
    Set wsMaster = ThisWorkbook.Sheets("マスタ")
    
    ' varlistテーブルを指定
    Dim tbl As ListObject
    Set tbl = wsMaster.ListObjects("varlist")
    
    ' varlistの各行をループ
    For i = 1 To tbl.ListRows.Count
        ' テンプレートシート名と出力名を取得
        templateName = tbl.ListColumns("テンプレート").DataBodyRange.Cells(i).Value
        outputName = tbl.ListColumns("出力名").DataBodyRange.Cells(i).Value
        outputType = tbl.ListColumns("出力タイプ").DataBodyRange.Cells(i).Value
        
        ' outputNameが空白またはNothingなら次の行へ
        If IsEmpty(outputName) Or outputName = "" Then
            GoTo NextRow
        End If
        
        ' テンプレートシートを指定
        Set wsTemplate = ThisWorkbook.Sheets(templateName)
        
        If outputType = "textFile" Then

            ' テンプレートシートの内容をテキストに変換
            For Each rng In wsTemplate.UsedRange
                textOutput = textOutput & rng.Value & vbTab
                If rng.Column = wsTemplate.UsedRange.Columns.Count Then
                    textOutput = textOutput & vbCrLf
                End If
            Next rng
            
            ' 4列目以降の列をループ
            For j = 4 To tbl.ListColumns.Count
                ' 変換元と変換先の文字列を取得
                replaceFrom = tbl.HeaderRowRange.Cells(1, j).Value
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

        Else

            ' シートをコピーして新しいシートを作成
            wsTemplate.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
            
            ' 新しいシートを選択して名前を変更
            ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count).Name = outputName
            
            ' 新しいシートを変数にセット
            Set wsNew = ThisWorkbook.Sheets(outputName)
            
            ' 4列目以降の列をループ
            For j = 4 To tbl.ListColumns.Count
                ' 変換元と変換先の文字列を取得
                replaceFrom = tbl.HeaderRowRange.Cells(1, j).Value
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

        End If

NextRow:
    Next i
    
    ' マスタシートをアクティブにする
    wsMaster.Activate
End Sub

Sub DeleteSheetsFromVarList()
    Dim wsMaster As Worksheet
    Dim i As Long
    Dim outputName As String
    
    ' マスタシートを指定
    Set wsMaster = ThisWorkbook.Sheets("マスタ")
    
    ' varlistテーブルを指定
    Dim tbl As ListObject
    Set tbl = wsMaster.ListObjects("varlist")
    
    ' varlistの各行をループ
    For i = 1 To tbl.ListRows.Count
        ' 出力名を取得
        outputName = tbl.ListColumns("出力名").DataBodyRange.Cells(i).Value
        
        ' 出力名が空白またはNothingなら次の行へ
        If IsEmpty(outputName) Or outputName = "" Then
            GoTo NextRow
        End If
        
        ' 出力名に該当するシートを削除
        DeleteSheet outputName
        
NextRow:
    Next i
End Sub

Sub DeleteSheet(sheetName As String)
    Dim ws As Worksheet
    
    ' シートの存在確認
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    
    ' シートが存在する場合、削除
    If Not ws Is Nothing Then
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
    End If
End Sub
