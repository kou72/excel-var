Attribute VB_Name = "Module1"
Sub CreateAndModifySheetsFromVarList()
    Dim wsMaster As Worksheet
    Dim wsTemplate As Worksheet
    Dim wsNew As Worksheet
    Dim rng As Range
    Dim i As Long
    Dim j As Long
    Dim templateName As String
    Dim outputName As String
    Dim replaceFrom As String
    Dim replaceTo As String

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
        
        ' テンプレートシートを指定
        Set wsTemplate = ThisWorkbook.Sheets(templateName)
        
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
            
            ' シート内の全てのセルを検索し、replaceFromをreplaceToに置換
            For Each rng In wsNew.UsedRange
                If rng.Value <> "" Then
                    rng.Value = Replace(rng.Value, replaceFrom, replaceTo)
                End If
            Next rng
        Next j
    Next i
End Sub

