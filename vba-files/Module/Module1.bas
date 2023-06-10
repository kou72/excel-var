Attribute VB_Name = "Module1"
Sub sample()
  MsgBox "hello vba"
End Sub

Sub CreateAndModifySheets()
    Dim i As Integer
    Dim ws As Worksheet
    Dim rng As Range
    
    ' 元のシートを指定
    Set ws = ThisWorkbook.Sheets("Sheet1")
    
    ' 繰り返し処理を開始
    For i = 1 To 3
        ' シートをコピーして新しいシートを作成
        ws.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
        
        ' 新しいシートを選択して名前を変更
        ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count).Name = "test-sheet" & i
        
        ' 新しいシートを変数にセット
        Set wsNew = ThisWorkbook.Sheets("test-sheet" & i)
        
        ' シート内の全てのセルを検索し、$var を test1 に置換
        For Each rng In wsNew.UsedRange
            If rng.Value <> "" Then
                rng.Value = Replace(rng.Value, "$var", "test1")
            End If
        Next rng
    Next i
End Sub