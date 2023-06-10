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

    ' �}�X�^�V�[�g���w��
    Set wsMaster = ThisWorkbook.Sheets("�}�X�^")
    
    ' varlist�e�[�u�����w��
    Dim tbl As ListObject
    Set tbl = wsMaster.ListObjects("varlist")
    
    ' varlist�̊e�s�����[�v
    For i = 1 To tbl.ListRows.Count
        ' �e���v���[�g�V�[�g���Əo�͖����擾
        templateName = tbl.ListColumns("�e���v���[�g").DataBodyRange.Cells(i).Value
        outputName = tbl.ListColumns("�o�͖�").DataBodyRange.Cells(i).Value
        
        ' �e���v���[�g�V�[�g���w��
        Set wsTemplate = ThisWorkbook.Sheets(templateName)
        
        ' �V�[�g���R�s�[���ĐV�����V�[�g���쐬
        wsTemplate.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
        
        ' �V�����V�[�g��I�����Ė��O��ύX
        ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count).Name = outputName
        
        ' �V�����V�[�g��ϐ��ɃZ�b�g
        Set wsNew = ThisWorkbook.Sheets(outputName)
        
        ' 4��ڈȍ~�̗�����[�v
        For j = 4 To tbl.ListColumns.Count
            ' �ϊ����ƕϊ���̕�������擾
            replaceFrom = tbl.HeaderRowRange.Cells(1, j).Value
            replaceTo = tbl.DataBodyRange.Cells(i, j).Value
            
            ' �V�[�g���̑S�ẴZ�����������AreplaceFrom��replaceTo�ɒu��
            For Each rng In wsNew.UsedRange
                If rng.Value <> "" Then
                    rng.Value = Replace(rng.Value, replaceFrom, replaceTo)
                End If
            Next rng
        Next j
    Next i
End Sub

