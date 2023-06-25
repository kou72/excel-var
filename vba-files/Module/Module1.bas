Attribute VB_Name = "Module1"

Sub CreateAndModifySheetsFromVarList()
    ' �X�N���[���X�V���I�t�ɂ���
    Application.ScreenUpdating = False

    ' �}�X�^�����Z�b�g
    Dim wsMaster As Worksheet
    Set wsMaster = ActiveSheet

    Dim templateName As String
    templateName = wsMaster.Range("template").Value

    Dim outputType As String
    outputType = wsMaster.Range("type").Value

    Dim tbl As ListObject
    Set tbl = wsMaster.ListObjects("varlist")
    
    ' varlist�̊e�s�����[�v
    Dim i As Long
    For i = 2 To tbl.ListRows.Count
        ' �e���v���[�g�V�[�g���Əo�͖����擾
        Dim outputName As String
        outputName = tbl.ListColumns(1).DataBodyRange.Cells(i).Value
        
        ' outputName���󔒂܂���Nothing�Ȃ玟�̍s��
        If IsEmpty(outputName) Or outputName = "" Then
            GoTo NextRow
        End If
        
        ' �e���v���[�g�V�[�g���w��
        Dim wsTemplate As Worksheet
        Set wsTemplate = ThisWorkbook.Sheets(templateName)
        
        If outputType = "textFile" Then
            ProcessAsTextFile wsTemplate, tbl, i, outputName
        Else
            ProcessAsWorksheet wsTemplate, tbl, i, outputName
        End If

    NextRow:
    Next i
    
    ' �}�X�^�V�[�g���A�N�e�B�u�ɂ���
    wsMaster.Activate

    ' �X�N���[���X�V���I���ɖ߂�
    Application.ScreenUpdating = True
End Sub

Sub ProcessAsTextFile(wsTemplate As Worksheet, tbl As ListObject, i As Long, outputName As String)
    Dim rng As Range
    Dim j As Long
    Dim replaceFrom As String
    Dim replaceTo As String
    Dim textOutput As String
    Dim fileName As String
    
    ' �e���v���[�g�V�[�g�̓��e���e�L�X�g�ɕϊ�
    For Each rng In wsTemplate.UsedRange
        textOutput = textOutput & rng.Value & vbTab
        If rng.Column = wsTemplate.UsedRange.Columns.Count Then
            textOutput = textOutput & vbCrLf
        End If
    Next rng
    
    ' 2��ڈȍ~�̗�����[�v
    For j = 2 To tbl.ListColumns.Count
        ' �ϊ����ƕϊ���̕�������擾
        replaceFrom = tbl.DataBodyRange.Cells(1, j).Value
        replaceTo = tbl.DataBodyRange.Cells(i, j).Value
        
        ' �ϊ����ƕϊ��悪��łȂ��ꍇ�̂ݒu�����s��
        If Not IsEmpty(replaceFrom) And Not IsEmpty(replaceTo) Then
            ' �e�L�X�g����replaceFrom��replaceTo�ɒu��
            textOutput = Replace(textOutput, replaceFrom, replaceTo)
        End If
    Next j
    
    ' �e�L�X�g�t�@�C������ݒ�
    fileName = ThisWorkbook.Path & "\" & outputName
    
    ' �e�L�X�g�t�@�C���ɏo��
    Open fileName For Output As #1
    Print #1, textOutput
    Close #1
    
    ' �e�L�X�g�o�͕ϐ������Z�b�g
    textOutput = ""
End Sub

Sub ProcessAsWorksheet(wsTemplate As Worksheet, tbl As ListObject, i As Long, outputName As String)
    Dim wsNew As Worksheet
    Dim rng As Range
    Dim j As Long
    Dim replaceFrom As String
    Dim replaceTo As String
    
    ' �V�[�g���R�s�[���ĐV�����V�[�g���쐬
    wsTemplate.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
    
    ' �V�����V�[�g��I�����Ė��O��ύX
    ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count).Name = outputName
    
    ' �V�����V�[�g��ϐ��ɃZ�b�g
    Set wsNew = ThisWorkbook.Sheets(outputName)
    
    ' 2��ڈȍ~�̗�����[�v
    For j = 2 To tbl.ListColumns.Count
        ' �ϊ����ƕϊ���̕�������擾
        replaceFrom = tbl.DataBodyRange.Cells(1, j).Value
        replaceTo = tbl.DataBodyRange.Cells(i, j).Value
        
        ' �ϊ����ƕϊ��悪��łȂ��ꍇ�̂ݒu�����s��
        If Not IsEmpty(replaceFrom) And Not IsEmpty(replaceTo) Then
            ' �V�[�g���̑S�ẴZ�����������AreplaceFrom��replaceTo�ɒu��
            For Each rng In wsNew.UsedRange
                If rng.Value <> "" Then
                    rng.Value = Replace(rng.Value, replaceFrom, replaceTo)
                End If
            Next rng
        End If
    Next j
End Sub

Sub SelectFileOrFolderAndWritePath()
    ' "type"�Ƃ������O�̃Z���̓��e���擾
    Dim selectType As Range
    Set selectType = ThisWorkbook.Names("type").RefersToRange

    ' "type"�̓��e�Ɋ�Â���FileDialog�̃^�C�v��ݒ�
    Dim fd As FileDialog
    If selectType.Value = "sheet" Then
        Set fd = Application.FileDialog(msoFileDialogFilePicker)
    ElseIf selectType.Value = "textFile" Then
        Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    Else
        MsgBox "���O�t���Z�� 'type' �̒l�������ł�"
        Exit Sub
    End If

    ' �_�C�A���O��\�����A�I�������p�X���擾
    Dim selectedPath As String
    With fd.Title = "Select Path".AllowMultiSelect = False
        If .Show = True Then
            selectedPath = .SelectedItems(1)
        End If
    End With

    ' �I�������p�X�� "path" �Ƃ������O�̃Z���ɏ�������
    Dim rng As Range
    Set rng = ThisWorkbook.Names("path").RefersToRange
    rng.Value = selectedPath
End Sub