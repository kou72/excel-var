Attribute VB_Name = "Module1"

Sub RunBothProcesses()
    ' �X�N���[���X�V���I�t�ɂ���
    Application.ScreenUpdating = False

    ' �v���Z�X�����s����
    Call DeleteSheetsFromVarList
    Call CreateAndModifySheetsFromVarList

    ' �X�N���[���X�V���I���ɖ߂�
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
        outputType = tbl.ListColumns("�o�̓^�C�v").DataBodyRange.Cells(i).Value
        
        ' outputName���󔒂܂���Nothing�Ȃ玟�̍s��
        If IsEmpty(outputName) Or outputName = "" Then
            GoTo NextRow
        End If
        
        ' �e���v���[�g�V�[�g���w��
        Set wsTemplate = ThisWorkbook.Sheets(templateName)
        
        If outputType = "textFile" Then

            ' �e���v���[�g�V�[�g�̓��e���e�L�X�g�ɕϊ�
            For Each rng In wsTemplate.UsedRange
                textOutput = textOutput & rng.Value & vbTab
                If rng.Column = wsTemplate.UsedRange.Columns.Count Then
                    textOutput = textOutput & vbCrLf
                End If
            Next rng
            
            ' 4��ڈȍ~�̗�����[�v
            For j = 4 To tbl.ListColumns.Count
                ' �ϊ����ƕϊ���̕�������擾
                replaceFrom = tbl.HeaderRowRange.Cells(1, j).Value
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

        Else

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

        End If

NextRow:
    Next i
    
    ' �}�X�^�V�[�g���A�N�e�B�u�ɂ���
    wsMaster.Activate
End Sub

Sub DeleteSheetsFromVarList()
    Dim wsMaster As Worksheet
    Dim i As Long
    Dim outputName As String
    
    ' �}�X�^�V�[�g���w��
    Set wsMaster = ThisWorkbook.Sheets("�}�X�^")
    
    ' varlist�e�[�u�����w��
    Dim tbl As ListObject
    Set tbl = wsMaster.ListObjects("varlist")
    
    ' varlist�̊e�s�����[�v
    For i = 1 To tbl.ListRows.Count
        ' �o�͖����擾
        outputName = tbl.ListColumns("�o�͖�").DataBodyRange.Cells(i).Value
        
        ' �o�͖����󔒂܂���Nothing�Ȃ玟�̍s��
        If IsEmpty(outputName) Or outputName = "" Then
            GoTo NextRow
        End If
        
        ' �o�͖��ɊY������V�[�g���폜
        DeleteSheet outputName
        
NextRow:
    Next i
End Sub

Sub DeleteSheet(sheetName As String)
    Dim ws As Worksheet
    
    ' �V�[�g�̑��݊m�F
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    
    ' �V�[�g�����݂���ꍇ�A�폜
    If Not ws Is Nothing Then
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
    End If
End Sub
