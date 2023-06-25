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
    Dim filePath As String
    filePath = ThisWorkbook.Names("path").RefersToRange.Value
    Dim tableDirection As String
    tableDirection = ThisWorkbook.Names("direction").RefersToRange.Value
    Dim tbl As ListObject
    Set tbl = wsMaster.ListObjects("varlist")

    ' ���s�O�Ɋm�F���b�Z�[�W��\��
    Dim msg As String
    msg = "�ȉ��̓��e�ŏ��������s���܂��B" & vbCrLf & vbCrLf & _
        "�e���v���[�g�V�[�g�F" & templateName & vbCrLf & _
        "�o�͌`���F" & outputType & vbCrLf & _
        "�o�͐�F" & filePath & vbCrLf & _
        "�o�͐��F" & GetOutputCount(tbl, tableDirection) & vbCrLf & _
        "�ϐ����F" & GetVarCount(tbl, tableDirection) & vbCrLf & _
        "��낵���ł����H"
    If MsgBox(msg, vbYesNo + vbQuestion, "�m�F") = vbNo Then Exit Sub

    ' �o�͐悪�L���ȃt�H���_�܂���Excel�t�@�C�����w���Ă��邩�m�F
    If Not CheckFilePath(filePath, outputType) Then Exit Sub

    ' �e�[�u���̕����ŏ����𕪊�
    Dim i As Long
    Dim disableFlag As String
    Dim outputName As String
    Dim wsTemplate As Worksheet
    If tableDirection = "�c����" Then
        ' varlist�̊e������[�v
        For i = 2 To tbl.ListColumns.Count
            ' "����flag"���󔒂łȂ��ꍇ�A���̗�̏������X�L�b�v
            disableFlag = tbl.Range.Cells(1, i).Offset(-1, 0).Value
            If disableFlag <> "" Then
                GoTo NextColumn
            End If

            ' �e���v���[�g�V�[�g���Əo�͖����擾
            outputName = tbl.ListColumns(i).DataBodyRange.Cells(1).Value

            ' outputName���󔒂܂���Nothing�Ȃ玟�̗��
            If IsEmpty(outputName) Or outputName = "" Then
                GoTo NextColumn
            End If

            ' �e���v���[�g�V�[�g���w��
            Set wsTemplate = ThisWorkbook.Sheets(templateName)
            
            If outputType = "textFile" Then
                ProcessAsTextFile wsTemplate, tbl, i, outputName, filePath, tableDirection
            Else
                ProcessAsWorksheet wsTemplate, tbl, i, outputName, filePath, tableDirection
            End If
        NextColumn:
        Next i
    Else
        ' varlist�̊e�s�����[�v
        For i = 2 To tbl.ListRows.Count
            ' "����flag"���󔒂łȂ��ꍇ�A���̍s�̏������X�L�b�v
            disableFlag = tbl.Range.Cells(i, 1).Offset(0, -1).Value
            If disableFlag <> "" Then
                GoTo NextRow
            End If

            ' �e���v���[�g�V�[�g���Əo�͖����擾
            outputName = tbl.ListColumns(1).DataBodyRange.Cells(i).Value
            
            ' outputName���󔒂܂���Nothing�Ȃ玟�̍s��
            If IsEmpty(outputName) Or outputName = "" Then
                GoTo NextRow
            End If

            ' �e���v���[�g�V�[�g���w��
            Set wsTemplate = ThisWorkbook.Sheets(templateName)
            
            If outputType = "textFile" Then
                ProcessAsTextFile wsTemplate, tbl, i, outputName, filePath, tableDirection
            Else
                ProcessAsWorksheet wsTemplate, tbl, i, outputName, filePath, tableDirection
            End If

        NextRow:
        Next i

    End If

    ' �X�N���[���X�V���I���ɖ߂�
    Application.ScreenUpdating = True
End Sub

' �o�͐����擾����֐�
Function GetOutputCount(tbl As ListObject, tableDirection As String) As Long
    ' �o�͐���������
    GetOutputCount = 0

    ' �e�[�u���̕����ŏ����𕪊�
    Dim i As Long
    Dim disableFlag As String
    Dim outputName As String
    If tableDirection = "�c����" Then
        ' varlist�̊e�s�����[�v
        For i = 2 To tbl.ListColumns.Count
            ' "����flag"���󔒂łȂ��ꍇ�A���̗�̏������X�L�b�v
            disableFlag = tbl.Range.Cells(1, i).Offset(-1, 0).Value
            If disableFlag <> "" Then
                GoTo NextColumn
            End If

            ' outputName���󔒂܂���Nothing�Ȃ玟�̗��
            outputName = tbl.ListColumns(i).DataBodyRange.Cells(1).Value
            If IsEmpty(outputName) Or outputName = "" Then
                GoTo NextColumn
            End If
            
            ' �o�͐����J�E���g
            GetOutputCount = GetOutputCount + 1

        NextColumn:
        Next i
    Else
        ' varlist�̊e�s�����[�v
        For i = 2 To tbl.ListRows.Count
            ' "����flag"���󔒂łȂ��ꍇ�A���̍s�̏������X�L�b�v
            disableFlag = tbl.Range.Cells(i, 1).Offset(0, -1).Value
            If disableFlag <> "" Then
                GoTo NextRow
            End If

            ' outputName���󔒂܂���Nothing�Ȃ玟�̍s��
            outputName = tbl.ListColumns(1).DataBodyRange.Cells(i).Value
            If IsEmpty(outputName) Or outputName = "" Then
                GoTo NextRow
            End If
            
            ' �o�͐����J�E���g
            GetOutputCount = GetOutputCount + 1

        NextRow:
        Next i
    End If
End Function

' �ϐ������擾����֐�
Function GetVarCount(tbl As ListObject, tableDirection As String) As Long
    ' �ϐ�����������
    GetVarCount = 0

    ' �e�[�u���̕����ŏ����𕪊�
    Dim j As Long
    Dim disableFlag As String
    Dim replaceFrom As String
    If tableDirection = "�c����" Then
        ' varlist�̊e�s�����[�v
        For j = 2 To tbl.ListRows.Count
            ' "����flag"���󔒂łȂ��ꍇ�A���̍s�̏������X�L�b�v
            disableFlag = tbl.Range.Cells(j, 1).Offset(0, -1).Value
            If disableFlag <> "" Then
                GoTo NextRow
            End If

            ' �ϊ������󔒂܂���Nothing�Ȃ玟�̍s��
            replaceFrom = tbl.DataBodyRange.Cells(j, 1).Value
            If IsEmpty(replaceFrom) Or replaceFrom = "" Then
                GoTo NextRow
            End If

            ' �ϐ������J�E���g
            GetVarCount = GetVarCount + 1

        NextRow:
        Next j
    Else
        ' varlist�̊e�s�����[�v
        For j = 2 To tbl.ListColumns.Count
            ' "����flag"���󔒂łȂ��ꍇ�A���̗�̏������X�L�b�v
            disableFlag = tbl.DataBodyRange.Cells(1, j).Offset(-1, 0).Value
            If disableFlag <> "" Then
                GoTo NextColumn
            End If

            ' �ϊ������󔒂܂���Nothing�Ȃ玟�̗��
            replaceFrom = tbl.DataBodyRange.Cells(1, j).Value
            If IsEmpty(replaceFrom) Or replaceFrom = "" Then
                GoTo NextColumn
            End If

            ' �ϐ������J�E���g
            GetVarCount = GetVarCount + 1

        NextColumn:
        Next j
    End If
End Function

' filePath���L���ȃt�H���_�܂���Excel�t�@�C�����w���Ă��邩�m�F����֐�
Function CheckFilePath(filePath As String, outputType As String) As Boolean
    If outputType = "sheet" Then ' Excel�t�@�C���̃`�F�b�N
        On Error Resume Next ' �G���[�n���h����L���ɂ���
        Dim wb As Workbook
        Set wb = Workbooks.Open(filePath) ' filePath���J��
        On Error GoTo 0 ' �G���[�n���h���𖳌��ɂ���
        If wb Is Nothing Then ' filePath�������Ȃ�G���[���b�Z�[�W��\�����ďI��
            MsgBox "�I�����ꂽ�t�@�C���������ł��B�L����Excel�t�@�C����I�����Ă��������B", vbCritical, "�G���["
            CheckFilePath = False
        Else ' filePath���L���Ȃ�True��Ԃ�
            ' wb.Close False
            CheckFilePath = True
        End If
    ElseIf outputType = "textFile" Then ' �t�H���_�̃`�F�b�N
        If Dir(filePath, vbDirectory) = "" Then
            MsgBox "�I�����ꂽ�p�X�������ł��B�L���ȃt�H���_��I�����Ă��������B", vbCritical, "�G���["
            CheckFilePath = False
        Else
            CheckFilePath = True
        End If
    End If
End Function

Sub ProcessAsTextFile(wsTemplate As Worksheet, tbl As ListObject, i As Long, outputName As String, filePath As String, tableDirection As String)
    ' �e���v���[�g�V�[�g�̓��e���e�L�X�g�ɕϊ�
    Dim rng As Range
    Dim textOutput As String
    For Each rng In wsTemplate.UsedRange
        textOutput = textOutput & rng.Value & vbTab
        If rng.Column = wsTemplate.UsedRange.Columns.Count Then
            textOutput = textOutput & vbCrLf
        End If
    Next rng
    
    ' �e�[�u���̕����ŏ����𕪊�
    Dim j As Long
    Dim disableFlag As String
    Dim replaceFrom As String
    Dim replaceTo As String
    If tableDirection = "�c����" Then
        ' 2��ڈȍ~�̍s�����[�v
        For j = 2 To tbl.ListRows.Count
            ' "����flag"���󔒂łȂ��ꍇ�A���̍s�̏������X�L�b�v
            disableFlag = tbl.Range.Cells(j, 1).Offset(0, -1).Value
            If disableFlag <> "" Then
                GoTo NextRow
            End If
            
            ' �ϊ����ƕϊ���̕�������擾
            replaceFrom = tbl.DataBodyRange.Cells(j, 1).Value
            replaceTo = tbl.DataBodyRange.Cells(j, i).Value
            
            ' �ϊ����ƕϊ��悪��łȂ��ꍇ�̂ݒu�����s��
            If Not IsEmpty(replaceFrom) And Not IsEmpty(replaceTo) Then
                ' �e�L�X�g����replaceFrom��replaceTo�ɒu��
                textOutput = Replace(textOutput, replaceFrom, replaceTo)
            End If
        NextRow:
        Next j
    Else
        ' 2��ڈȍ~�̗�����[�v
        For j = 2 To tbl.ListColumns.Count
            ' "����flag"���󔒂łȂ��ꍇ�A���̗�̏������X�L�b�v
            disableFlag = tbl.DataBodyRange.Cells(1, j).Offset(-1, 0).Value
            If disableFlag <> "" Then
                GoTo NextColumn
            End If
            
            ' �ϊ����ƕϊ���̕�������擾
            replaceFrom = tbl.DataBodyRange.Cells(1, j).Value
            replaceTo = tbl.DataBodyRange.Cells(i, j).Value
            
            ' �ϊ����ƕϊ��悪��łȂ��ꍇ�̂ݒu�����s��
            If Not IsEmpty(replaceFrom) And Not IsEmpty(replaceTo) Then
                ' �e�L�X�g����replaceFrom��replaceTo�ɒu��
                textOutput = Replace(textOutput, replaceFrom, replaceTo)
            End If
        NextColumn:
        Next j

    End If
        
    ' �e�L�X�g�t�@�C������ݒ�
    Dim fileName As String
    Dim uniqueNum As Integer
    uniqueNum = 0
    fileName = filePath & "\" & outputName
    
    ' �t�@�C�������d�����Ă���ꍇ�͒ʂ��ԍ���t�^
    While Dir(fileName) <> ""
        uniqueNum = uniqueNum + 1
        Dim arr() As String
        arr = Split(outputName, ".")
        fileName = filePath & "\" & arr(0) & "(" & uniqueNum & ")." & arr(1)
    Wend
    
    ' �e�L�X�g�t�@�C���ɏo��
    Open fileName For Output As #1
    Print #1, textOutput
    Close #1
    
    ' �e�L�X�g�o�͕ϐ������Z�b�g
    textOutput = ""
End Sub

Sub ProcessAsWorksheet(wsTemplate As Worksheet, tbl As ListObject, i As Long, outputName As String, filePath As String, tableDirection As String)
    ' �w�肳�ꂽ�p�X��Workbook���J��
    Dim wbTarget As Workbook
    Set wbTarget = Workbooks.Open(filePath)
    
    ' �d�����Ȃ��V�[�g���������邽�߂̃��[�v
    Dim suffix As Integer
    suffix = 0
    Dim newSheetName As String
    newSheetName = outputName
    Do While WorksheetExists(wbTarget, newSheetName)
        suffix = suffix + 1
        newSheetName = outputName & " (" & suffix & ")"
    Loop
    
    ' �V�[�g���R�s�[���ĐV�����V�[�g���쐬���A���O��ύX
    wsTemplate.Copy After:=wbTarget.Sheets(wbTarget.Sheets.Count)
    wbTarget.Sheets(wbTarget.Sheets.Count).Name = newSheetName
    Dim wsNew As Worksheet
    Set wsNew = wbTarget.Sheets(newSheetName)
    
    
    ' �e�[�u���̕����ŏ����𕪊�
    Dim j As Long
    Dim disableFlag As String
    Dim replaceFrom As String
    Dim replaceTo As String
    Dim rng As Range
    If tableDirection = "�c����" Then
        ' 2��ڈȍ~�̍s�����[�v
        For j = 2 To tbl.ListRows.Count
            ' "����flag"���󔒂łȂ��ꍇ�A���̍s�̏������X�L�b�v
            disableFlag = tbl.Range.Cells(j, 1).Offset(0, -1).Value
            If disableFlag <> "" Then
                GoTo NextRow
            End If

            ' �ϊ����ƕϊ���̕�������擾
            replaceFrom = tbl.DataBodyRange.Cells(j, 1).Value
            replaceTo = tbl.DataBodyRange.Cells(j, i).Value

            ' �ϊ����ƕϊ��悪��łȂ��ꍇ�̂ݒu�����s��
            If Not IsEmpty(replaceFrom) And Not IsEmpty(replaceTo) Then
                ' �V�[�g���̑S�ẴZ�����������AreplaceFrom��replaceTo�ɒu��
                For Each rng In wsNew.UsedRange
                    If rng.Value <> "" Then
                        rng.Value = Replace(rng.Value, replaceFrom, replaceTo)
                    End If
                Next rng
            End If
        NextRow:
        Next j
    Else
        ' 2��ڈȍ~�̗�����[�v
        For j = 2 To tbl.ListColumns.Count
            ' "����flag"���󔒂łȂ��ꍇ�A���̗�̏������X�L�b�v
            disableFlag = tbl.DataBodyRange.Cells(1, j).Offset(-1, 0).Value
            If disableFlag <> "" Then
                GoTo NextColumn
            End If

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
        NextColumn:
        Next j
    End If
End Sub

' �V�[�g�����݂��邩�ǂ������m�F���邽�߂̊֐�
Function WorksheetExists(wb As Workbook, wsName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Sheets(wsName)
    On Error GoTo 0
    WorksheetExists = Not ws Is Nothing
End Function


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
    With fd
        .Title = "Select Path"
        .AllowMultiSelect = False
        
        If .Show = True Then
            selectedPath = .SelectedItems(1)
        End If
    End With

    ' �I�������p�X�� "path" �Ƃ������O�̃Z���ɏ�������
    Dim rng As Range
    Set rng = ThisWorkbook.Names("path").RefersToRange
    rng.Value = selectedPath
End Sub

Sub SetSheetNamesAsDropdownOptions()
    ' �V�[�g�̐����擾
    Dim sheetCount As Integer
    sheetCount = ThisWorkbook.Sheets.Count
    
    ' �V�[�g����ێ����邽�߂̔z����쐬
    Dim sheetNames() As String
    ReDim sheetNames(1 To sheetCount - 1)  ' �A�N�e�B�u�V�[�g�����������Ŕz���������
    
    ' �e�V�[�g�̖��O��z��Ɋi�[
    Dim i As Integer
    Dim index As Integer
    index = 1
    For i = 1 To sheetCount
        If ThisWorkbook.Sheets(i).Name <> ActiveSheet.Name Then  ' �A�N�e�B�u�V�[�g�ȊO�̖��O��ǉ�
            sheetNames(index) = ThisWorkbook.Sheets(i).Name
            index = index + 1
        End If
    Next i
    
    ' �f�[�^���؃��X�g�̃Z�����w��
    Dim rng As Range
    Set rng = ActiveSheet.Range("template")  ' �A�N�e�B�u�V�[�g��"template"�Z�����w��
    
    ' ���łɃf�[�^���؂��ݒ肳��Ă���ꍇ�͂�����폜
    rng.Validation.Delete
    
    ' �f�[�^���؂̃��X�g��ݒ�
    Dim strList As String
    strList = Join(sheetNames, ",")  ' �z����J���}�ŘA������������ɕϊ�
    rng.Validation.Add Type:=xlValidateList, Formula1:=strList

    ' �������ʂ����b�Z�[�W�{�b�N�X�ŕ\��
    MsgBox "�ȉ��V�[�g�����e���v���[�g���X�g�ɐݒ肵�܂����B" & vbCrLf & vbCrLf & Join(sheetNames, vbCrLf), vbInformation, "����"
End Sub

' �s������ւ���֐�
Sub TransposeTable()
    ' ���s���邩�m�F
    If MsgBox("�e�[�u���̍s�Ɨ�����ւ��܂��B��낵���ł����H", vbYesNo + vbQuestion, "�m�F") = vbNo Then Exit Sub

    ' ���O���t����ꂽ�͈�"varlist"���擾
    Dim wsMaster As Worksheet
    Set wsMaster = ActiveSheet
    Dim tbl As ListObject
    Set tbl = wsMaster.ListObjects("varlist")
    Dim rng As Range
    Set rng = tbl.DataBodyRange

    ' ���̃f�[�^��z��ɓǂݍ���
    Dim arrOriginal As Variant
    arrOriginal = rng.Value

    ' �s�Ɨ�����ւ����V�����z����쐬����
    Dim arrTransposed As Variant
    ReDim arrTransposed(1 To UBound(arrOriginal, 2), 1 To UBound(arrOriginal, 1))
    Dim i As Long
    Dim j As Long
    For i = 1 To UBound(arrOriginal, 1)
        For j = 1 To UBound(arrOriginal, 2)
            arrTransposed(j, i) = arrOriginal(i, j)
        Next j
    Next i

    ' �V�����s���E�񐔂ɍ��킹�ăe�[�u�������T�C�Y
    tbl.Resize tbl.Range.Resize(tbl.ListColumns.Count, tbl.ListRows.Count)

    ' �V�����z�����������
    For i = 1 To UBound(arrTransposed, 1)
        For j = 1 To UBound(arrTransposed, 2)
            tbl.Range.Cells(i, j).Value = arrTransposed(i, j)
        Next j
    Next i

    ' direction�Z���̒l���擾
    Dim rngDirection As Range
    Set rngDirection = wsMaster.Range("direction")
    
    ' direction�Z���̒l�ɉ����Ēl����������
    If rngDirection.Value = "������" Then
        rngDirection.Value = "�c����"
    ElseIf rngDirection.Value = "�c����" Then
        rngDirection.Value = "������"
    Else
        ' direction�Z���̒l���z��O�̏ꍇ�̓G���[���b�Z�[�W��\��
        MsgBox "direction�Z���̒l���z��O�ł��B�l���m�F���Ă��������B"
    End If

End Sub
