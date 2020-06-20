Attribute VB_Name = "PasteName"
Option Explicit

Dim nameNo As Integer         '�\��t���閼�O�͉����ڂ��̃J�E���g
Dim ROW_OF_NAME As Integer    '1�l�̖��O�Ɏg���s��

'***************************************************************************************************
'* �@�\�����F���O�\��t�������i�N���X�P�ʁj
'* ���ӎ����F�Ȃ�
'***************************************************************************************************
Sub pstNameStudents(students As Collection)
    
    '������
    pstStudentsNameInit
    
    Dim student As student
    nameNo = 0
    
    '�擾�������k�����J��Ԃ�
    For Each student In students
        nameNo = nameNo + 1
        pstNameStudent student
    Next
    

End Sub

Sub pstStudentsNameInit()
    
    ROW_OF_NAME = 2
    
End Sub


'***************************************************************************************************
'* �@�\�����F���O�\��t�������i���k�P�ʁj
'* ���ӎ����F�Ȃ�
'***************************************************************************************************
Sub pstNameStudent(student As student)

    Dim pstCellAddress As String
    
    '���O�̓\��t���ʒu�擾
    pstCellAddress = getPastNameCellAddress
    
    '���O��\��t��
    pstName pstCellAddress, student
    

End Sub

'***************************************************************************************************
'* �@�\�����F���O�ʒu�擾
'* ���ӎ����F�Ȃ�
'***************************************************************************************************
Function getPastNameCellAddress()
    
    Dim colIdx As Integer
    Dim colIdxTmp As Integer
    Dim rowIdx As Integer
    Dim rowIdxTmp As Integer
    Dim tmp As Integer
    
    '**************
    '�s�ԍ�
    '**************
    
    '���s�ڂ̖��O��
    '�\��t���閼�O�͉��ڂ��̃J�E���g / ���O�𒣂�t����� �̐؂�グ
    rowIdxTmp = Application.WorksheetFunction.RoundUp(nameNo / COL_CNT, 0)

    '���\��t����s�܂łɎg�p�����s��(1���O�ɂ�2�s�g��)��ݒ肷��
    If (rowIdxTmp = 1) Then
        '��s�ڂ̏ꍇ
        rowIdx = 0
    Else
        rowIdx = ((rowIdxTmp - 1) * ROW_OF_NAME)
    End If
    
    '�w�b�_�[�𑫂�
    rowIdx = rowIdx + ROW_OF_HEADDER
    
    '���ׂĂ̎ʐ^�Ɏg���s���𑫂�
    rowIdx = rowIdx + ROW_OF_ALL_IMG
    
    '�ʐ^�g�Ɩ��O�g�̊Ԃ�1�s�𑫂�
    rowIdx = rowIdx + 1
    
    '���̎��̍s�ɓ\��t����
    rowIdx = rowIdx + 1
    
    
    '**************
    '��ԍ�
    '**************
    
    '����ڂ̖��O��
    '�\��t���閼�O�͉����ڂ��̃J�E���g / ���O�𒣂�t����� �̂��܂�
    colIdxTmp = nameNo Mod COL_CNT
    If (colIdxTmp = 0) Then
        '0 �̏ꍇ�͈�ԍŌ�̗�
        colIdxTmp = COL_CNT
    End If

    '���\��t�����܂łɎg�p������(1���O�ɂ�5��g��)��ݒ肷��
    If (colIdxTmp = 1) Then
        '���ڂ̏ꍇ
        colIdx = 0
    Else
        colIdx = ((colIdxTmp - 1) * COL_OF_IMG)
    End If
    
    '������\�閼�O�̍����1���񂪂���̂ő���
    colIdx = colIdx + 1
    
    '�ŏ���A��𑫂�
    colIdx = colIdx + COL_OF_LEFT
    
    '���̎��̗�ɓ\��t����
    colIdx = colIdx + 1
    
    
    '**************
    '�\��t����Z����ݒ�
    '**************
    'Cells �� Address ���擾�ł��Ȃ��̂ŁAA1 �Z������ Offset �Ŗ������擾����
    getPastNameCellAddress = Range("A1").Offset(rowIdx - 1, colIdx - 1).Address

End Function

'***************************************************************************************************
'* �@�\�����F���O�\��t��
'* ���ӎ����F�Ȃ�
'***************************************************************************************************
Sub pstName(cellAddress As String, student As student)

    Dim cellOfNo As String
    Dim cellOfPhonetic As String
    Dim cellOfName As String
    Dim cellOfSex As String
    
    
    cellOfNo = cellAddress                                  'cellAddress���ԍ��Z���ɂȂ�
    cellOfPhonetic = Range(cellOfNo).Offset(0, 1).Address   '�ԍ��Z���̉E
    cellOfName = Range(cellOfPhonetic).Offset(1, 0).Address '�ӂ肪�ȃZ���̉�
    cellOfSex = Range(cellOfPhonetic).Offset(0, 1).Address  '�ӂ肪�ȃZ���̉E
    
    
    outputSheet.Range(cellOfNo).MergeArea.value = student.�ԍ�
    outputSheet.Range(cellOfPhonetic).value = student.�ӂ肪��
    outputSheet.Range(cellOfName).value = student.���O
    outputSheet.Range(cellOfSex).MergeArea.value = student.����
    
End Sub
