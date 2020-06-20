Attribute VB_Name = "PasteImage"
Option Explicit

Public ROW_CNT As Integer        '�ʐ^�𒣂�t����s��
Public COL_CNT As Integer        '�ʐ^�𒣂�t�����
Public ROW_OF_IMG As Integer     '1���̎ʐ^�Ɏg���s��
Public COL_OF_IMG As Integer     '1���̎ʐ^�Ɏg����
Public ROW_OF_ALL_IMG As Integer '���ׂĂ̎ʐ^�Ɏg���s��
Public ROW_OF_HEADDER As Integer '�w�b�_�[�Ɏg���s��
Public COL_OF_LEFT As Integer    '�����̗]���̗�
Dim imgNo As Integer             '�\��t����ʐ^�͉����ڂ��̃J�E���g


'***************************************************************************************************
'* �@�\�����F�ʐ^�\��t�������i�N���X�P�ʁj
'* ���ӎ����F�Ȃ�
'***************************************************************************************************
Sub pstImgStudents(students As Collection)
    
    '������
    pstImgStudentsInit
    
    Dim student As student
    imgNo = 0
    
    '�擾�������k�����J��Ԃ�
    For Each student In students
        imgNo = imgNo + 1
        pstImgStudent student
    Next

End Sub

'***************************************************************************************************
'* �@�\�����F����������
'* ���ӎ����F�Ȃ�
'***************************************************************************************************
Sub pstImgStudentsInit()
    
    ROW_CNT = imageRowCnt
    COL_CNT = imageColCnt
    ROW_OF_IMG = 3
    COL_OF_IMG = 5
    ROW_OF_ALL_IMG = ROW_CNT * ROW_OF_IMG
    ROW_OF_HEADDER = 2
    COL_OF_LEFT = 1
    
End Sub

'***************************************************************************************************
'* �@�\�����F����������
'* ���ӎ����F�Ȃ�
'***************************************************************************************************
Sub pstImgStudent(student As student)
    
    Dim imgPath As String
    Dim imgExistFlg As Boolean
    Dim pstCellAddress As String
    
    
    '���k�̎ʐ^�̃p�X���쐬
    imgPath = getImgPath(student)
    
    '�ʐ^�̓\��t���ʒu�擾
    pstCellAddress = getPastImgCellAddress
    
    '�ʐ^��\��t��
    pstImg pstCellAddress, imgPath

End Sub

'***************************************************************************************************
'* �@�\�����F�ʐ^�̃t���p�X���擾
'* ���ӎ����F�Ȃ�
'***************************************************************************************************
Function getImgPath(student As student)

    '���k�̎ʐ^�̃p�X:���[�g�f�B���N�g��\�w�N\�g\�c�����O
    Dim fileName As String
    Dim imgPath As String
    Dim extension As Variant
    
    '�ʐ^�̊g���q("|"��؂�ŕ�����)�����[�v
    For Each extension In imageExtensionArr()
    
        '�t�@�C����
        fileName = student.���O & extension
        
        '�p�X
        imgPath = fso.BuildPath(imageDir, student.�w�N)
        imgPath = fso.BuildPath(imgPath, student.�g)
        imgPath = fso.BuildPath(imgPath, fileName)
        
        '�ʐ^�����݂�����I��
        If (chkExistimg(imgPath)) Then
            getImgPath = imgPath
            Exit Function
        End If
    
    Next extension
    
    '�ʐ^�����݂��Ȃ��������
    getImgPath = ""
    
End Function

'***************************************************************************************************
'* �@�\�����F�ʐ^�\��t���ʒu�擾
'* ���ӎ����F�Ȃ�
'***************************************************************************************************
Function getPastImgCellAddress()
    
    Dim colIdx As Integer
    Dim colIdxTmp As Integer
    Dim rowIdx As Integer
    Dim rowIdxTmp As Integer
    Dim tmp As Integer
    '**************
    '�s�ԍ�
    '**************
    
    '���s�ڂ̎ʐ^��
    '�\��t����ʐ^�͉����ڂ��̃J�E���g / �ʐ^�𒣂�t����� �̐؂�グ
    rowIdxTmp = Application.WorksheetFunction.RoundUp(imgNo / COL_CNT, 0)

    '���\��t����s�܂łɎg�p�����s��(1�ʐ^�ɂ�3�s�g��)��ݒ肷��
    If (rowIdxTmp = 1) Then
        '��s�ڂ̏ꍇ
        rowIdx = 0
    Else
        rowIdx = ((rowIdxTmp - 1) * ROW_OF_IMG)
    End If
    
    '������\��ʐ^�̏�s��1�s��s������̂ő���
    rowIdx = rowIdx + 1
    
    '�w�b�_�[�𑫂�
    rowIdx = rowIdx + ROW_OF_HEADDER
    
    '���̎��̍s�ɓ\��t����
    rowIdx = rowIdx + 1
    
    
    '**************
    '��ԍ�
    '**************
    
    '����ڂ̎ʐ^��
    '�\��t����ʐ^�͉����ڂ��̃J�E���g / �ʐ^�𒣂�t����� �̂��܂�
    colIdxTmp = imgNo Mod COL_CNT
    If (colIdxTmp = 0) Then
        '0 �̏ꍇ�͈�ԍŌ�̗�
        colIdxTmp = COL_CNT
    End If

    '���\��t�����܂łɎg�p������(1�ʐ^�ɂ�5��g��)��ݒ肷��
    If (colIdxTmp = 1) Then
        '���ڂ̏ꍇ
        colIdx = 0
    Else
        colIdx = ((colIdxTmp - 1) * COL_OF_IMG)
    End If
    
    '������\��ʐ^�̍����1���񂪂���̂ő���
    colIdx = colIdx + 1
    
    '�ŏ���A��𑫂�
    colIdx = colIdx + COL_OF_LEFT
    
    '���̎��̗�ɓ\��t����
    colIdx = colIdx + 1
    
    
    '**************
    '�\��t����Z����ݒ�
    '**************
    'Cells �� Address ���擾�ł��Ȃ��̂ŁAA1 �Z������ Offset �Ŗ������擾����
    getPastImgCellAddress = Range("A1").Offset(rowIdx - 1, colIdx - 1).Address

End Function

'***************************************************************************************************
'* �@�\�����F�ʐ^���݃`�F�b�N
'* ���ӎ����F�Ȃ�
'***************************************************************************************************
Function chkExistimg(imgPath As String)

    chkExistimg = fso.FileExists(imgPath)

End Function

'***************************************************************************************************
'* �@�\�����F�ʐ^�\��t��
'* ���ӎ����F�Ȃ�
'***************************************************************************************************
Sub pstImg(cellAddress As String, imgPath As String)

    Dim objShape As Object
    
    If (imgPath = "") Then
        Exit Sub
    End If

    Set objShape = outputSheet.Shapes.AddPicture( _
                    fileName:=imgPath, _
                    LinkToFile:=False, _
                    SaveWithDocument:=True, _
                    Left:=Range(cellAddress).Left, _
                    Top:=Range(cellAddress).Top, _
                    Width:=Range(cellAddress).MergeArea.Width, _
                    Height:=Range(cellAddress).MergeArea.Height)
 
End Sub
