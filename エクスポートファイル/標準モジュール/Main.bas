Attribute VB_Name = "Main"
Option Explicit

Dim targetYear As String
Dim targetClass As String
Dim completedClassList As Collection
Dim formatSheet As Worksheet
Dim studentsInfoListBookPath As String
Public imageDir As String
Public imageExtensionArr() As String
Dim outputBookPath As String
Dim outputBook As Workbook
Public outputSheet As Worksheet
Public studentsInfoListSheetName As String
Dim studentsInfoListWorkBook As Workbook
Public imageRowCnt As Integer
Public imageColCnt As Integer
Public fso As Object


'***************************************************************************************************
'* �@�\�����F�ʐ^�ꗗ�쐬���� - ���C��
'* ���ӎ����F�Ȃ�
'***************************************************************************************************
Public Sub main()

    Application.ScreenUpdating = False '�������Ή�

    Dim studentsInfo As Collection
    Dim targetStudensInfo As Collection
    
    '��������
    mainInit
    mainSetting
    
    '���k���ꗗ�G�N�Z����ǂݎ���p�ŊJ��
    Dim studentsInfoListWorkBook As Workbook
    Set studentsInfoListWorkBook = Workbooks.Open(fileName:=studentsInfoListBookPath, ReadOnly:=True)
    
    '�S�Z���k���擾
    Set studentsInfo = GetStudentsInfo(studentsInfoListWorkBook.Worksheets(studentsInfoListSheetName))
    
    '���k���ꗗ�G�N�Z�������
    studentsInfoListWorkBook.Close
    
    '�o�͗p�̃��[�N�u�b�N���쐬
    Set outputBook = Workbooks.Add
    
    
    Do
        '�N���X���Ƃɏ������邽�߁A���[�v����
        
        '���̏����Ώۂ̑g�̐��k�����擾
        Set targetStudensInfo = GetStudentsInfoNextClass(studentsInfo)
        
        If targetStudensInfo.Count = 0 Then
            '�S�N���X����
            Exit Do
        End If
        
        '�o�͗p�̃��[�N�u�b�N�Ƀt�H�[�}�b�g���R�s�[
        formatSheet.Copy after:=outputBook.Sheets(Worksheets.Count)
        Set outputSheet = ActiveSheet
        outputSheet.Name = targetStudensInfo(1).�w�N & targetStudensInfo(1).�g
    
        '�w�b�_�[�ɃN���X��������
        outputSheet.Range("E2").MergeArea.value = targetStudensInfo(1).�w�N & "  " & targetStudensInfo(1).�g
        
        '�ʐ^�\��t��(�W�����W���[��:PasteImage)
        pstImgStudents targetStudensInfo
        
        '���O�𒣂�t��(�W�����W���[��:PasteName)
        pstNameStudents targetStudensInfo
    
    Loop
    
    '�o�͗p�̃��[�N�u�b�N��1�V�[�g�ڂ͋�̃V�[�g�Ȃ̂ō폜����
    Application.DisplayAlerts = False '�m�F���b�Z�[�W�I�t
    outputBook.Worksheets(1).Delete
    
    '���ׂẴV�[�g���z�[���|�W�V�����ɂ���
    homePosition outputBook
    
    '�o�͗p�̃��[�N�u�b�N�̑��݃`�F�b�N
    Dim saveFlg As Integer: saveFlg = 1
    If fso.FileExists(outputBookPath) Then
        saveFlg = MsgBox("�u" & outputBookPath & "�v" & "�͊��ɑ��݂��Ă��܂����u�������܂����H", vbOKCancel)
    End If
    
    '�o�͗p�̃��[�N�u�b�N��ۑ�
    If saveFlg = 1 Then
        outputBook.SaveAs outputBookPath
    End If
    
    outputBook.Close
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True '�������Ή�

End Sub

'***************************************************************************************************
'* �@�\�����F����������
'* ���ӎ����F�Ȃ�
'***************************************************************************************************
Sub mainInit()

    Set completedClassList = New Collection
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set formatSheet = ThisWorkbook.Worksheets("�ʐ^�ꗗ_�t�H�[�}�b�g")

End Sub


'***************************************************************************************************
'* �@�\�����F�ݒ�
'* ���ӎ����F�Ȃ�
'***************************************************************************************************
Sub mainSetting()

    imageDir = ThisWorkbook.Worksheets("�ʐ^�ꗗ�쐬").Range("J1").value                      '�ʐ^�i�[�t�H���_�p�X
    imageExtensionArr = Split(ThisWorkbook.Worksheets("�ʐ^�ꗗ�쐬").Range("J2").value, "|") '�ʐ^�̊g���q("|"��؂�ŕ�����)
    outputBookPath = fso.BuildPath( _
                            ThisWorkbook.Worksheets("�ʐ^�ꗗ�쐬").Range("J3").value, _
                            ThisWorkbook.Worksheets("�ʐ^�ꗗ�쐬").Range("J4").value)      '���k���ꗗ�V�[�g��
    studentsInfoListBookPath = ThisWorkbook.Worksheets("�ʐ^�ꗗ�쐬").Range("J5").value    '�o�̓t�@�C����
    studentsInfoListSheetName = ThisWorkbook.Worksheets("�ʐ^�ꗗ�쐬").Range("J6").value   '���k���ꗗ�G�N�Z���p�X
    imageRowCnt = ThisWorkbook.Worksheets("�ʐ^�ꗗ�쐬").Range("J7").value                 '�ʐ^�ꗗ�t�H�[�}�b�g�̎ʐ^�̍s��
    imageColCnt = ThisWorkbook.Worksheets("�ʐ^�ꗗ�쐬").Range("J8").value                 '�ʐ^�ꗗ�t�H�[�}�b�g�̎ʐ^�̗�

End Sub


'***************************************************************************************************
'* �@�\�����F���̃N���X�̐��k���擾
'* ���ӎ����F�Ȃ�
'***************************************************************************************************
Function GetStudentsInfoNextClass(students As Collection)
    
    Dim student As student
    Dim targetStudentsInfo As Collection: Set targetStudentsInfo = New Collection
    Dim targetYearLast As String
    Dim targetClassLast As String
    
    '�S�Z���k����A�N���X�P�ʂŒ��o����
    For Each student In students
        
        '�܂��������Ă��Ȃ��g�̏ꍇ
        If Not (isExists(completedClassList, student.�w�N & student.�g)) Then
        
            '����ȊO���A�w�N�܂��͑g���ς������I��
            If (targetYearLast <> "" And (targetYearLast <> student.�w�N Or targetClassLast <> student.�g)) Then
                completedClassList.Add targetYearLast & targetClassLast
                Set GetStudentsInfoNextClass = targetStudentsInfo
                Exit Function
            End If
            
            targetStudentsInfo.Add student.Self
            
            targetYearLast = student.�w�N
            targetClassLast = student.�g
        End If
        
    Next
    
    completedClassList.Add targetYearLast & targetClassLast
                
    Set GetStudentsInfoNextClass = targetStudentsInfo
    
End Function


'***************************************************************************************************
'* �@�\�����F���X�g�����݃`�F�b�N
'* ���ӎ����F�Ȃ�
'***************************************************************************************************
Function isExists(col As Collection, item As Variant) As Boolean

    If col Is Nothing Then
        isExists = False
        Exit Function
    End If

    Dim Var As Variant
    For Each Var In col
        If Var = item Then
            isExists = True
            Exit Function
        End If
    Next Var
    
    isExists = False
    
End Function

'***************************************************************************************************
'* �@�\�����F���ׂẴV�[�g���z�[���|�W�V�����ɂ���
'* ���ӎ����F�Ȃ�
'***************************************************************************************************
Sub homePosition(wb As Workbook)
    wb.Sheets(1).Activate
    Dim i As Worksheet
    For Each i In wb.Worksheets
        i.Activate
        Range("G7").Select
        Range("A1").Select
    Next
    wb.Sheets(1).Activate
End Sub
