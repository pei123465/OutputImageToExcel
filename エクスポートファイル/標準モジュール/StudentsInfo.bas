Attribute VB_Name = "StudentsInfo"
Option Explicit

Dim studentsInfoBookPath As String
Dim studentsInfoSheetName As String
Dim studentsInfoBook As Workbook

'***************************************************************************************************
'* �@�\�����F���k���̗���
'* ���ӎ����FA �񂩂珇�ɋL�ڂ��邱�Ɓ�
'***************************************************************************************************
Enum eInfoNo
    ��_�ԍ� = 1
    ��_���O
    ��_�ӂ肪��
    ��_����
    ��_�w�N
    ��_�g
    eInfoNo_End = ��_�g '��ԍŌ�̍��ڂ̒l������
End Enum

'***************************************************************************************************
'* �@�\�����F���k���z����쐬����
'* ���ӎ����F�Ȃ�
'***************************************************************************************************
Public Function GetStudentsInfo(ws As Worksheet) As Collection
    Dim arr: arr = GetDataAsArray(ws)
    Dim C As Collection: Set C = New Collection
    Dim i, j
    For i = LBound(arr, 1) + 1 To UBound(arr, 1)
        With New student
            
            For j = 1 To eInfoNo.eInfoNo_End
                .LetParameter j, arr(i, j)
            Next
            C.Add .Self
        End With
    Next
    Set GetStudentsInfo = C
End Function

'***************************************************************************************************
'* �@�\�����FA1 �Z���� Ctrl + A �������͈̔͂̒l���擾���A�z��ɃZ�b�g����
'* ���ӎ����F�Ȃ�
'***************************************************************************************************
Function GetDataAsArray(ws As Worksheet) As Variant
    GetDataAsArray = ws.Range("A1").CurrentRegion.value
End Function
