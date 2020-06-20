Attribute VB_Name = "StudentsInfo"
Option Explicit

Dim studentsInfoBookPath As String
Dim studentsInfoSheetName As String
Dim studentsInfoBook As Workbook

'***************************************************************************************************
'* 機能説明：生徒情報の列情報
'* 注意事項：A 列から順に記載すること★
'***************************************************************************************************
Enum eInfoNo
    列_番号 = 1
    列_名前
    列_ふりがな
    列_性別
    列_学年
    列_組
    eInfoNo_End = 列_組 '一番最後の項目の値を入れる
End Enum

'***************************************************************************************************
'* 機能説明：生徒情報配列を作成する
'* 注意事項：なし
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
'* 機能説明：A1 セルで Ctrl + A した時の範囲の値を取得し、配列にセットする
'* 注意事項：なし
'***************************************************************************************************
Function GetDataAsArray(ws As Worksheet) As Variant
    GetDataAsArray = ws.Range("A1").CurrentRegion.value
End Function
