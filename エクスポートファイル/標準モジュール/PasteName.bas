Attribute VB_Name = "PasteName"
Option Explicit

Dim nameNo As Integer         '貼り付ける名前は何枚目かのカウント
Dim ROW_OF_NAME As Integer    '1人の名前に使う行数

'***************************************************************************************************
'* 機能説明：名前貼り付け処理（クラス単位）
'* 注意事項：なし
'***************************************************************************************************
Sub pstNameStudents(students As Collection)
    
    '初期化
    pstStudentsNameInit
    
    Dim student As student
    nameNo = 0
    
    '取得した生徒数分繰り返し
    For Each student In students
        nameNo = nameNo + 1
        pstNameStudent student
    Next
    

End Sub

Sub pstStudentsNameInit()
    
    ROW_OF_NAME = 2
    
End Sub


'***************************************************************************************************
'* 機能説明：名前貼り付け処理（生徒単位）
'* 注意事項：なし
'***************************************************************************************************
Sub pstNameStudent(student As student)

    Dim pstCellAddress As String
    
    '名前の貼り付け位置取得
    pstCellAddress = getPastNameCellAddress
    
    '名前を貼り付け
    pstName pstCellAddress, student
    

End Sub

'***************************************************************************************************
'* 機能説明：名前位置取得
'* 注意事項：なし
'***************************************************************************************************
Function getPastNameCellAddress()
    
    Dim colIdx As Integer
    Dim colIdxTmp As Integer
    Dim rowIdx As Integer
    Dim rowIdxTmp As Integer
    Dim tmp As Integer
    
    '**************
    '行番号
    '**************
    
    '何行目の名前か
    '貼り付ける名前は何個目かのカウント / 名前を張り付ける列数 の切り上げ
    rowIdxTmp = Application.WorksheetFunction.RoundUp(nameNo / COL_CNT, 0)

    '今貼り付ける行までに使用した行数(1名前につき2行使う)を設定する
    If (rowIdxTmp = 1) Then
        '一行目の場合
        rowIdx = 0
    Else
        rowIdx = ((rowIdxTmp - 1) * ROW_OF_NAME)
    End If
    
    'ヘッダーを足す
    rowIdx = rowIdx + ROW_OF_HEADDER
    
    'すべての写真に使う行数を足す
    rowIdx = rowIdx + ROW_OF_ALL_IMG
    
    '写真枠と名前枠の間の1行を足す
    rowIdx = rowIdx + 1
    
    'その次の行に貼り付ける
    rowIdx = rowIdx + 1
    
    
    '**************
    '列番号
    '**************
    
    '何列目の名前か
    '貼り付ける名前は何枚目かのカウント / 名前を張り付ける列数 のあまり
    colIdxTmp = nameNo Mod COL_CNT
    If (colIdxTmp = 0) Then
        '0 の場合は一番最後の列
        colIdxTmp = COL_CNT
    End If

    '今貼り付ける列までに使用した列数(1名前につき5列使う)を設定する
    If (colIdxTmp = 1) Then
        '一列目の場合
        colIdx = 0
    Else
        colIdx = ((colIdxTmp - 1) * COL_OF_IMG)
    End If
    
    '今から貼る名前の左列に1列空列があるので足す
    colIdx = colIdx + 1
    
    '最初のA列を足す
    colIdx = colIdx + COL_OF_LEFT
    
    'その次の列に貼り付ける
    colIdx = colIdx + 1
    
    
    '**************
    '貼り付けるセルを設定
    '**************
    'Cells で Address が取得できないので、A1 セルから Offset で無理やり取得する
    getPastNameCellAddress = Range("A1").Offset(rowIdx - 1, colIdx - 1).Address

End Function

'***************************************************************************************************
'* 機能説明：名前貼り付け
'* 注意事項：なし
'***************************************************************************************************
Sub pstName(cellAddress As String, student As student)

    Dim cellOfNo As String
    Dim cellOfPhonetic As String
    Dim cellOfName As String
    Dim cellOfSex As String
    
    
    cellOfNo = cellAddress                                  'cellAddressが番号セルになる
    cellOfPhonetic = Range(cellOfNo).Offset(0, 1).Address   '番号セルの右
    cellOfName = Range(cellOfPhonetic).Offset(1, 0).Address 'ふりがなセルの下
    cellOfSex = Range(cellOfPhonetic).Offset(0, 1).Address  'ふりがなセルの右
    
    
    outputSheet.Range(cellOfNo).MergeArea.value = student.番号
    outputSheet.Range(cellOfPhonetic).value = student.ふりがな
    outputSheet.Range(cellOfName).value = student.名前
    outputSheet.Range(cellOfSex).MergeArea.value = student.性別
    
End Sub
