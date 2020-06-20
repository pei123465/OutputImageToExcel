Attribute VB_Name = "PasteImage"
Option Explicit

Public ROW_CNT As Integer        '写真を張り付ける行数
Public COL_CNT As Integer        '写真を張り付ける列数
Public ROW_OF_IMG As Integer     '1枚の写真に使う行数
Public COL_OF_IMG As Integer     '1枚の写真に使う列数
Public ROW_OF_ALL_IMG As Integer 'すべての写真に使う行数
Public ROW_OF_HEADDER As Integer 'ヘッダーに使う行数
Public COL_OF_LEFT As Integer    '左側の余白の列数
Dim imgNo As Integer             '貼り付ける写真は何枚目かのカウント


'***************************************************************************************************
'* 機能説明：写真貼り付け処理（クラス単位）
'* 注意事項：なし
'***************************************************************************************************
Sub pstImgStudents(students As Collection)
    
    '初期化
    pstImgStudentsInit
    
    Dim student As student
    imgNo = 0
    
    '取得した生徒数分繰り返し
    For Each student In students
        imgNo = imgNo + 1
        pstImgStudent student
    Next

End Sub

'***************************************************************************************************
'* 機能説明：初期化処理
'* 注意事項：なし
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
'* 機能説明：初期化処理
'* 注意事項：なし
'***************************************************************************************************
Sub pstImgStudent(student As student)
    
    Dim imgPath As String
    Dim imgExistFlg As Boolean
    Dim pstCellAddress As String
    
    
    '生徒の写真のパスを作成
    imgPath = getImgPath(student)
    
    '写真の貼り付け位置取得
    pstCellAddress = getPastImgCellAddress
    
    '写真を貼り付け
    pstImg pstCellAddress, imgPath

End Sub

'***************************************************************************************************
'* 機能説明：写真のフルパスを取得
'* 注意事項：なし
'***************************************************************************************************
Function getImgPath(student As student)

    '生徒の写真のパス:ルートディレクトリ\学年\組\苗字名前
    Dim fileName As String
    Dim imgPath As String
    Dim extension As Variant
    
    '写真の拡張子("|"区切りで複数可)分ループ
    For Each extension In imageExtensionArr()
    
        'ファイル名
        fileName = student.名前 & extension
        
        'パス
        imgPath = fso.BuildPath(imageDir, student.学年)
        imgPath = fso.BuildPath(imgPath, student.組)
        imgPath = fso.BuildPath(imgPath, fileName)
        
        '写真が存在したら終了
        If (chkExistimg(imgPath)) Then
            getImgPath = imgPath
            Exit Function
        End If
    
    Next extension
    
    '写真が存在しなかったら空欄
    getImgPath = ""
    
End Function

'***************************************************************************************************
'* 機能説明：写真貼り付け位置取得
'* 注意事項：なし
'***************************************************************************************************
Function getPastImgCellAddress()
    
    Dim colIdx As Integer
    Dim colIdxTmp As Integer
    Dim rowIdx As Integer
    Dim rowIdxTmp As Integer
    Dim tmp As Integer
    '**************
    '行番号
    '**************
    
    '何行目の写真か
    '貼り付ける写真は何枚目かのカウント / 写真を張り付ける列数 の切り上げ
    rowIdxTmp = Application.WorksheetFunction.RoundUp(imgNo / COL_CNT, 0)

    '今貼り付ける行までに使用した行数(1写真につき3行使う)を設定する
    If (rowIdxTmp = 1) Then
        '一行目の場合
        rowIdx = 0
    Else
        rowIdx = ((rowIdxTmp - 1) * ROW_OF_IMG)
    End If
    
    '今から貼る写真の上行に1行空行があるので足す
    rowIdx = rowIdx + 1
    
    'ヘッダーを足す
    rowIdx = rowIdx + ROW_OF_HEADDER
    
    'その次の行に貼り付ける
    rowIdx = rowIdx + 1
    
    
    '**************
    '列番号
    '**************
    
    '何列目の写真か
    '貼り付ける写真は何枚目かのカウント / 写真を張り付ける列数 のあまり
    colIdxTmp = imgNo Mod COL_CNT
    If (colIdxTmp = 0) Then
        '0 の場合は一番最後の列
        colIdxTmp = COL_CNT
    End If

    '今貼り付ける列までに使用した列数(1写真につき5列使う)を設定する
    If (colIdxTmp = 1) Then
        '一列目の場合
        colIdx = 0
    Else
        colIdx = ((colIdxTmp - 1) * COL_OF_IMG)
    End If
    
    '今から貼る写真の左列に1列空列があるので足す
    colIdx = colIdx + 1
    
    '最初のA列を足す
    colIdx = colIdx + COL_OF_LEFT
    
    'その次の列に貼り付ける
    colIdx = colIdx + 1
    
    
    '**************
    '貼り付けるセルを設定
    '**************
    'Cells で Address が取得できないので、A1 セルから Offset で無理やり取得する
    getPastImgCellAddress = Range("A1").Offset(rowIdx - 1, colIdx - 1).Address

End Function

'***************************************************************************************************
'* 機能説明：写真存在チェック
'* 注意事項：なし
'***************************************************************************************************
Function chkExistimg(imgPath As String)

    chkExistimg = fso.FileExists(imgPath)

End Function

'***************************************************************************************************
'* 機能説明：写真貼り付け
'* 注意事項：なし
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
