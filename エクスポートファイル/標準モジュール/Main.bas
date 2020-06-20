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
'* 機能説明：写真一覧作成処理 - メイン
'* 注意事項：なし
'***************************************************************************************************
Public Sub main()

    Application.ScreenUpdating = False '高速化対応

    Dim studentsInfo As Collection
    Dim targetStudensInfo As Collection
    
    '初期処理
    mainInit
    mainSetting
    
    '生徒情報一覧エクセルを読み取り専用で開く
    Dim studentsInfoListWorkBook As Workbook
    Set studentsInfoListWorkBook = Workbooks.Open(fileName:=studentsInfoListBookPath, ReadOnly:=True)
    
    '全校生徒情報取得
    Set studentsInfo = GetStudentsInfo(studentsInfoListWorkBook.Worksheets(studentsInfoListSheetName))
    
    '生徒情報一覧エクセルを閉じる
    studentsInfoListWorkBook.Close
    
    '出力用のワークブックを作成
    Set outputBook = Workbooks.Add
    
    
    Do
        'クラスごとに処理するため、ループする
        
        '次の処理対象の組の生徒情報を取得
        Set targetStudensInfo = GetStudentsInfoNextClass(studentsInfo)
        
        If targetStudensInfo.Count = 0 Then
            '全クラス完了
            Exit Do
        End If
        
        '出力用のワークブックにフォーマットをコピー
        formatSheet.Copy after:=outputBook.Sheets(Worksheets.Count)
        Set outputSheet = ActiveSheet
        outputSheet.Name = targetStudensInfo(1).学年 & targetStudensInfo(1).組
    
        'ヘッダーにクラス名を入れる
        outputSheet.Range("E2").MergeArea.value = targetStudensInfo(1).学年 & "  " & targetStudensInfo(1).組
        
        '写真貼り付け(標準モジュール:PasteImage)
        pstImgStudents targetStudensInfo
        
        '名前を張り付け(標準モジュール:PasteName)
        pstNameStudents targetStudensInfo
    
    Loop
    
    '出力用のワークブックの1シート目は空のシートなので削除する
    Application.DisplayAlerts = False '確認メッセージオフ
    outputBook.Worksheets(1).Delete
    
    'すべてのシートをホームポジションにする
    homePosition outputBook
    
    '出力用のワークブックの存在チェック
    Dim saveFlg As Integer: saveFlg = 1
    If fso.FileExists(outputBookPath) Then
        saveFlg = MsgBox("「" & outputBookPath & "」" & "は既に存在していますが置き換えますか？", vbOKCancel)
    End If
    
    '出力用のワークブックを保存
    If saveFlg = 1 Then
        outputBook.SaveAs outputBookPath
    End If
    
    outputBook.Close
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True '高速化対応

End Sub

'***************************************************************************************************
'* 機能説明：初期化処理
'* 注意事項：なし
'***************************************************************************************************
Sub mainInit()

    Set completedClassList = New Collection
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set formatSheet = ThisWorkbook.Worksheets("写真一覧_フォーマット")

End Sub


'***************************************************************************************************
'* 機能説明：設定
'* 注意事項：なし
'***************************************************************************************************
Sub mainSetting()

    imageDir = ThisWorkbook.Worksheets("写真一覧作成").Range("J1").value                      '写真格納フォルダパス
    imageExtensionArr = Split(ThisWorkbook.Worksheets("写真一覧作成").Range("J2").value, "|") '写真の拡張子("|"区切りで複数可)
    outputBookPath = fso.BuildPath( _
                            ThisWorkbook.Worksheets("写真一覧作成").Range("J3").value, _
                            ThisWorkbook.Worksheets("写真一覧作成").Range("J4").value)      '生徒情報一覧シート名
    studentsInfoListBookPath = ThisWorkbook.Worksheets("写真一覧作成").Range("J5").value    '出力ファイル名
    studentsInfoListSheetName = ThisWorkbook.Worksheets("写真一覧作成").Range("J6").value   '生徒情報一覧エクセルパス
    imageRowCnt = ThisWorkbook.Worksheets("写真一覧作成").Range("J7").value                 '写真一覧フォーマットの写真の行数
    imageColCnt = ThisWorkbook.Worksheets("写真一覧作成").Range("J8").value                 '写真一覧フォーマットの写真の列数

End Sub


'***************************************************************************************************
'* 機能説明：次のクラスの生徒情報取得
'* 注意事項：なし
'***************************************************************************************************
Function GetStudentsInfoNextClass(students As Collection)
    
    Dim student As student
    Dim targetStudentsInfo As Collection: Set targetStudentsInfo = New Collection
    Dim targetYearLast As String
    Dim targetClassLast As String
    
    '全校生徒から、クラス単位で抽出する
    For Each student In students
        
        'まだ処理していない組の場合
        If Not (isExists(completedClassList, student.学年 & student.組)) Then
        
            '初回以外かつ、学年または組が変わったら終了
            If (targetYearLast <> "" And (targetYearLast <> student.学年 Or targetClassLast <> student.組)) Then
                completedClassList.Add targetYearLast & targetClassLast
                Set GetStudentsInfoNextClass = targetStudentsInfo
                Exit Function
            End If
            
            targetStudentsInfo.Add student.Self
            
            targetYearLast = student.学年
            targetClassLast = student.組
        End If
        
    Next
    
    completedClassList.Add targetYearLast & targetClassLast
                
    Set GetStudentsInfoNextClass = targetStudentsInfo
    
End Function


'***************************************************************************************************
'* 機能説明：リスト内存在チェック
'* 注意事項：なし
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
'* 機能説明：すべてのシートをホームポジションにする
'* 注意事項：なし
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
