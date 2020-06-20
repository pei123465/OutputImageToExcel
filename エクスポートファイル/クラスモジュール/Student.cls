VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Student"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'***************************************************************************************************
'* 機能説明：生徒情報
'***************************************************************************************************

Private Parameter(1 To eInfoNo.eInfoNo_End)


'***************************************************************************************************
'* Property - Start
'* 注意事項：列を追加した場合、Property を追加すること★
'***************************************************************************************************
Property Get 番号() As String
    番号 = Parameter(eInfoNo.列_番号)
End Property
Property Get 名前() As String
    名前 = Parameter(eInfoNo.列_名前)
End Property
Property Get ふりがな() As String
    ふりがな = Parameter(eInfoNo.列_ふりがな)
End Property
Property Get 性別() As String
    性別 = Parameter(eInfoNo.列_性別)
End Property
Property Get 学年() As String
    学年 = Parameter(eInfoNo.列_学年)
End Property
Property Get 組() As String
    組 = Parameter(eInfoNo.列_組)
End Property
'***************************************************************************************************
'* Property - End
'***************************************************************************************************


Property Get Self() As Object
    Set Self = Me
End Property

Sub LetParameter(key, value)
    Parameter(key) = value
End Sub

Function GetParameter(key) As Variant
    GetParameter = Parameter(key)
End Function
