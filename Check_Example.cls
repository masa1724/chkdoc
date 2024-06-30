VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Check_Example"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements ICheckLogic

Private re_t As Object
Private re_c1 As Object
Private beforeBranch As String
Private beforeFilePath As String

Public Sub Class_Initialize()
    Set re_t = CreateRegExp(".*cont.*")
    Set re_c1 = CreateRegExp(".*cont.*")
    beforeBranch = ""
    beforeFilePath = ""
End Sub

Private Function ICheckLogic_Check(line As String, allLines As collection, lineNo As Long) As Long
    ICheckLogic_Check = -1

    If Not re_t.test(line) Then
        Exit Function
    End If
    
    Dim offsets As New collection
    offsets.Add -1
    offsets.Add 0
    offsets.Add 1

    ICheckLogic_Check = CheckRange(allLines, re_c1, lineNo, offsets)
End Function

Private Function ICheckLogic_SkipCheck(branch As String, filePath As String, line As String, allLines As collection, lineNo As Long) As Boolean
    ' 同じブランチの同じファイルに対しては1回のみチェック
    If beforeBranch = branch And beforeFilePath = filePath Then
        ICheckLogic_SkipCheck = True
    Else
        beforeBranch = branch
        beforeFilePath = filePath
        ICheckLogic_SkipCheck = False
    End If
End Function

Private Function ICheckLogic_GetErrMsg() As String
    ICheckLogic_GetErrMsg = "has cont"
End Function
