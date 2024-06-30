Attribute VB_Name = "Main"
Option Explicit

' �R�}���h���s����
Public Type CommandOutput
    exitCode As Long
    stdOut As New collection
    stdErr As New collection
End Type

Public Sub Main()
    Dim wb As Workbook
    Dim ws As Worksheet
    
    Set wb = ThisWorkbook
    Set ws = wb.Worksheets(1)
    
    '
    ' �ݒ�l�̎擾
    '
    ' [���|�W�g��URL]
    Dim repoUrl As String
    repoUrl = Trim(ws.Range("B1"))
    
    If repoUrl = "" Then
        MsgBox "[���|�W�g��URL]" & "����͂��Ă�������"
        Exit Sub
    End If
    
    ' ���|�W�g��URL����clone��p�X�𐶐�
    Dim idx As Long
    Dim dirName As String
    Dim cloneDirPath As String
    
    idx = InStrRev(repoUrl, "/")
    If idx = 0 Then
        MsgBox "[���|�W�g��URL]�̓��͒l���s���ł�"
        Exit Sub
    End If
    
    dirName = Replace(Mid(repoUrl, idx + 1), ".git", "")
    cloneDirPath = wb.path & "\" & dirName
    Debug.Print "cloneDirPath=" & cloneDirPath

    ' [�`�F�b�N�Ώۃt�H���_]
    Dim checkTargetDirs As collection
    Set checkTargetDirs = SplitStringToCollection(Trim(ws.Range("B2")))
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
      
    Dim checkTargetDirPaths As collection
    Set checkTargetDirPaths = New collection
    
    If checkTargetDirs.Count = 0 Then
        checkTargetDirPaths.Add cloneDirPath
    Else
        Dim path As Variant
        For Each path In checkTargetDirs
            checkTargetDirPaths.Add fso.BuildPath(cloneDirPath, CStr(path))
        Next
    End If

    '
    ' git clone or fetch
    '
    Dim cmdOutput As CommandOutput

    If fso.FolderExists(cloneDirPath) Then
        cmdOutput = FetchRepository(cloneDirPath)
        
        If cmdOutput.exitCode <> 0 Then
            MsgBox "git clone�Ɏ��s���܂����B"
            Exit Sub
        End If
    
    Else
        cmdOutput = CloneRepository(repoUrl, cloneDirPath)
        
        If cmdOutput.exitCode <> 0 Then
            MsgBox "git clone�Ɏ��s���܂����B"
            Exit Sub
        End If
    End If
    
    '
    ' git branch
    '
    cmdOutput = GetAllRemoteBranches(cloneDirPath)
    
    If cmdOutput.exitCode <> 0 Then
        MsgBox "git branch�Ɏ��s���܂����B"
        Exit Sub
    End If
    
    ' �u�����`���̈ꗗ�𐶐�
    Dim branches As collection
    Set branches = ConvLocalBranchNames(cmdOutput.stdOut)
    
    Dim resultList As collection ' As Dictionary<String,CodeCheckResult>
    Dim result As CodeCheckResult
    Dim branch As Variant
    
    Set resultList = New collection
    
    For Each branch In branches
        '
        ' git checkout
        '
        cmdOutput = CheckoutBranch(CStr(branch), cloneDirPath)
        
        If cmdOutput.exitCode <> 0 Then
            Set result = New CodeCheckResult
            result.branch = branch
            result.filePath = "-"
            result.lineNo = "-"
            result.errMsg = "git checkout�Ɏ��s���܂���"
            result.lineContents = "-"
            
            resultList.Add result
            GoTo Continue
        End If

        '
        ' git pull
        '
        cmdOutput = PullRepository(cloneDirPath)
        
        If cmdOutput.exitCode <> 0 Then
            Set result = New CodeCheckResult
            result.branch = branch
            result.filePath = "-"
            result.lineNo = "-"
            result.errMsg = "git pull�Ɏ��s���܂���"
            result.lineContents = "-"
            
            resultList.Add result
            GoTo Continue
        End If
        
        '
        ' check code
        '
        Dim path2 As Variant
        For Each path2 In checkTargetDirPaths
            If fso.FolderExists(path2) Then
                CheckCode CStr(branch), CStr(path2), resultList, fso
            Else
                Set result = New CodeCheckResult
                result.branch = branch
                result.filePath = path2
                result.lineNo = "-"
                result.errMsg = "���̃t�H���_�͑��݂��܂���B"
                result.lineContents = "-"
                
                resultList.Add result
            End If
        Next
Continue:
    Next
        
    '
    ' create result sheet
    '
    CreateResultSheet resultList
    
    Debug.Print "fin."
End Sub

Private Function ConvLocalBranchNames(stdOut As collection) As collection
    Dim branches As collection
    Dim line As Variant
    
    Set branches = New collection
    For Each line In stdOut
        line = Trim(line)
        
        ' HEAD��main�͏��O
        If InStr(line, "origin/HEAD") >= 1 Or InStr(line, "origin/main") >= 1 Then
            GoTo Continue
        End If
        
        If InStr(line, "feature/") <> 0 Then
            branches.Add Replace(line, "origin/", "")
        End If
Continue:
    Next
    
    Set ConvLocalBranchNames = branches
End Function

Private Sub CheckCode(branch As String, checkTargetDirPath As String, resultList As collection, fso As Object)
    
    Dim procId As String
    procId = Replace(branch, "-", "")
    If Len(procId) >= 7 Then
        procId = LCase(Right(procId, 7))
    End If
    
    Dim allFilePaths As collection
    Set allFilePaths = GetAllFilePaths(checkTargetDirPath, fso)
    
    Dim checkLogics As New collection
    checkLogics.Add New Check_Ex_BE_SCE
    checkLogics.Add New Check_Ex_SE_SCW
    checkLogics.Add New Check_Example
    
    Dim result As CodeCheckResult
    Dim filePath As Variant
    Dim emptyFiles As Boolean
    
    emptyFiles = True
    
    For Each filePath In allFilePaths
        
        Dim isCheckError As Boolean
        isCheckError = False
        
        ' �Ώۃt�@�C��
        Dim isTarget As Boolean
        isTarget = (InStr(CStr(filePath), procId) >= 1)
        isTarget = True
        If isTarget Then
            Debug.Print "Checking... Branch=" & branch & ", File=" & filePath
            emptyFiles = False
        
            Dim allLines As collection
            Set allLines = ReadFileLinesToCollection(ReadUTF8File(CStr(filePath)))
            
            Dim lineNo As Long
            For lineNo = 1 To allLines.Count
            
                Dim checkLogic As ICheckLogic
                
                For Each checkLogic In checkLogics
                    
                    If checkLogic.SkipCheck(branch, CStr(filePath), allLines(lineNo), allLines, lineNo) Then
                        GoTo NextLine
                    End If
                    
                    Dim errLineNo As Long
                    errLineNo = checkLogic.Check(allLines(lineNo), allLines, lineNo)
                    
                    If errLineNo <> -1 Then
                        isCheckError = True

                        Set result = New CodeCheckResult
                        result.branch = branch
                        result.filePath = filePath
                        result.lineNo = lineNo
                        result.errMsg = checkLogic.GetErrMsg()
                        
                        If lineNo = errLineNo Then
                            result.lineContents = allLines(lineNo)
                        Else
                            result.lineContents = "L" & lineNo & " " & allLines(lineNo) & vbLf & _
                                                  "L" & errLineNo & " " & allLines(errLineNo)
                        End If
                        
                        resultList.Add result
                    End If
                Next
NextLine:
            Next
            
            If Not isCheckError Then
                Set result = New CodeCheckResult
                result.branch = branch
                result.filePath = filePath
                result.lineNo = "-"
                result.errMsg = "-"
                result.lineContents = "�`�F�b�N�G���[�Ȃ�"
                resultList.Add result
            End If
        End If
    Next

    If emptyFiles Then
        Set result = New CodeCheckResult
        result.branch = branch
        result.filePath = "-"
        result.lineNo = "-"
        result.errMsg = "-"
        result.lineContents = "�`�F�b�N�Ώۃt�@�C����0��"
        resultList.Add result
    End If
End Sub

Private Sub CreateResultSheet(resultList As collection)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.Name = "result" & Format(Now, "yyyymmddhhnnss")
    
    ' �w�b�_�[��ݒ�
    Dim header As Variant
    Dim col As Long
    header = Array("Branch", "FilePath", "LineNo", "ErrorMessage", "LineContents")
    For col = LBound(header) To UBound(header)
        ws.Cells(1, col + 1).value = header(col)
        
        With ws.Cells(1, col + 1).Interior
            .Color = RGB(220, 230, 241)
        End With
        
        With ws.Cells(1, col + 1).Font
            .Bold = True
        End With
    Next
    
    ' �f�[�^���V�[�g�ɏ�������
    Dim row As Long
    Dim result As CodeCheckResult
    row = 2
    For Each result In resultList
        ws.Cells(row, 1).value = result.branch
        ws.Cells(row, 2).value = result.filePath
        ws.Cells(row, 3).value = result.lineNo
        ws.Cells(row, 4).value = result.errMsg
        ws.Cells(row, 5).value = result.lineContents
        row = row + 1
    Next
    
    ' �񕝂���������
    ws.Columns("A:E").AutoFit
    
    ' �t�B���^�[��ݒ�
    ws.Range("A1:E1").AutoFilter
    
    ' 1�s�ڂ��Œ�\��
    ws.Rows("2:2").Select
    ActiveWindow.FreezePanes = True
    
    ' �r��������
    With ws.UsedRange.Borders
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    ' �t�H���g��ύX
    With ws.UsedRange.Font
        .Name = "Meiryo" ' �t�H���g�����C���I�ɐݒ�
    End With
End Sub

