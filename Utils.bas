Attribute VB_Name = "Utils"
Option Explicit

Public Function ExecuteCommand(command As String, args As collection, currentDir As String) As CommandOutput
    Dim wsh As Object
    Dim exec As Object
    Dim output As CommandOutput
    Dim commandLine As String
    Dim arg As Variant
    
    Set wsh = CreateObject("WScript.Shell")
    Set output.stdOut = New collection
    Set output.stdErr = New collection
    
    If Trim(currentDir) <> "" Then
        command = "cd /d """ & currentDir & """ && " & command
    End If

    commandLine = command
    
    For Each arg In args
        If arg = "2>NUL" Or arg = "||" Or arg = "&&" Then
            commandLine = commandLine & " " & arg
        Else
            commandLine = commandLine & " """ & arg & """"
        End If
    Next
    
    ' コマンドを実行
    Set exec = wsh.exec("cmd /c " & commandLine)
    
    ' 標準出力を読み取り
    Do While Not exec.stdOut.AtEndOfStream
        output.stdOut.Add exec.stdOut.ReadLine()
    Loop
    
    ' 標準エラーを読み取り
    Do While Not exec.stdErr.AtEndOfStream
        output.stdErr.Add exec.stdErr.ReadLine()
    Loop
    
    ' 結果コードを取得
    output.exitCode = exec.exitCode
    
    DebugPrintCommandOutput commandLine, output

    ExecuteCommand = output
End Function

Public Function CloneRepository(repoUrl As String, clonePath) As CommandOutput
    Dim cmdArgs As collection
    Dim cmdOutput As CommandOutput

    Set cmdArgs = New collection
    cmdArgs.Add "clone"
    cmdArgs.Add repoUrl
    cmdArgs.Add clonePath
    cmdOutput = ExecuteCommand("git", cmdArgs, "")

    CloneRepository = cmdOutput
End Function

Public Function FetchRepository(currentDir As String) As CommandOutput
    Dim cmdArgs As collection
    Dim cmdOutput As CommandOutput

    Set cmdArgs = New collection
    cmdArgs.Add "fetch"
    cmdOutput = ExecuteCommand("git", cmdArgs, currentDir)
    
    FetchRepository = cmdOutput
End Function

Public Function GetAllRemoteBranches(currentDir As String) As CommandOutput
    Dim cmdArgs As collection
    Dim cmdOutput As CommandOutput

    Set cmdArgs = New collection
    cmdArgs.Add "branch"
    cmdArgs.Add "-r"
    cmdOutput = ExecuteCommand("git", cmdArgs, currentDir)
    
    GetAllRemoteBranches = cmdOutput
End Function

Public Function CheckoutBranch(branchName As String, currentDir As String) As CommandOutput
    Dim cmdArgs As collection
    Dim cmdOutput As CommandOutput

    Set cmdArgs = New collection
    cmdArgs.Add "checkout"
    cmdArgs.Add branchName
    cmdArgs.Add "2>NUL"
    cmdArgs.Add "||"
    cmdArgs.Add "git"
    cmdArgs.Add "checkout"
    cmdArgs.Add "-b"
    cmdArgs.Add branchName
    cmdArgs.Add "origin/" & branchName
    cmdOutput = ExecuteCommand("git", cmdArgs, currentDir)
    CheckoutBranch = cmdOutput

End Function

Public Function PullRepository(currentDir As String) As CommandOutput
    Dim cmdArgs As collection
    Dim cmdOutput As CommandOutput

    Set cmdArgs = New collection
    cmdArgs.Add "pull"
    cmdOutput = ExecuteCommand("git", cmdArgs, currentDir)
    
    PullRepository = cmdOutput
End Function

Public Function ReadFileLinesToCollection(fileContents As String) As collection
    
    Dim tempLines As Variant
    fileContents = Replace(fileContents, vbCrLf, vbLf)
    tempLines = Split(fileContents, vbLf)
    
    Dim lines As collection
    Set lines = New collection
    
    Dim line As Variant
    For Each line In tempLines
        lines.Add line
    Next line
    
    Set ReadFileLinesToCollection = lines
End Function

Public Function GetAllFilePaths(ByRef dirPath As String, ByRef fso As Object) As collection
    Dim folder As Object
    Dim file As Object
    Dim subFolder As Object
    Dim filePaths As collection
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set filePaths = New collection

    ' 指定フォルダ内のすべてのファイルを取得
    Set folder = fso.GetFolder(dirPath)
    For Each file In folder.Files
        If file.Name <> ".git" Then
            filePaths.Add file.path
        End If
    Next

    ' 指定フォルダ内のすべてのサブフォルダを再帰的に処理
    For Each subFolder In folder.SubFolders
        If subFolder.Name <> ".git" Then
            Dim subFilePaths As collection
            Set subFilePaths = GetAllFilePaths(subFolder.path, fso)
            
            Dim subFilePath As Variant
            For Each subFilePath In subFilePaths
                filePaths.Add subFilePath
            Next subFilePath
        End If
    Next

    Set GetAllFilePaths = filePaths
    Exit Function
End Function

Public Function ReadUTF8File(filePath As String) As String
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")

    With stream
        .Type = 1 ' adTypeBinary
        .Open
        .LoadFromFile filePath
        ' UTF-8 BOMをスキップするために先頭3バイトを確認
        If .size >= 3 Then
            .Position = 0
            Dim bom As Variant
            bom = .Read(3)
            If bom(0) = &HEF And bom(1) = &HBB And bom(2) = &HBF Then
                ' UTF-8 BOMが存在する場合
                .Position = 3
            Else
                ' UTF-8 BOMが存在しない場合
                .Position = 0
            End If
        End If
        .Type = 2 ' adTypeText
        .Charset = "utf-8"
    End With

    Dim content As String
    content = stream.ReadText(-1) ' -1: adReadAll

    stream.Close
    Set stream = Nothing

    ReadUTF8File = content
End Function

Private Sub DebugPrintCommandOutput(commandLine As String, cmdOutout As CommandOutput)
    Dim output As String
    Dim line As Variant
    
    output = "[Command]" & vbCrLf & "  " & commandLine & vbCrLf
    output = output & "[Exit code]" & vbCrLf & "  " & cmdOutout.exitCode & vbCrLf
    
    output = output & "[StdOut]" & vbCrLf
    For Each line In cmdOutout.stdOut
        output = output & "  " & line & vbCrLf
    Next line
    
    output = output & "[StdErr]" & vbCrLf
    For Each line In cmdOutout.stdErr
        output = output & "  " & line & vbCrLf
    Next line
    
    Debug.Print output
End Sub

Public Function SplitStringToCollection(str As String) As collection
    Dim collection As collection
    Set collection = New collection
    
    If str = "" Then
        Set SplitStringToCollection = collection
        Exit Function
    End If

    Dim items() As String
    Dim st As Long
    Dim ed As Long
    
    items = Split(Replace(Replace(str, ",", vbLf), "、", vbLf), vbLf)
    st = LBound(items)
    ed = UBound(items)
    
    Dim i As Long
    
    For i = st To ed
        Dim s As String
        s = Trim(items(i))
        
        If s <> "" Then
            collection.Add s
        End If
    Next
    
    Set SplitStringToCollection = collection
End Function
