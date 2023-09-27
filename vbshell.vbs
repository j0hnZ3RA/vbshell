Dim objShell, objExecObject, strCommand, currentDir, output

Set objShell = CreateObject("WScript.Shell")
currentDir = objShell.CurrentDirectory

strCommand = InputBox(currentDir & ">")

While strCommand <> ""
    On Error Resume Next
    
    If LCase(Left(strCommand, 5)) = "mkdir" Then
        CreateDirectory Trim(Mid(strCommand, 6))
    ElseIf LCase(Left(strCommand, 2)) = "cd" Then
        ChangeDirectory Trim(Mid(strCommand, 3))
    ElseIf LCase(strCommand) = "dir" Then
        ListDirectory
    ElseIf LCase(Left(strCommand, 4)) = "copy" Then
        CopyFile Trim(Mid(strCommand, 5))
    Else
        ' Usando Exec para capturar a saída
        Set objExecObject = objShell.Exec(strCommand)
        
        Do While Not objExecObject.StdOut.AtEndOfStream
            output = objExecObject.StdOut.ReadAll()
            If output <> "" Then
                MsgBox output
            End If
        Loop
    End If
    
    On Error GoTo 0
    strCommand = InputBox(currentDir & ">")
Wend

Sub CreateDirectory(dirName)
    Dim objFSO
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    On Error Resume Next
    objFSO.CreateFolder(currentDir & "\" & dirName)
    If Err.Number <> 0 Then
        MsgBox "Erro ao criar diretório: " & Err.Description
    End If
    On Error GoTo 0
End Sub

Sub ChangeDirectory(dirPath)
    Dim objFSO
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    If dirPath = ".." Then
        currentDir = objFSO.GetParentFolderName(currentDir)
    ElseIf objFSO.FolderExists(dirPath) Then
        currentDir = dirPath
    Else
        MsgBox "Diretório não encontrado."
    End If
End Sub

Sub ListDirectory()
    Dim objFSO, objFolder, objFile, objSubFolder
    Dim strList
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFolder = objFSO.GetFolder(currentDir)
    
    strList = " Lista de Diretórios:" & vbCrLf & vbCrLf
    For Each objSubFolder in objFolder.SubFolders
        strList = strList & "[DIR] " & objSubFolder.Name & vbCrLf
    Next
    
    strList = strList & vbCrLf & " Lista de Arquivos:" & vbCrLf & vbCrLf
    For Each objFile in objFolder.Files
        strList = strList & objFile.Name & vbCrLf
    Next
    
    MsgBox strList
End Sub

Sub CopyFile(commandArgs)
    Dim objFSO, sourceFile, destFile
    Dim args
    
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    args = Split(commandArgs, " ")
    
    If UBound(args) >= 1 Then
        sourceFile = args(0)
        destFile = args(1)
        
        If objFSO.FileExists(sourceFile) Then
            On Error Resume Next
            objFSO.CopyFile sourceFile, destFile
            If Err.Number <> 0 Then
                MsgBox "Erro ao copiar arquivo: " & Err.Description
            End If
            On Error GoTo 0
        Else
            MsgBox "Arquivo de origem não encontrado."
        End If
    Else
        MsgBox "Sintaxe incorreta. Use: copy [source] [destination]"
    End If
End Sub
