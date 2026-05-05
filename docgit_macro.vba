Sub DocGitCommit()
    Dim doc As Document
    Set doc = ActiveDocument
    
    ' Check if the document has been saved to disk
    If doc.Path = "" Then
        MsgBox "Please save your document first before committing to docgit.", vbExclamation, "docgit"
        Exit Sub
    End If
    
    ' Save the document automatically to ensure docgit grabs the latest changes
    doc.Save
    
    ' Ask the user for a commit message
    Dim message As String
    message = InputBox("Enter a commit message for docgit:", "docgit Commit")
    
    ' Cancel if no message was entered
    If message = "" Then
        MsgBox "Commit cancelled. A message is required.", vbInformation, "docgit"
        Exit Sub
    End If
    
    ' Prepare the shell command
    ' We use Chr(34) to wrap paths and messages in double quotes
    Dim shellCommand As String
    shellCommand = "cmd.exe /c docgit commit " & Chr(34) & doc.FullName & Chr(34) & " -m " & Chr(34) & message & Chr(34)
    
    ' Run the command in the background
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    
    ' 0 = Hide window, True = Wait for command to finish
    Dim returnCode As Integer
    returnCode = wsh.Run(shellCommand, 0, True)
    
    ' Check if the command succeeded
    If returnCode = 0 Then
        MsgBox "Document successfully committed!", vbInformation, "docgit"
    Else
        MsgBox "Commit failed. Are you sure you have run 'docgit init' in this folder?", vbCritical, "docgit Error"
    End If
End Sub

Sub DocGitStatus()
    Dim doc As Document
    Set doc = ActiveDocument
    
    If doc.Path = "" Then
        MsgBox "Please save your document first.", vbExclamation, "docgit"
        Exit Sub
    End If
    
    ' Open a visible command prompt to show the status
    Dim shellCommand As String
    shellCommand = "cmd.exe /k cd /d " & Chr(34) & doc.Path & Chr(34) & " && docgit status"
    
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    wsh.Run shellCommand, 1, False
End Sub
