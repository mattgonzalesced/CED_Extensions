' Re-runs the CED Content Collection folder check from a Check Again button click.
' Sets CONTENTCOLLECTIONFOUND to the path if linked, or empty if not.
' Also clears WIXUI_EXITDIALOGOPTIONALTEXT on success so the warning text hides.

Function ReCheckContent()
    Dim fso, contentPath
    contentPath = "C:\ACC\ACCDocs\CoolSys\CED Content Collection"
    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FolderExists(contentPath) Then
        Session.Property("CONTENTCOLLECTIONFOUND") = contentPath
        Session.Property("WIXUI_EXITDIALOGOPTIONALTEXT") = ""
        ' Success: no MsgBox here — the Check Again button publish chain
        ' transitions to the compact CEDExitDialogOK which has the
        ' "Setup completed" message and a real Finish button.
    Else
        Session.Property("CONTENTCOLLECTIONFOUND") = ""
        MsgBox _
            "CED Content Collection folder still NOT FOUND at:" & vbCrLf & _
            contentPath & vbCrLf & vbCrLf & _
            "Link it in Autodesk Docs, then click Check Again.", _
            vbExclamation + vbOKOnly, _
            "Folder Still Missing"
    End If

    ReCheckContent = 0
End Function
