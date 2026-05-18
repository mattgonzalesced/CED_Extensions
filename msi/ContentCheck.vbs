Function CheckContent()
    Dim fso, contentPath
    contentPath = "C:\ACC\ACCDocs\CoolSys\CED Content Collection"
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(contentPath) Then
        MsgBox _
            "CED Content Collection folder NOT FOUND!" & vbCrLf & vbCrLf & _
            "Expected at:" & vbCrLf & contentPath & vbCrLf & vbCrLf & _
            "Please link the CED Content Collection to your Autodesk Docs." & vbCrLf & _
            "DO NOT IGNORE THIS MESSAGE!", _
            vbExclamation + vbOKOnly, _
            "CED pyRevit Extensions Installer"
    End If

    CheckContent = 0
End Function
