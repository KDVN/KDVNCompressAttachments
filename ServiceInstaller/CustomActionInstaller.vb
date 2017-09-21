Imports System.IO

Public Class CustomActionInstaller
    'Custom for Commit & Install
    <Security.Permissions.SecurityPermission(Security.Permissions.SecurityAction.Demand)>
    Public Overrides Sub Install(ByVal stateSaver As System.Collections.IDictionary)
        MyBase.Install(stateSaver)
    End Sub

    <Security.Permissions.SecurityPermission(Security.Permissions.SecurityAction.Demand)>
    Public Overrides Sub Commit(ByVal savedState As System.Collections.IDictionary)
        MyBase.Commit(savedState)
    End Sub


    'Custom for Uninstall
    <Security.Permissions.SecurityPermission(Security.Permissions.SecurityAction.Demand)>
    Public Overrides Sub Uninstall(ByVal savedState As System.Collections.IDictionary)
        Dim PathToDelete As String = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly.Location)
        'If MsgBox("Do you want to remove all files and folder " & PathToDelete, MsgBoxStyle.YesNo, "KDVN Information") = MsgBoxResult.Yes Then
        If ((Not (PathToDelete) Is Nothing) _
            AndAlso Directory.Exists(PathToDelete)) Then
                ' delete all the file inside this folder except SID.SetupSupport
                For Each objFile In Directory.GetFiles(PathToDelete)
                    ' MessageBox.Show(file);
                    If Not objFile.Contains(System.Reflection.Assembly.GetAssembly(GetType(CustomActionInstaller)).GetName.Name) Then
                        SafeDeleteFile(objFile)
                    End If

                Next
                For Each objDir In Directory.GetDirectories(PathToDelete)
                    SafeDeleteDirectory(objDir)
                Next
            End If
        'End If
        MyBase.Uninstall(savedState)
    End Sub

    Private Shared Sub SafeDeleteFile(ByVal strfile As String)
        Try
            File.Delete(strfile)
        Catch

        End Try

    End Sub

    Private Shared Sub SafeDeleteDirectory(ByVal strDirectory As String)
        Try
            Directory.Delete(strDirectory, True)
        Catch

        End Try

    End Sub

End Class
