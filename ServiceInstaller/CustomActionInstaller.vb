Imports System.IO
Imports System.IO.Directory

Public Class CustomActionInstaller
    Sub unsintallVSTO(ByVal vtsoFile As String)
        ' Try uninstall VSTO
        Try
            Dim vstoPath = SearchForFiles("C:\Program Files\Common Files\microsoft shared\VSTO\", "VSTOInstaller.exe")
            For Each fileVSTOInstaller In vstoPath
                Dim aroras As Process = Process.Start(fileVSTOInstaller, "/s /u " & vtsoFile)
                aroras.WaitForExit()
            Next
            vstoPath = SearchForFiles("C:\Program Files(x86)\Common Files\microsoft shared\VSTO\", "VSTOInstaller.exe")
            For Each fileVSTOInstaller In vstoPath
                Dim aroras1 As Process = Process.Start(fileVSTOInstaller, "/s /u " & vtsoFile)
                aroras1.WaitForExit()
            Next
        Catch ex As Exception

        End Try
        'Try clean AppCache
        Try
            Dim cleanApp As Process = Process.Start(Environment.GetFolderPath(Environment.SpecialFolder.System) & "\rundll32.exe", "dfshim CleanOnlineAppCache")
        Catch ex As Exception
        End Try
    End Sub
    Function SearchForFiles(ByVal RootFolder As String, ByVal FileFilter As String) As List(Of String)
        Dim ReturnedData As New List(Of String)                             'List to hold the search results
        Dim FolderStack As New Stack(Of String)                             'Stack for searching the folders
        FolderStack.Push(RootFolder)                                        'Start at the specified root folder
        Do While FolderStack.Count > 0                                      'While there are things in the stack
            Dim ThisFolder As String = FolderStack.Pop                      'Grab the next folder to process
            Try                                                             'Use a try to catch any errors
                For Each SubFolder In GetDirectories(ThisFolder)            'Loop through each sub folder in this folder
                    FolderStack.Push(SubFolder)                             'Add to the stack for further processing
                Next                                                        'Process next sub folder
                ReturnedData.AddRange(GetFiles(ThisFolder, FileFilter))    'Search for and return the matched file names
            Catch ex As Exception                                           'For simplicity sake
            End Try                                                         'We'll ignore the errors
        Loop                                                                'Process next folder in the stack
        Return ReturnedData                                                 'Return the list of files that match
    End Function

    'Custom for Commit & Install
    <Security.Permissions.SecurityPermission(Security.Permissions.SecurityAction.Demand)>
    Public Overrides Sub Install(ByVal stateSaver As System.Collections.IDictionary)
        'Try
        Dim cleanApp As Process = Process.Start(Environment.GetFolderPath(Environment.SpecialFolder.System) & "\rundll32.exe", "dfshim CleanOnlineAppCache")
            SafeDeleteDirectory(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) & "\apps\2.0")

            Dim procs As Process() = Process.GetProcessesByName("dfsvc.exe")
            For Each proc As Process In procs
                proc.Kill()
            Next
        '   Catch ex As Exception
        '  End Try
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
                If Path.GetExtension(objFile).ToUpper() = ".VSTO" Then unsintallVSTO(objFile)
                If Not objFile.Contains(System.Reflection.Assembly.GetAssembly(GetType(CustomActionInstaller)).GetName.Name) Then
                    SafeDeleteFile(objFile)
                End If

            Next
            For Each objDir In Directory.GetDirectories(PathToDelete)
                SafeDeleteDirectory(objDir)
            Next
        End If
        MyBase.Uninstall(savedState)
        'End If

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
