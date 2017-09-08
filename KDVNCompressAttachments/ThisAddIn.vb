Imports System.IO
Imports SevenZip



Public Class ThisAddIn
    Public Const strComproressFile As String = "Attachment.zip"
    Public Const strSubject As String = "Password for email "

    Private Sub ThisAddIn_Startup() Handles Me.Startup

    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown

    End Sub
    'Delete a dir
    Private Sub DeleteDirectory(path As String)
        If Directory.Exists(path) Then
            'Delete all files from the Directory
            For Each filepath As String In Directory.GetFiles(path)
                File.Delete(filepath)
            Next
            'Delete all child Directories
            For Each dir As String In Directory.GetDirectories(path)
                DeleteDirectory(dir)
            Next
            'Delete a Directory
            Directory.Delete(path)
        End If
    End Sub

    Function makeTempDir() As String
        Dim newTempDir = Path.GetTempPath() & RandomString(10)
        If (System.IO.Directory.Exists(newTempDir)) Then
            DeleteDirectory(newTempDir)
        End If
        System.IO.Directory.CreateDirectory(newTempDir)
        System.IO.Directory.CreateDirectory(newTempDir & "\Old")
        System.IO.Directory.CreateDirectory(newTempDir & "\New")
        Return newTempDir & "\"
    End Function

    'Generate Ramdom string
    Function RandomString(lenChars As Integer)
        Dim s As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
        Static r As New Random
        Dim chactersInString As Integer = r.Next(lenChars, lenChars)
        Dim sb As New StringBuilder
        For i As Integer = 1 To chactersInString
            Dim idx As Integer = r.Next(0, s.Length)
            sb.Append(s.Substring(idx, 1))
        Next
        Return sb.ToString()
    End Function

    Private Sub CreateZipFile(ByVal sourcePath As String, ByVal destFilename As String, ByVal zipPassword As String)
        'Config 
        ''Call to set DLL depending on processor type''
        Try
            SevenZip.SevenZipCompressor.SetLibraryPath("C:\Program Files\7-Zip\7z.dll")
        Catch
            Try
                If Environment.Is64BitProcess Then
                    SevenZip.SevenZipCompressor.SetLibraryPath("C:\Program Files (x86)\7-Zip\7z.dll")
                Else
                    MsgBox("Please install 7zip")
                    Stop
                End If
            Catch
                MsgBox("Please install 7zip")
                Stop
            End Try
        End Try

        Dim SevenZipCompressor As New SevenZipCompressor()
        SevenZipCompressor.CompressionLevel = CompressionLevel.Normal
        SevenZipCompressor.CompressionMethod = CompressionMethod.Deflate
        SevenZipCompressor.ArchiveFormat = OutArchiveFormat.Zip
        SevenZipCompressor.ZipEncryptionMethod = ZipEncryptionMethod.Aes256


        SevenZipCompressor.CompressDirectory(sourcePath, destFilename, zipPassword)
    End Sub

    Private Sub moveAttach(ByRef Item As Outlook.MailItem, strPath As String)
        Dim sFile As String
        Dim colAtts As Outlook.Attachments
        colAtts = Item.Attachments

        If colAtts.Count > 0 Then
            Item.Categories = "Compressed"
            For Each oAtt In colAtts
                ' This code looks at the last 4 characters in a filename
                'sFileType = LCase$(Right$(oAtt.FileName, 4))
                sFile = strPath & oAtt.FileName
                oAtt.SaveAsFile(sFile)
            Next
        End If
        'Remove all attachments
        While colAtts.Count > 0
            colAtts.Remove(1)
        End While
    End Sub

    Private Sub Application_ItemSend(Item As Outlook.MailItem, ByRef Cancel As Boolean) Handles Application.ItemSend
        'Cancel = True
        'CreateZipFile("")
        On Error GoTo Fail
        If Item.Attachments.Count > 0 Then
            If MsgBox("Are you sure you want to compress this message?", vbYesNo + vbQuestion _
                , "SEND CONFIRMATION") = vbYes Then
                Dim tmpAttsDir = makeTempDir()
                Dim ZipPassword = RandomString(10)
                'Extract files to temp folder
                moveAttach(Item, tmpAttsDir & "Old\")
                CreateZipFile(tmpAttsDir & "Old", tmpAttsDir & "New\" & strComproressFile, ZipPassword)
                Dim newMail As Outlook.MailItem
                newMail = Item.Copy()
                newMail.Subject = strSubject & newMail.Subject
                newMail.Body = ZipPassword
                newMail.Send()
                Item.Attachments.Add(tmpAttsDir & "New\" & strComproressFile)
                Item.Save()
                ' MsgBox("Compress",, "KDVN Notification")
            End If
        End If
        Exit Sub
Fail:
        MsgBox("Failed")
        Cancel = True
    End Sub
End Class
