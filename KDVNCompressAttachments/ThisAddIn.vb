Imports System.IO
Imports SevenZip

Public Class ThisAddIn
    Public Const strComproressFile As String = "Attachment.zip",
        strConfirmMessage As String = "- This email has attachment(s). Would you like to zip & secure with a random password?" &
        Chr(13) &
        "- If you choose YES, anther email with the password will be sent automatically !" & Chr(13) &
        "- If you choose NO, Just send it"

    Public Const strSubject As String = "Attachment Password for ! $CONTENT ! ",
        strBodyEmailPassword As String = "Below is the password for the attachment of the email having subject of ! $CONTENT !"
    Public Const strMessagePS As String = Chr(13) & "NOTE: Please check next email for the password to open the Attachment"


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

    Private Sub KDVNWarning(strMessage As String, ByRef Cancel As Boolean)
        Cancel = True
        MsgBox(strMessage,, "KDVN Warning")
    End Sub

    Private Sub Application_ItemSend(Item As Outlook.MailItem, ByRef Cancel As Boolean) Handles Application.ItemSend
        'Cancel = True
        'Check 7Zip
        ''Call to set DLL depending on processor type''
        Try
            SevenZip.SevenZipCompressor.SetLibraryPath("C:\Program Files\7-Zip\7z.dll")
        Catch
            Try
                If Environment.Is64BitOperatingSystem Then
                    SevenZip.SevenZipCompressor.SetLibraryPath("C:\Program Files (x86)\7-Zip\7z.dll")
                Else
                    KDVNWarning("Please Install 7zip", Cancel)
                    Exit Sub
                End If
            Catch ex As Exception
                KDVNWarning(ex.Message & Chr(13) & "Please Install 7zip", Cancel)
                Exit Sub
            End Try
        End Try
        Try
            If Item.Attachments.Count > 0 Then
                If MsgBox(strConfirmMessage, vbYesNo + vbQuestion _
                    , "KDVN SEND CONFIRMATION") = vbYes Then
                    Dim tmpAttsDir = makeTempDir()
                    Dim ZipPassword = RandomString(10)
                    'Extract files to temp folder
                    moveAttach(Item, tmpAttsDir & "Old\")
                    CreateZipFile(tmpAttsDir & "Old", tmpAttsDir & "New\" & strComproressFile, ZipPassword)
                    Dim newMail As Outlook.MailItem
                    newMail = Item.Copy()
                    Dim originalSubject = Item.Subject

                    Dim newstrBodyEmailPassword = strBodyEmailPassword.Replace("$CONTENT", originalSubject)
                    Dim newSubject = strSubject.Replace("$CONTENT", originalSubject)
                    newMail.Subject = newSubject
                    newMail.Body = newstrBodyEmailPassword & Chr(13) & ZipPassword

                    Item.Body = Item.Body & strMessagePS
                    Item.Attachments.Add(tmpAttsDir & "New\" & strComproressFile)
                    Item.Save()

                    newMail.Send()

                    DeleteDirectory(tmpAttsDir)
                    ' MsgBox("Compress",, "KDVN Notification")
                End If
            End If
        Catch ex As Exception
            KDVNWarning(ex.Message, Cancel)
        End Try
    End Sub
End Class
