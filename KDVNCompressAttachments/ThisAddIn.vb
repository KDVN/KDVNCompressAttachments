Imports System.Drawing
Imports System.IO

Imports ICSharpCode.SharpZipLib.Core
Imports ICSharpCode.SharpZipLib.Zip
Imports KDVNMyMsgBox


'Imports SevenZip

Public Class ThisAddIn
    Private Const strComproressFile As String = "Attachment.zip"
    Private Const strConfirmMessage As String = "- This email has attachment(s). Would you like to zip & secure with a random password?" &
        Chr(13) &
        "- If you choose YES, another email with the password will be sent automatically !" & Chr(13) &
        "- If you choose NO, Just send it"
    Private Const strSubject As String = "Attachment Password for [$CONTENT] "
    Private Const strBodyEmailPassword As String = "Below is the password for the attachment of the email having subject of [$CONTENT] "
    Private Const strMessagePS As String = Chr(13) & "NOTE: Please check next email for the password to open the Attachment"

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

    ' Copy from https://github.com/icsharpcode/SharpZipLib/wiki/Zip-Samples
    ' Compresses the files in the nominated folder, and creates a zip file on disk named as outPathname.
    '
    Public Sub CreateZipFile(folderName As String, outPathname As String, password As String)

        Dim fsOut As FileStream = File.Create(outPathname)
        Dim zipStream As New ZipOutputStream(fsOut)

        zipStream.SetLevel(3)       '0-9, 9 being the highest level of compression
        zipStream.Password = password   ' optional. Null is the same as not setting.

        ' This setting will strip the leading part of the folder path in the entries, to
        ' make the entries relative to the starting folder.
        ' To include the full path for each entry up to the drive root, assign folderOffset = 0.
        Dim folderOffset As Integer = folderName.Length + (If(folderName.EndsWith("\"), 0, 1))

        CompressFolder(folderName, zipStream, folderOffset)

        zipStream.IsStreamOwner = True
        ' Makes the Close also Close the underlying stream
        zipStream.Close()
    End Sub
    ' Recurses down the folder structure
    '
    Private Sub CompressFolder(path As String, zipStream As ZipOutputStream, folderOffset As Integer)

        Dim files As String() = Directory.GetFiles(path)

        For Each filename As String In files

            Dim fi As New FileInfo(filename)

            Dim entryName As String = filename.Substring(folderOffset)  ' Makes the name in zip based on the folder
            entryName = ZipEntry.CleanName(entryName)       ' Removes drive from name and fixes slash direction
            Dim newEntry As New ZipEntry(entryName)
            newEntry.DateTime = fi.LastWriteTime            ' Note the zip format stores 2 second granularity
            newEntry.IsUnicodeText = True
            ' Specifying the AESKeySize triggers AES encryption. Allowable values are 0 (off), 128 or 256.
            '   newEntry.AESKeySize = 256;

            ' To permit the zip to be unpacked by built-in extractor in WinXP and Server2003, WinZip 8, Java, and other older code,
            ' you need to do one of the following: Specify UseZip64.Off, or set the Size.
            ' If the file may be bigger than 4GB, or you do not need WinXP built-in compatibility, you do not need either,
            ' but the zip will be in Zip64 format which not all utilities can understand.
            '   zipStream.UseZip64 = UseZip64.Off;
            newEntry.Size = fi.Length

            zipStream.PutNextEntry(newEntry)

            ' Zip the file in buffered chunks
            ' the "using" will close the stream even if an exception occurs
            Dim buffer As Byte() = New Byte(4095) {}
            Using streamReader As FileStream = File.OpenRead(filename)
                StreamUtils.Copy(streamReader, zipStream, buffer)
            End Using
            zipStream.CloseEntry()
        Next
        Dim folders As String() = Directory.GetDirectories(path)
        For Each folder As String In folders
            CompressFolder(folder, zipStream, folderOffset)
        Next
    End Sub
    'Private Sub CreateZipFileTemp(ByVal sourcePath As String, ByVal destFilename As String, ByVal zipPassword As String)
    '    'Config
    '    Dim SevenZipCompressor As New SevenZipCompressor()
    '    SevenZipCompressor.CompressionLevel = CompressionLevel.Normal
    '    SevenZipCompressor.CompressionMethod = CompressionMethod.Deflate
    '    SevenZipCompressor.ArchiveFormat = OutArchiveFormat.Zip
    '    SevenZipCompressor.ZipEncryptionMethod = ZipEncryptionMethod.Aes256


    '    SevenZipCompressor.CompressDirectory(sourcePath, destFilename, zipPassword)
    'End Sub

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
            If Item.Attachments.Count > 0 Then

                Dim kdvnFont As Font = New Font(FontFamily.GenericSansSerif, 10, FontStyle.Regular)

                'Dim result As DialogResult
                'Cai nay
                Dim waitTime As Integer = 5 'Second
                Dim result = myMsgBox.Show(strConfirmMessage, kdvnFont,
                                        "KDVN SEND CONFIRMATION",
                                        myMsgBox.CustomButtons.Button3, "&Yes", "&NO, Just send it", "&Cancel",
                                        myMsgBox.DefaultButton.Button2,
                                        myMsgBox.Icon.Question, waitTime)
                If result = 3 Then
                    Cancel = True
                    Exit Sub
                ElseIf result = 1 Then
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
            Exit Sub
        End Try
    End Sub
End Class
