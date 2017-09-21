
Imports KDVNMyMsgBox

Public Class Form1
    Private Sub Form1_Click(sender As Object, e As EventArgs) Handles Me.Click
        Dim kdvnFont As Font = New Font(FontFamily.GenericSansSerif, 10, FontStyle.Regular)
        myMsgBox.Show("Testing", kdvnFont,
                                        "KDVN's Confirmation",
        myMsgBox.CustomButtons.Button3, "&Yes", "&NO, Just send it", "&Cancel",
                                        myMsgBox.DefaultButton.Button2,
                                        myMsgBox.Icon.Question)
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class
