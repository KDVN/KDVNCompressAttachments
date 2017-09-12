

Public Class myMsgBox

    Const frmminWidth As Integer = 345
    Const frmminHeight As Integer = 160

    Const frmmaxWidth As Integer = 800
    Const frmmaxHeight As Integer = 600

    Const txtminWidth As Integer = 250
    Const txtminHeight As Integer = 60

    Const txtmaxWidth As Integer = 700
    Const txtmaxHeight As Integer = 500

    'changes the text of the tests according to your language and preference
    Const txt_OK As String = "Ok"
    Const txt_Cancel As String = "Cancel"
    Const txt_Abort As String = "Interrompi"
    Const txt_Retry As String = "Riprova"
    Const txt_Ignore As String = "Ignora"
    Const txt_Yes As String = "Yes"
    Const txt_No As String = "No"

#Region "Enumerator"

    ''' <summary>
    ''' Specifica le costanti che definiscono i pulsanti da visualizzare in myMsgBox
    ''' </summary>
    Public Enum Buttons
        '
        ' Riepilogo:
        '     La finestra di messaggio contiene un pulsante OK.
        OK = 0
        '
        ' Riepilogo:
        '     La finestra di messaggio contiene i pulsanti OK e Annulla.
        OKCancel = 1
        '
        ' Riepilogo:
        '     La finestra di messaggio contiene i pulsanti Interrompi, Riprova e Ignora.
        AbortRetryIgnore = 2
        '
        ' Riepilogo:
        '     La finestra di messaggio contiene i pulsanti Sì, No e Annulla.
        YesNoCancel = 3
        '
        ' Riepilogo:
        '     La finestra di messaggio contiene i pulsanti Sì e No.
        YesNo = 4
        '
        ' Riepilogo:
        '     La finestra di messaggio contiene i pulsanti Riprova e Annulla.
        RetryCancel = 5
    End Enum

    ''' <summary>
    ''' Specifica le costanti che definiscono i pulsanti personalizzabili da visualizzare in myMsgBox
    ''' </summary>
    Public Enum CustomButtons
        '
        ' Riepilogo:
        '     La finestra di messaggio contiene un pulsante con testo personalizzato.
        Button1 = 6
        '
        ' Riepilogo:
        '     La finestra di messaggio contiene due pulsanto con testo personalizzato.
        Button2 = 7
        '
        ' Riepilogo:
        '     La finestra di messaggio contiene tre pulsanti con testo personalizzato.
        Button3 = 8
    End Enum

    ''' <summary>
    ''' Specifica le costanti che definiscono il pulsante predefinito in myMsgBox
    ''' </summary>
    Public Enum DefaultButton
        ''' <summary>
        ''' Il primo pulsante sulla finestra di messaggio è il pulsante predefinito.
        ''' </summary>
        Button1 = 0
        ''' <summary>
        ''' Il secondo pulsante sulla finestra di messaggio è il pulsante predefinito.
        ''' </summary>
        Button2 = 256
        ''' <summary>
        '''  Il terzo pulsante sulla finestra di messaggio è il pulsante predefinito.
        ''' </summary>
        Button3 = 512
    End Enum


    ''' <summary>
    ''' Specifica le costanti che definiscono le icone da visualizzare in myMsgBox
    ''' </summary>
    Public Enum Icon
        ''' <summary>
        ''' La finestra di messaggio non contiene simboli.
        ''' </summary>
        None = 0

        ''' <summary>
        ''' Simbolo di spunta bianco su cerchio verde
        ''' </summary>
        Confirm = 6

        ''' <summary>
        ''' Cerchio rosso con X bianca
        ''' </summary>
        [Error] = 16

        ''' <summary>
        ''' Punto interrogativo bianco in cerchio verde
        ''' </summary>
        Question = 32

        ''' <summary>
        ''' Punto interrogativo bianco in cerchio viola
        ''' </summary>
        Question2 = 33

        ''' <summary>
        ''' Punto esclamativo bianco in cerchio giallo
        ''' </summary>
        Exclamation = 48

        ''' <summary>
        ''' Cerchio rosso con simbolo negativo bianco
        ''' </summary>
        Warning = 49

        ''' <summary>
        ''' lettera minuscola i dentro un cerchio azzurro
        ''' </summary>
        Information = 64

    End Enum

    '''' <summary>
    '''' Specifica le costanti della scelta dell'utente in myMsgBox
    '''' </summary>
    'Public Enum Result
    '    None = 0
    '    OK = 1
    '    Cancel = 2
    '    Abort = 3
    '    Retry = 4
    '    Ignore = 5
    '    Yes = 6
    '    No = 7
    '    Button1 = 1
    '    Button2 = 2
    '    Button3 = 3
    'End Enum

#End Region

    <CompilerServices.DesignerGenerated()>
    Partial Private Class CustomMessageForm
        Inherits System.Windows.Forms.Form

        ' Form shadows dispose to clean up the component list.
        <System.Diagnostics.DebuggerNonUserCode()>
        Public Shadows Sub Dispose()
            Try
                If components IsNot Nothing Then
                    components.Dispose()
                End If
            Finally
                MyBase.Dispose(True)
            End Try
        End Sub


        'Richiesto da Windows Form Designer
        Private components As System.ComponentModel.IContainer

        <System.Diagnostics.DebuggerStepThrough()>
        Private Sub InitializeComponent()
            'Form  myMsgBox

            SuspendLayout()

            StartPosition = FormStartPosition.CenterScreen
            FormBorderStyle = FormBorderStyle.FixedDialog
            TopMost = True
            ControlBox = False
            ShowInTaskbar = False
            AutoScaleDimensions = New SizeF(6.0!, 13.0!)
            AutoScaleMode = AutoScaleMode.Font
            'Me.ClientSize = New Size(frmminWidth, frmminHeight)


            Width = frmminWidth
            Height = frmminHeight
            Name = "myMsgBox"
            Text = "myMsgBox"
            BackColor = Color.White

            ResumeLayout(False)
            Dim customFont = New Font(FontFamily.GenericSansSerif, 10, FontStyle.Regular)

            Button1 = New Button()
            Button1.Font = customFont
            Button1.Name = "Button1"
            Button1.Size = New Size(95, 40)
            Button1.TabIndex = 0
            Button1.Text = txt_OK
            Button1.UseVisualStyleBackColor = True
            Button1.Anchor = AnchorStyles.Right Or AnchorStyles.Top
            Button1.DialogResult = DialogResult.OK


            Button2 = New Button()
            Button2.Name = "Button2"
            Button2.Size = New Size(95, 40)
            Button2.TabIndex = 1
            Button2.Text = txt_Yes
            Button2.UseVisualStyleBackColor = True
            Button2.Anchor = AnchorStyles.Right Or AnchorStyles.Top
            Button2.DialogResult = DialogResult.Yes
            Button2.Visible = False

            Button3 = New Button()
            Button3.Name = "Button3"
            Button3.Size = New Size(95, 40)
            Button3.TabIndex = 2
            Button3.Text = txt_No
            Button3.UseVisualStyleBackColor = True
            Button3.Anchor = AnchorStyles.Right Or AnchorStyles.Top
            Button3.DialogResult = DialogResult.No
            Button3.Visible = False


            Picture1 = New PictureBox()
            Picture1.Location = New Point(5, 5)
            Picture1.Size = New Size(48, 48)
            Picture1.TabIndex = 3
            Picture1.TabStop = False

            'textbox
            txtMsg = New TextBox()
            txtMsg.Text = ""
            txtMsg.Location = New Point(60, 10)
            txtMsg.Size = New Size(txtminWidth, txtminHeight)
            txtMsg.WordWrap = True
            txtMsg.Multiline = True
            txtMsg.BackColor = BackColor
            txtMsg.ReadOnly = True
            txtMsg.BorderStyle = BorderStyle.None
            txtMsg.BackColor = Color.White
            txtMsg.Anchor = AnchorStyles.Right Or AnchorStyles.Left Or AnchorStyles.Top Or AnchorStyles.Bottom
            txtMsg.TabStop = False
            txtMsg.HideSelection = True


            'checkbox
            checkBx = New CheckBox
            checkBx.Text = ""
            checkBx.TabIndex = 4
            checkBx.BackColor = SystemColors.Control 'Color.White
            checkBx.Dock = DockStyle.Bottom
            checkBx.Padding = New System.Windows.Forms.Padding(8, 0, 0, 0)
            checkBx.Visible = False


            bottomPanel = New Panel
            bottomPanel.Controls.Add(Button1)
            bottomPanel.Controls.Add(Button2)
            bottomPanel.Controls.Add(Button3)
            'bottomPanel.Location = New Point(0, frmminHeight - 50)
            ' bottomPanel.Anchor = AnchorStyles.Bottom Or AnchorStyles.Right Or AnchorStyles.Left
            bottomPanel.Size = New Size(frmminWidth, 50)
            bottomPanel.Dock = DockStyle.Bottom
            bottomPanel.TabIndex = 5
            bottomPanel.BackColor = SystemColors.Control


            Controls.Add(txtMsg)
            Controls.Add(checkBx) 'change the order of checkBx/bottomPanel if you want the checkbox below to Buttons
            Controls.Add(bottomPanel)
            Controls.Add(Picture1)

            ResumeLayout(False)
            PerformLayout()

        End Sub
        '****************************** G. Ravi Kumar * parts *************************************************************
        '
        '' capture more form event and Disable double click on title bar and form moving. Commenting this sub it is not considered useful
        ' thanks to G. Ravi Kumar for this sub
        '  http://www.codeproject.com/script/Membership/View.aspx?mid=2748957

        Protected Overloads Overrides Sub WndProc(ByRef m As Message)
            'Define DoubleClick...
            Const WM_NCLBUTTONDBLCLK As Integer = &HA3
            'Define LeftButtonDown event...
            Const WM_NCLBUTTONDOWN As Integer = 161
            'Define MOVE action...
            Const WM_SYSCOMMAND As Integer = 274
            'Define that the WM_NCLBUTTONDOWN is at TitleBar...
            Const HTCAPTION As Integer = 2
            'Trap MOVE action...
            Const SC_MOVE As Integer = 61456
            'Disable moving TitleBar...
            If (m.Msg = WM_SYSCOMMAND) AndAlso (m.WParam.ToInt32() = SC_MOVE) Then
                Exit Sub
            End If
            'Track whether clicked on TitleBar...
            If (m.Msg = WM_NCLBUTTONDOWN) AndAlso (m.WParam.ToInt32() = HTCAPTION) Then
                Exit Sub
            End If
            'Disable double click on TitleBar...
            If (m.Msg = WM_NCLBUTTONDBLCLK) Then
                Exit Sub
            End If
            MyBase.WndProc(m)
        End Sub
        '****************************** G. Ravi Kumar * end ************************************************************
        Friend WithEvents Button1 As Button
        Friend WithEvents Button2 As Button
        Friend WithEvents Button3 As Button

        Friend txtMsg As TextBox
        Friend lblMsg As Label

        Friend WithEvents checkBx As CheckBox

        Friend WithEvents bottomPanel As Panel
        Friend WithEvents Picture1 As PictureBox

        Friend strDefaultButton As String
        Friend btndefaultButton As Button
        Friend intWaittime As Integer


        Private Sub HandleTimerTick(ByVal sender As System.Object, ByVal e As System.EventArgs)
            If intWaittime > 0 Then
                intWaittime -= 1
                btndefaultButton.Text = strDefaultButton & " (" & Trim(Str(intWaittime)) & ")"
            Else
                btndefaultButton.PerformClick()
            End If
        End Sub

        Friend Sub MakeTimer(defaultButton As DefaultButton, waitTime As Integer)
            If waitTime <= 0 Then
                Exit Sub
            Else
                Select Case defaultButton
                    Case DefaultButton.Button1
                        btndefaultButton = Button1

                    Case DefaultButton.Button2
                        btndefaultButton = Button2

                    Case DefaultButton.Button3
                        btndefaultButton = Button3
                End Select

                Dim t As Timer = New Timer()
                intWaittime = waitTime
                strDefaultButton = btndefaultButton.Text

                btndefaultButton.Text = strDefaultButton & " (" & Trim(Str(intWaittime)) & ")"
                t.Interval = 1000

                AddHandler t.Tick, AddressOf HandleTimerTick
                t.Start()
            End If
        End Sub

        Friend Sub MakeButtons(ButtonType As Buttons, defaultButton As DefaultButton)

            Select Case ButtonType

                Case Buttons.OK
                    Button1.Location = New Point(bottomPanel.Width - 110, bottomPanel.Height - 40)

                Case Buttons.OKCancel
                    Button1.Location = New Point(bottomPanel.Width - 210, bottomPanel.Height - 40)
                    Button2.Location = New Point(bottomPanel.Width - 110, bottomPanel.Height - 40)
                    Button2.Text = txt_Cancel
                    Button2.DialogResult = DialogResult.Cancel
                    Button2.Visible = True


                Case Buttons.AbortRetryIgnore
                    Button1.Location = New Point(bottomPanel.Width - 310, bottomPanel.Height - 40)
                    Button2.Location = New Point(bottomPanel.Width - 210, bottomPanel.Height - 40)
                    Button3.Location = New Point(bottomPanel.Width - 110, bottomPanel.Height - 40)


                    Button1.Text = txt_Abort
                    Button1.DialogResult = DialogResult.Abort

                    Button2.Text = txt_Retry
                    Button2.DialogResult = DialogResult.Retry
                    Button2.Visible = True

                    Button3.Text = txt_Ignore
                    Button3.DialogResult = DialogResult.Ignore
                    Button3.Visible = True


                Case Buttons.YesNoCancel
                    Button1.Location = New Point(bottomPanel.Width - 310, bottomPanel.Height - 40)
                    Button2.Location = New Point(bottomPanel.Width - 210, bottomPanel.Height - 40)
                    Button3.Location = New Point(bottomPanel.Width - 110, bottomPanel.Height - 40)


                    Button1.Text = txt_Yes
                    Button1.DialogResult = DialogResult.Yes

                    Button2.Text = txt_No
                    Button2.DialogResult = DialogResult.No
                    Button2.Visible = True

                    Button3.Text = txt_Cancel
                    Button3.DialogResult = DialogResult.Cancel
                    Button3.Visible = True

                Case Buttons.YesNo
                    Button1.Location = New Point(bottomPanel.Width - 210, bottomPanel.Height - 40)
                    Button2.Location = New Point(bottomPanel.Width - 110, bottomPanel.Height - 40)


                    Button1.Text = txt_Yes
                    Button1.DialogResult = DialogResult.Yes

                    Button2.Text = txt_No
                    Button2.DialogResult = DialogResult.No
                    Button2.Visible = True

                Case Buttons.RetryCancel
                    Button1.Location = New Point(bottomPanel.Width - 210, bottomPanel.Height - 40)
                    Button2.Location = New Point(bottomPanel.Width - 110, bottomPanel.Height - 40)


                    Button1.Text = txt_Retry
                    Button1.DialogResult = DialogResult.Retry

                    Button2.Text = txt_Cancel
                    Button2.DialogResult = DialogResult.Cancel
                    Button2.Visible = True

            End Select



            Select Case defaultButton
                Case DefaultButton.Button1
                    Button1.Select()

                Case DefaultButton.Button2

                    Button2.TabIndex = 0
                    Button3.TabIndex = 1
                    Button1.TabIndex = 2
                    Button2.Select()

                Case DefaultButton.Button3

                    Button3.TabIndex = 0
                    Button1.TabIndex = 1
                    Button2.TabIndex = 2
                    Button3.Select()

            End Select

            '  bottomPanel.Top = 10
        End Sub

        Friend Sub MakeCustomButtons(ButtonType As CustomButtons, defaultButton As DefaultButton, Optional txtButton1 As String = "", Optional txtButton2 As String = "", Optional txtButton3 As String = "")

            Select Case ButtonType

                Case CustomButtons.Button1
                    Button1.Location = New Point(bottomPanel.Width - 110, bottomPanel.Height - 40)
                    Button1.Text = txtButton1
                    Button1.DialogResult = DialogResult.OK

                Case CustomButtons.Button2
                    Button1.Location = New Point(bottomPanel.Width - 210, bottomPanel.Height - 40)
                    Button2.Location = New Point(bottomPanel.Width - 110, bottomPanel.Height - 40)

                    Button1.Text = txtButton1
                    Button1.DialogResult = DialogResult.OK

                    Button2.Text = txtButton2
                    Button2.DialogResult = DialogResult.Cancel
                    Button2.Visible = True

                Case CustomButtons.Button3
                    Button1.Location = New Point(bottomPanel.Width - 310, bottomPanel.Height - 40)
                    Button2.Location = New Point(bottomPanel.Width - 210, bottomPanel.Height - 40)
                    Button3.Location = New Point(bottomPanel.Width - 110, bottomPanel.Height - 40)

                    Button1.Text = txtButton1
                    Button1.DialogResult = DialogResult.OK

                    Button2.Text = txtButton2
                    Button2.DialogResult = DialogResult.Cancel
                    Button2.Visible = True

                    Button3.Text = txtButton3
                    Button3.DialogResult = DialogResult.Abort
                    Button3.Visible = True
            End Select

            Select Case defaultButton
                Case DefaultButton.Button1
                    Button1.Select()

                Case DefaultButton.Button2

                    Button2.TabIndex = 0
                    Button3.TabIndex = 1
                    Button1.TabIndex = 2
                    Button2.Select()

                Case DefaultButton.Button3

                    Button3.TabIndex = 0
                    Button1.TabIndex = 1
                    Button2.TabIndex = 2
                    Button3.Select()

            End Select

            '  bottomPanel.Top = 10
        End Sub


        Friend Sub SetIcon(icon As Icon)

            Select Case icon
                Case Icon.None
                    Picture1.Visible = False

                Case Icon.Confirm
                    Picture1.BackgroundImage = My.Resources.msg_ok_48

                Case Icon.Error
                    Picture1.BackgroundImage = My.Resources.msg_cancel_48

                Case Icon.Question
                    Picture1.BackgroundImage = My.Resources.msg_question_48

                Case Icon.Question2
                    Picture1.BackgroundImage = My.Resources.msg_question2_48

                Case Icon.Exclamation
                    Picture1.BackgroundImage = My.Resources.msg_alert_48

                Case Icon.Warning
                    Picture1.BackgroundImage = My.Resources.msg_remove_48

                Case Icon.Information
                    Picture1.BackgroundImage = My.Resources.msg_info_48


                Case Else
                    Picture1.Visible = False
            End Select


        End Sub

        Friend Sub MakeForm(text As String, title As String, size As Size, font As Font, backColor As Color, foreColor As Color, Optional textCheckBox As String = "")
            Dim width As Integer
            Dim height As Integer

            If String.IsNullOrEmpty(title) Then title = " "
            Me.Text = title
            ' Me.Font = font
            Me.BackColor = backColor
            Me.ForeColor = foreColor


            txtMsg.Text = text
            txtMsg.Font = font
            txtMsg.BackColor = backColor
            txtMsg.ForeColor = foreColor

            Dim graph As Graphics = CreateGraphics()
            Dim textSize As SizeF = graph.MeasureString(text, font, txtmaxWidth)

            width = CInt(textSize.Width)
            height = CInt(textSize.Height)

            If height > txtmaxHeight Then height = txtmaxHeight : txtMsg.ScrollBars = ScrollBars.Vertical

            If height > (Me.Height - 150) Then Me.Height = height + 150
            If width > (Me.Width - 100) Then Me.Width = width + 100

            'avoid form resize
            Me.MaximumSize = New Size(Me.Width, Me.Height)
            Me.MinimumSize = New Size(Me.Width, Me.Height)

            ' checkbox control required
            If textCheckBox <> "" Then
                bottomPanel.Size = New Size(frmminWidth, 40)
                checkBx.Text = textCheckBox
                checkBx.Visible = True
            End If

            graph.Dispose()

        End Sub

    End Class


#Region "Show_Function"

    Public Shared Shadows Function Show(Text As String) As DialogResult
        Dim dlg As New CustomMessageForm()
        dlg.MakeButtons(Buttons.OK, DefaultButton.Button1)
        dlg.MakeForm(Text, " ", dlg.Size, dlg.txtMsg.Font, dlg.BackColor, dlg.ForeColor)

        Return dlg.ShowDialog
        dlg.Dispose()

    End Function

    Public Shared Shadows Function Show(Text As String, Title As String) As DialogResult
        Dim dlg As New CustomMessageForm()

        dlg.MakeButtons(Buttons.OK, DefaultButton.Button1)
        dlg.MakeForm(Text, Title, dlg.Size, dlg.txtMsg.Font, dlg.BackColor, dlg.ForeColor)

        Return dlg.ShowDialog
        dlg.Dispose()

    End Function
    Public Shared Shadows Function Show(Text As String, Title As String, Icon As Icon) As DialogResult
        Dim dlg As New CustomMessageForm()

        dlg.MakeButtons(Buttons.OK, DefaultButton.Button1)
        dlg.SetIcon(Icon)

        dlg.MakeForm(Text, Title, dlg.Size, dlg.txtMsg.Font, dlg.BackColor, dlg.ForeColor)

        Return dlg.ShowDialog
        dlg.Dispose()

    End Function

    Public Shared Shadows Function Show(Text As String, TextFont As Font, Title As String, Icon As Icon) As DialogResult
        Dim dlg As New CustomMessageForm()

        dlg.MakeButtons(Buttons.OK, DefaultButton.Button1)
        dlg.SetIcon(Icon)

        dlg.MakeForm(Text, Title, dlg.Size, TextFont, dlg.BackColor, dlg.ForeColor)

        Return dlg.ShowDialog
        dlg.Dispose()

    End Function

    Public Shared Shadows Function Show(Text As String, Title As String, Buttons As CustomButtons, txtButton1 As String) As DialogResult
        Dim dlg As New CustomMessageForm()

        dlg.MakeCustomButtons(Buttons, DefaultButton.Button1, txtButton1)
        dlg.MakeForm(Text, Title, dlg.Size, dlg.txtMsg.Font, dlg.BackColor, dlg.ForeColor)

        Return dlg.ShowDialog
        dlg.Dispose()

    End Function

    Public Shared Shadows Function Show(Text As String, Title As String, Buttons As CustomButtons, txtButton1 As String, txtButton2 As String) As DialogResult
        Dim dlg As New CustomMessageForm()

        dlg.MakeCustomButtons(Buttons, DefaultButton.Button1, txtButton1, txtButton2)
        dlg.MakeForm(Text, Title, dlg.Size, dlg.txtMsg.Font, dlg.BackColor, dlg.ForeColor)

        Return dlg.ShowDialog
        dlg.Dispose()

    End Function

    Public Shared Shadows Function Show(Text As String, Title As String, Buttons As CustomButtons, txtButton1 As String, txtButton2 As String, txtButton3 As String) As DialogResult
        Dim dlg As New CustomMessageForm()

        dlg.MakeCustomButtons(Buttons, DefaultButton.Button1, txtButton1, txtButton2, txtButton3)
        dlg.MakeForm(Text, Title, dlg.Size, dlg.txtMsg.Font, dlg.BackColor, dlg.ForeColor)

        Return dlg.ShowDialog
        dlg.Dispose()

    End Function


    Public Shared Shadows Function Show(Text As String, Title As String, Buttons As Buttons, DefaultButton As DefaultButton) As DialogResult
        Dim dlg As New CustomMessageForm()

        dlg.MakeButtons(Buttons, DefaultButton)
        dlg.MakeForm(Text, Title, dlg.Size, dlg.txtMsg.Font, dlg.BackColor, dlg.ForeColor)

        Return dlg.ShowDialog
        dlg.Dispose()

    End Function
    Public Shared Shadows Function Show(Text As String, Title As String, Buttons As CustomButtons, txtButton1 As String, DefaultButton As DefaultButton) As DialogResult
        Dim dlg As New CustomMessageForm()

        dlg.MakeCustomButtons(Buttons, DefaultButton, txtButton1)
        dlg.MakeForm(Text, Title, dlg.Size, dlg.txtMsg.Font, dlg.BackColor, dlg.ForeColor)

        Return dlg.ShowDialog
        dlg.Dispose()

    End Function

    Public Shared Shadows Function Show(Text As String, Title As String, Buttons As CustomButtons, txtButton1 As String, txtButton2 As String, DefaultButton As DefaultButton) As DialogResult
        Dim dlg As New CustomMessageForm()

        dlg.MakeCustomButtons(Buttons, DefaultButton, txtButton1, txtButton2)
        dlg.MakeForm(Text, Title, dlg.Size, dlg.txtMsg.Font, dlg.BackColor, dlg.ForeColor)

        Return dlg.ShowDialog
        dlg.Dispose()

    End Function

    Public Shared Shadows Function Show(Text As String, Title As String, Buttons As CustomButtons, txtButton1 As String, txtButton2 As String, txtButton3 As String, DefaultButton As DefaultButton) As DialogResult
        Dim dlg As New CustomMessageForm()

        dlg.MakeCustomButtons(Buttons, DefaultButton, txtButton1, txtButton2, txtButton3)
        dlg.MakeForm(Text, Title, dlg.Size, dlg.txtMsg.Font, dlg.BackColor, dlg.ForeColor)

        Return dlg.ShowDialog
        dlg.Dispose()

    End Function



    Public Shared Shadows Function Show(Text As String, Title As String, Buttons As Buttons, Icon As Icon) As DialogResult
        Dim dlg As New CustomMessageForm()

        dlg.MakeButtons(Buttons, DefaultButton.Button1)
        dlg.SetIcon(Icon)

        dlg.MakeForm(Text, Title, dlg.Size, dlg.txtMsg.Font, dlg.BackColor, dlg.ForeColor)


        Return dlg.ShowDialog
        dlg.Dispose()

    End Function

    Public Shared Shadows Function Show(Text As String, Title As String, Buttons As CustomButtons, txtButton1 As String, Icon As Icon) As DialogResult
        Dim dlg As New CustomMessageForm()

        dlg.MakeCustomButtons(Buttons, DefaultButton.Button1, txtButton1)
        dlg.SetIcon(Icon)

        dlg.MakeForm(Text, Title, dlg.Size, dlg.txtMsg.Font, dlg.BackColor, dlg.ForeColor)


        Return dlg.ShowDialog
        dlg.Dispose()

    End Function
    Public Shared Shadows Function Show(Text As String, Title As String, Buttons As CustomButtons, txtButton1 As String, txtButton2 As String, Icon As Icon) As DialogResult
        Dim dlg As New CustomMessageForm()

        dlg.MakeCustomButtons(Buttons, DefaultButton.Button1, txtButton1, txtButton2)
        dlg.SetIcon(Icon)

        dlg.MakeForm(Text, Title, dlg.Size, dlg.txtMsg.Font, dlg.BackColor, dlg.ForeColor)


        Return dlg.ShowDialog
        dlg.Dispose()

    End Function
    Public Shared Shadows Function Show(Text As String, Title As String, Buttons As CustomButtons, txtButton1 As String, txtButton2 As String, txtButton3 As String, Icon As Icon) As DialogResult
        Dim dlg As New CustomMessageForm()

        dlg.MakeCustomButtons(Buttons, DefaultButton.Button1, txtButton1, txtButton2, txtButton3)
        dlg.SetIcon(Icon)

        dlg.MakeForm(Text, Title, dlg.Size, dlg.txtMsg.Font, dlg.BackColor, dlg.ForeColor)


        Return dlg.ShowDialog
        dlg.Dispose()

    End Function



    Public Shared Shadows Function Show(Text As String, Title As String, Buttons As Buttons, DefaultButton As DefaultButton, Icon As Icon) As DialogResult
        Dim dlg As New CustomMessageForm()

        dlg.MakeButtons(Buttons, DefaultButton)
        dlg.SetIcon(Icon)

        dlg.MakeForm(Text, Title, dlg.Size, dlg.txtMsg.Font, dlg.BackColor, dlg.ForeColor)


        Return dlg.ShowDialog
        dlg.Dispose()

    End Function

    Public Shared Shadows Function Show(Text As String, TextFont As Font, Title As String, Buttons As Buttons, DefaultButton As DefaultButton, Icon As Icon) As DialogResult
        Dim dlg As New CustomMessageForm()

        dlg.MakeButtons(Buttons, DefaultButton)
        dlg.SetIcon(Icon)

        dlg.MakeForm(Text, Title, dlg.Size, TextFont, dlg.BackColor, dlg.ForeColor)


        Return dlg.ShowDialog
        dlg.Dispose()

    End Function

    Public Shared Shadows Function Show(Text As String, Title As String, Buttons As CustomButtons, txtButton1 As String, DefaultButton As DefaultButton, Icon As Icon) As DialogResult
        Dim dlg As New CustomMessageForm()
        MsgBox("HEdfasfdsafllo")
        dlg.MakeCustomButtons(Buttons, DefaultButton, txtButton1)
        dlg.SetIcon(Icon)

        dlg.MakeForm(Text, Title, dlg.Size, dlg.txtMsg.Font, dlg.BackColor, dlg.ForeColor)


        Return dlg.ShowDialog
        dlg.Dispose()

    End Function

    Public Shared Shadows Function Show(Text As String, TextFont As Font, Title As String, Buttons As CustomButtons, txtButton1 As String, DefaultButton As DefaultButton, Icon As Icon) As DialogResult
        Dim dlg As New CustomMessageForm()
        MsgBox("HEllo4")
        dlg.MakeCustomButtons(Buttons, DefaultButton, txtButton1)
        dlg.SetIcon(Icon)

        dlg.MakeForm(Text, Title, dlg.Size, TextFont, dlg.BackColor, dlg.ForeColor)


        Return dlg.ShowDialog
        dlg.Dispose()

    End Function

    Public Shared Shadows Function Show(Text As String, Title As String, Buttons As CustomButtons, txtButton1 As String, txtButton2 As String, DefaultButton As DefaultButton, Icon As Icon) As DialogResult
        Dim dlg As New CustomMessageForm()

        dlg.MakeCustomButtons(Buttons, DefaultButton, txtButton1, txtButton2)
        dlg.SetIcon(Icon)

        dlg.MakeForm(Text, Title, dlg.Size, dlg.txtMsg.Font, dlg.BackColor, dlg.ForeColor)


        Return dlg.ShowDialog
        dlg.Dispose()

    End Function

    Public Shared Shadows Function Show(Text As String, TextFont As Font, Title As String, Buttons As CustomButtons, txtButton1 As String, txtButton2 As String, DefaultButton As DefaultButton, Icon As Icon) As DialogResult
        Dim dlg As New CustomMessageForm()
        MsgBox("HEllo3")
        dlg.MakeCustomButtons(Buttons, DefaultButton, txtButton1, txtButton2)
        dlg.SetIcon(Icon)

        dlg.MakeForm(Text, Title, dlg.Size, TextFont, dlg.BackColor, dlg.ForeColor)


        Return dlg.ShowDialog
        dlg.Dispose()

    End Function

    Public Shared Shadows Function Show(Text As String, Title As String, Buttons As CustomButtons, txtButton1 As String, txtButton2 As String, txtButton3 As String, DefaultButton As DefaultButton, Icon As Icon) As DialogResult
        Dim dlg As New CustomMessageForm()
        MsgBox("HEllo2")
        dlg.MakeCustomButtons(Buttons, DefaultButton, txtButton1, txtButton2, txtButton3)
        dlg.SetIcon(Icon)

        dlg.MakeForm(Text, Title, dlg.Size, dlg.txtMsg.Font, dlg.BackColor, dlg.ForeColor)


        Return dlg.ShowDialog
        dlg.Dispose()

    End Function

    Public Shared Shadows Function Show(Text As String, TextFont As Font, Title As String, Buttons As CustomButtons, txtButton1 As String, txtButton2 As String, txtButton3 As String, DefaultButton As DefaultButton, Icon As Icon, Optional interval As Integer = 0) As DialogResult
        Dim dlg As New CustomMessageForm()

        dlg.MakeCustomButtons(Buttons, DefaultButton, txtButton1, txtButton2, txtButton3)
        dlg.SetIcon(Icon)
        dlg.MakeTimer(DefaultButton, interval)
        dlg.MakeForm(Text, Title, dlg.Size, TextFont, dlg.BackColor, dlg.ForeColor)

        Return dlg.ShowDialog
        dlg.Dispose()

    End Function


    Public Shared Shadows Function Show(Text As String, Title As String, Buttons As Buttons, Icon As Icon, textCheckBox As String, ByRef Check As Boolean) As DialogResult
        Dim dlg As New CustomMessageForm()

        dlg.MakeButtons(Buttons, DefaultButton.Button1)
        dlg.SetIcon(Icon)

        dlg.MakeForm(Text, Title, dlg.Size, dlg.txtMsg.Font, dlg.BackColor, dlg.ForeColor, textCheckBox)


        Dim dlgresult As DialogResult = dlg.ShowDialog
        Check = dlg.checkBx.Checked
        Return dlgresult
        dlg.Dispose()

    End Function



    Public Shared Shadows Function Show(Text As String, Title As String, Buttons As Buttons, DefaultButton As DefaultButton, Icon As Icon, textCheckBox As String, ByRef Check As Boolean) As DialogResult
        Dim dlg As New CustomMessageForm()

        dlg.MakeButtons(Buttons, DefaultButton)
        dlg.SetIcon(Icon)

        dlg.MakeForm(Text, Title, dlg.Size, dlg.txtMsg.Font, dlg.BackColor, dlg.ForeColor, textCheckBox)


        Dim dlgresult As DialogResult = dlg.ShowDialog
        Check = dlg.checkBx.Checked
        Return dlgresult
        dlg.Dispose()

    End Function

    Public Shared Shadows Function Show(Text As String, TextFont As Font, Title As String, Buttons As Buttons, DefaultButton As DefaultButton, Icon As Icon, textCheckBox As String, ByRef Check As Boolean) As DialogResult
        Dim dlg As New CustomMessageForm()

        dlg.MakeButtons(Buttons, DefaultButton)
        dlg.SetIcon(Icon)

        dlg.MakeForm(Text, Title, dlg.Size, TextFont, dlg.BackColor, dlg.ForeColor, textCheckBox)


        Dim dlgresult As DialogResult = dlg.ShowDialog
        Check = dlg.checkBx.Checked
        Return dlgresult
        dlg.Dispose()

    End Function

    Public Shared Shadows Function Show(Text As String, Title As String, Buttons As CustomButtons, txtButton1 As String, DefaultButton As DefaultButton, Icon As Icon, textCheckBox As String, ByRef Check As Boolean) As DialogResult
        Dim dlg As New CustomMessageForm()

        dlg.MakeCustomButtons(Buttons, DefaultButton, txtButton1)
        dlg.SetIcon(Icon)

        dlg.MakeForm(Text, Title, dlg.Size, dlg.txtMsg.Font, dlg.BackColor, dlg.ForeColor, textCheckBox)


        Dim dlgresult As DialogResult = dlg.ShowDialog
        Check = dlg.checkBx.Checked
        Return dlgresult
        dlg.Dispose()

    End Function

    Public Shared Shadows Function Show(Text As String, TextFont As Font, Title As String, Buttons As CustomButtons, txtButton1 As String, DefaultButton As DefaultButton, Icon As Icon, textCheckBox As String, ByRef Check As Boolean) As DialogResult
        Dim dlg As New CustomMessageForm()

        dlg.MakeCustomButtons(Buttons, DefaultButton, txtButton1)
        dlg.SetIcon(Icon)

        dlg.MakeForm(Text, Title, dlg.Size, TextFont, dlg.BackColor, dlg.ForeColor, textCheckBox)


        Dim dlgresult As DialogResult = dlg.ShowDialog
        Check = dlg.checkBx.Checked
        Return dlgresult
        dlg.Dispose()

    End Function



    Public Shared Shadows Function Show(Text As String, Title As String, Buttons As CustomButtons, txtButton1 As String, txtButton2 As String, DefaultButton As DefaultButton, Icon As Icon, textCheckBox As String, ByRef Check As Boolean) As DialogResult
        Dim dlg As New CustomMessageForm()

        dlg.MakeCustomButtons(Buttons, DefaultButton, txtButton1, txtButton2)
        dlg.SetIcon(Icon)

        dlg.MakeForm(Text, Title, dlg.Size, dlg.txtMsg.Font, dlg.BackColor, dlg.ForeColor, textCheckBox)


        Dim dlgresult As DialogResult = dlg.ShowDialog
        Check = dlg.checkBx.Checked
        Return dlgresult
        dlg.Dispose()

    End Function

    Public Shared Shadows Function Show(Text As String, TextFont As Font, Title As String, Buttons As CustomButtons, txtButton1 As String, txtButton2 As String, DefaultButton As DefaultButton, Icon As Icon, textCheckBox As String, ByRef Check As Boolean) As DialogResult
        Dim dlg As New CustomMessageForm()

        dlg.MakeCustomButtons(Buttons, DefaultButton, txtButton1, txtButton2)
        dlg.SetIcon(Icon)

        dlg.MakeForm(Text, Title, dlg.Size, TextFont, dlg.BackColor, dlg.ForeColor, textCheckBox)


        Dim dlgresult As DialogResult = dlg.ShowDialog
        Check = dlg.checkBx.Checked
        Return dlgresult
        dlg.Dispose()

    End Function



    Public Shared Shadows Function Show(Text As String, Title As String, Buttons As CustomButtons, txtButton1 As String, txtButton2 As String, txtButton3 As String, DefaultButton As DefaultButton, Icon As Icon, textCheckBox As String, ByRef Check As Boolean) As DialogResult
        Dim dlg As New CustomMessageForm()

        dlg.MakeCustomButtons(Buttons, DefaultButton, txtButton1, txtButton2, txtButton3)
        dlg.SetIcon(Icon)

        dlg.MakeForm(Text, Title, dlg.Size, dlg.txtMsg.Font, dlg.BackColor, dlg.ForeColor, textCheckBox)


        Dim dlgresult As DialogResult = dlg.ShowDialog
        Check = dlg.checkBx.Checked
        Return dlgresult
        dlg.Dispose()

    End Function

    Public Shared Shadows Function Show(Text As String, TextFont As Font, Title As String, Buttons As CustomButtons, txtButton1 As String, txtButton2 As String, txtButton3 As String, DefaultButton As DefaultButton, Icon As Icon, textCheckBox As String, ByRef Check As Boolean) As DialogResult
        Dim dlg As New CustomMessageForm()

        dlg.MakeCustomButtons(Buttons, DefaultButton, txtButton1, txtButton2, txtButton3)
        dlg.SetIcon(Icon)

        dlg.MakeForm(Text, Title, dlg.Size, TextFont, dlg.BackColor, dlg.ForeColor, textCheckBox)


        Dim dlgresult As DialogResult = dlg.ShowDialog
        Check = dlg.checkBx.Checked
        Return dlgresult
        dlg.Dispose()

    End Function

#End Region

End Class