Imports System.Net.Mail
Public Class Form1
    Public Function GenerateMessage()
        Dim BodyMessage As String = ""
        Dim sourceString As String = New Net.WebClient().DownloadString("email.html")
        BodyMessage = sourceString.Replace("dynamicusername", TextBox1.Text)
        BodyMessage = BodyMessage.Replace("dynamicpassword", TextBox2.Text)
        Return BodyMessage
    End Function

    Public Sub SendEmail(ByVal NetCredentialEmail As String, ByVal NetCredentialPassword As String, ByVal SMTPPort As String, ByVal SMTPSSL As Boolean, ByVal SMTPHost As String, ByVal EmailFrom As String, ByVal EmailTo As String, ByVal EmailSubject As String)
        For i = 0 To 9
            Try
                Dim SmtpServer As New SmtpClient()
                Dim Mail As New MailMessage()
                With SmtpServer
                    .Credentials = New Net.NetworkCredential(NetCredentialEmail, NetCredentialPassword)
                    .Port = SMTPPort
                    .EnableSsl = SMTPSSL
                    .Host = SMTPHost
                    Mail = New MailMessage()
                    With Mail
                        .From = New MailAddress(EmailFrom)
                        .To.Add(EmailTo)
                        .Subject = EmailSubject
                        .IsBodyHtml = True
                        .Body = GenerateMessage()
                    End With
                    .Send(Mail)
                End With
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
        Next

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        SendEmail("e2e@csucarig.edu.ph", "uhyijvzwgxtrjxfq", 587, True, "smtp.gmail.com", "e2e@csucarig.edu.ph", "vincemgrm000@gmail.com", "No-Reply")
    End Sub
End Class