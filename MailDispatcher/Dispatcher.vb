Imports System.Configuration
Imports System.net.Mail


Public Class Dispatcher

    Shared UserName As String = IIf(Not ConfigurationManager.AppSettings("SMTPUserName") Is Nothing, ConfigurationManager.AppSettings("SMTPUserName"), "")
    Shared Password As String = IIf(Not ConfigurationManager.AppSettings("SMTPPassword") Is Nothing, ConfigurationManager.AppSettings("SMTPPassword"), "")
    Shared SMTPID As String = IIf(Not ConfigurationManager.AppSettings("SMTPID") Is Nothing, ConfigurationManager.AppSettings("SMTPID"), "localhost")
    Shared Port As String = IIf(Not ConfigurationManager.AppSettings("SMTPPort") Is Nothing, ConfigurationManager.AppSettings("SMTPPort"), "25")

    Public Shared Sub SendEmail(ByVal sBody As String, ByVal sSubject As String, ByVal sFrom As String, ByVal sTo As String)
        'Dim Mail As New System.Net.Mail.MailMessage(sFrom, sTo, sSubject, sBody)
        'Dim SmtpMail As New System.Net.Mail.SmtpClient(SMTPID, Port)

        'SmtpMail.Credentials = New Net.NetworkCredential(UserName, Password)
        'SmtpMail.Send(Mail)

        ' Create an Outlook application.
        Dim oApp As Microsoft.Office.Interop.Outlook._Application
        oApp = New Microsoft.Office.Interop.Outlook.Application()

        ' Create a new MailItem.
        Dim oMsg As Microsoft.Office.Interop.Outlook._MailItem
        oMsg = oApp.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem)
        oMsg.Subject = sSubject
        oMsg.Body = sBody

        ' TODO: Replace with a valid e-mail address.
        oMsg.To = sTo

        Dim sBodyLen As String = sBody.Length
        ' Send
        oMsg.Send()

        ' Clean up
        oApp = Nothing
        oMsg = Nothing
    End Sub
End Class
