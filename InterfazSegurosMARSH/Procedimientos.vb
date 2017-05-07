Imports System.Net.Mail

Public Class Codigos

    Shared Sub PrintLog(ByVal Linea As String, ByVal NombreArchivo As String)
        Try

            System.IO.File.AppendAllText(System.Windows.Forms.Application.StartupPath & "\" & Year(Now) & Month(Now) & Day(Now) & NombreArchivo, Now & ": " & Linea & Chr(13) & Chr(10))

            System.Console.WriteLine(Now & ": " & Linea)
        Catch ex As Exception

        End Try

    End Sub

    Shared Function PrintLogCorreo(ByVal Procedimiento As String, ByVal Linea As String) As String
        Try

            ' System.IO.File.AppendAllText(System.Windows.Forms.Application.StartupPath & "\" & Year(Now) & Month(Now) & Day(Now) & NombreArchivo, Now & ": " & Linea & Chr(13) & Chr(10))

            Return (Now & ": " & Procedimiento & ": " & Linea & "<br>" & Chr(10) & Chr(13))
        Catch ex As Exception

        End Try

    End Function

    Shared Sub SendMail(ByVal MSG As String, ByVal Subject As String)
        Dim TablaConfiguracion As SegurosMarshTableAdapters.SegurosMARSH_ConfiguracionTableAdapter
        Dim Correos As SegurosMarsh.SegurosMARSH_ConfiguracionDataTable
        Dim Correo As SegurosMarsh.SegurosMARSH_ConfiguracionRow
        Dim ServidorSMTP As String



        TablaConfiguracion = New SegurosMarshTableAdapters.SegurosMARSH_ConfiguracionTableAdapter()
        Correos = TablaConfiguracion.GetCorreos("General")
        ServidorSMTP = TablaConfiguracion.GetValor("CONFIGURACION", "SMTP")



        Dim oMail As MailMessage = New MailMessage()
        Try
            oMail.From = New System.Net.Mail.MailAddress("ELLIPSE.Santiago@glencore.cl")
            For Each Correo In Correos

                oMail.To.Add(New System.Net.Mail.MailAddress(Correo.Valor2, Correo.Valor))
            Next

            oMail.Priority = MailPriority.Normal
            oMail.IsBodyHtml = True
            oMail.Subject = Subject
            oMail.Body = MSG

            Dim oSmtp As SmtpClient = New System.Net.Mail.SmtpClient
            oSmtp.DeliveryMethod = SmtpDeliveryMethod.Network
            oSmtp.Host = ServidorSMTP
            ' oSmtp.Credentials = New System.Net.NetworkCredential("ellipse", "")
            oSmtp.EnableSsl = False
            oSmtp.Port = 25
            oSmtp.Send(oMail)

        Catch ex As Exception

        End Try
    End Sub

    Shared Sub SendMailAdmin(ByVal MSG As String, ByVal Subject As String, ByVal Distrito As String)
        Dim TablaConfiguracion As SegurosMarshTableAdapters.SegurosMARSH_ConfiguracionTableAdapter
        Dim Correos As SegurosMarsh.SegurosMARSH_ConfiguracionDataTable
        Dim Correo As SegurosMarsh.SegurosMARSH_ConfiguracionRow
        Dim ServidorSMTP As String



        TablaConfiguracion = New SegurosMarshTableAdapters.SegurosMARSH_ConfiguracionTableAdapter()
        Correos = TablaConfiguracion.GetCorreos("Admin" & Distrito)
        ServidorSMTP = TablaConfiguracion.GetValor("CONFIGURACION", "SMTP")



        Dim oMail As MailMessage = New MailMessage()
        Try
            oMail.From = New System.Net.Mail.MailAddress("ELLIPSE.Santiago@glencore.cl")
            For Each Correo In Correos

                oMail.To.Add(New System.Net.Mail.MailAddress(Correo.Valor2, Correo.Valor))
            Next

            oMail.Priority = MailPriority.Normal
            oMail.IsBodyHtml = True
            oMail.Subject = Subject
            oMail.Body = MSG

            Dim oSmtp As SmtpClient = New System.Net.Mail.SmtpClient
            oSmtp.DeliveryMethod = SmtpDeliveryMethod.Network
            oSmtp.Host = ServidorSMTP
            ' oSmtp.Credentials = New System.Net.NetworkCredential("ellipse", "")
            oSmtp.EnableSsl = False
            oSmtp.Port = 25
            oSmtp.Send(oMail)

        Catch ex As Exception

        End Try
    End Sub

    Shared Sub SendMail(ByVal Emails As ArrayList, ByVal MSG As String, ByVal Subject As String)
        Dim TablaConfiguracion As SegurosMarshTableAdapters.SegurosMARSH_ConfiguracionTableAdapter
        Dim Correos As ArrayList
        Dim Correo As String
        Dim ServidorSMTP As String



        TablaConfiguracion = New SegurosMarshTableAdapters.SegurosMARSH_ConfiguracionTableAdapter()
        Correos = Emails
        ServidorSMTP = TablaConfiguracion.GetValor("CONFIGURACION", "SMTP")



        Dim oMail As MailMessage = New MailMessage()
        Try
            oMail.From = New System.Net.Mail.MailAddress("ELLIPSE.Santiago@glencore.cl")
            For Each Correo In Correos

                oMail.To.Add(New System.Net.Mail.MailAddress(Correo))
            Next

            oMail.Priority = MailPriority.Normal
            oMail.IsBodyHtml = True
            oMail.Subject = Subject
            oMail.Body = MSG

            Dim oSmtp As SmtpClient = New System.Net.Mail.SmtpClient
            oSmtp.DeliveryMethod = SmtpDeliveryMethod.Network
            oSmtp.Host = ServidorSMTP
            ' oSmtp.Credentials = New System.Net.NetworkCredential("ellipse", "")
            oSmtp.EnableSsl = False
            oSmtp.Port = 25
            oSmtp.Send(oMail)

        Catch ex As Exception

        End Try
    End Sub
End Class