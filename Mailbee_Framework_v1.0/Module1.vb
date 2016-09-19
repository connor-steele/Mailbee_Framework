Imports System.Configuration
Module Module1

    Sub Main()
        MailBee.Global.LicenseKey = licsence_key
        Dim smtp_configuration As New SMTP_settings
        Dim pop_configuration As New POP_settings
        Dim error_msg As New err_msgs
        Dim parsed_variables As New parsed_mail
        Dim unparsed_variables As New Unparsed_Mail
        Dim pop As New MailBee.Pop3Mail.Pop3


    End Sub

End Module
