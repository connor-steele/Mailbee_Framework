
'Module General_Usage
'    Dim smtp_configuration As New SMTP_settings
'    Dim pop_configuration As New POP_settings
'    Dim error_msg As New err_msgs
'    Dim parsed_variables As New parsed_mail
'    Dim unparsed_variables As New unparsed_mail
'    Dim error_msgs As New err_msgs


'    ' SMTP TESTS

'    '*************************************************SEND AN EMAIL FROM A @UNISERVETEAM.COM***********************************************
'    Public Function send_a_test_gmail_email_to_smtp1()
'        '*************************************************Debug Options***********************************************
'        If debug_menu = True Then
'            Console.WriteLine("To: " & SMTP_UniserveTeam_Username1 & Environment.NewLine &
'         "From: " & SMTP_GMail_Username1 & Environment.NewLine &
'           "From Password: " & SMTP_GMail_Password1 & Environment.NewLine &
'            "Body: " & "This is a Test Body" & Environment.NewLine &
'           "Subject: " & "FROM A @GMAIL.COM" & Environment.NewLine &
'           "Host: " & SMTP_GMail_Host1 & Environment.NewLine &
'           "Port: " & SMTP_Gmail_Port1 & Environment.NewLine &
'     smtp_configuration.Send_to_Support.ToString())
'        End If
'        '*************************************************SENDING USAGE***********************************************
'        Try
'            'MsgBox(SMTP_GMail_Host1)
'            send_an_email(SMTP_UniserveTeam_Username1, SMTP_GMail_Username1, SMTP_GMail_Password1, SMTP_GMail_Host1, SMTP_Gmail_Port1, sGUID, "FROM A @GMAIL.COM")
'            'send_an_email(SMTP_UniserveTeam_Username1, SMTP_UniserveTeam_Username1, SMTP_UniserveTeam_Password1, SMTP_UniserveTeam_Host1, SMTP_UniserveTeam_Port1, "This is a Test Body", "TEST SUBJECT")
'        Catch ex As Exception
'            error_msg.errmsg += ex.Message
'        End Try
'        check_errmsg_object(error_msg.errmsg)
'    End Function

'End Module
