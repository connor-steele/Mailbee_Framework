Imports System.Data.SqlClient
Imports System.IO
Imports System.Net
Imports MailBee
Imports MailBee.Mime
Imports MailBee.Pop3Mail
Imports MailBee.SmtpMail

Module Email_Functions
    'Setup variables and test variables
    Public sConnectionString As String = "Persist Security Info=True;User ID=*****;Password=*****; Initial Catalog=databasename; Data Source=sqlservername"
    Public Test_Subject As String = "TEST"
    Public Test_Body As String = "this email is a test."
    Public Subject As String = "TEST SERVER"
    Public Date_Now As DateTime = DateTime.UtcNow


    'Setup Objects
    Public CMD As SqlCommand
    Dim SMTP_Configuration As New SMTP_Settings
    Dim POP_Configuration As New POP_Settings
    Dim Error_Msg As New Err_Msgs
    Dim Parsed_Variables As New Parsed_Mail
    Dim Unparsed_Variables As New Unparsed_Mail

    'Generates a GUID for uniquely identifying each email
    Public Sub GenerateGUID()
        Dim sGUID As String
        sGUID = System.Guid.NewGuid.ToString()
        If Debug_Menu = True Then
            Console.WriteLine("GUID: " + System.Guid.NewGuid().ToString())
        End If
        GUID = ""
        GUID = sGUID
    End Sub


    Public Function Query_Unresolved_Emails(Email_Address As String, Password As String, POP_Host As String, POP_Port As String, POP As MailBee.Pop3Mail.Pop3) As Query_Unresolved_Emails_Results
        Dim Query_Results As New Query_Unresolved_Emails_Results
        Dim SMTP_Configuration As New SMTP_Settings
        Dim POP_Configuration As New POP_Settings
        Dim SQLCON As SqlConnection

        Dim SQL_Query As String = "SELECT uid, mail_to FROM [database].[updateemail_log]  WHERE sent = 1 and RECIEVED = 0"

        If Debug_Menu = True Then
            Console.WriteLine("test_subject: " & sGUID)
        End If

        SQLCON = New SqlConnection(sConnectionString)
        Dim SQLCommand1 As New SqlCommand(SQL_Query, SQLCON)
        If SQLCON.State = ConnectionState.Closed Then
            Try
                SQLCON.Open()
            Catch ex As Exception
                If Debug_Menu = True Then
                    Console.WriteLine("Invalid DB SqlConnnection" + vbCrLf + Err.Description, "DB Connection Test")
                End If
            End Try
        End If

        Try
            Using (SQLCON)
                Dim Count As Integer = 0
                Dim SQLReader4 As SqlDataReader = SQLCommand1.ExecuteReader()
                If SQLReader4.HasRows Then


                    If Debug_Menu = True Then
                        Console.WriteLine(SQL_Query)
                    End If
                    'MsgBox("Successfull DB Connnection to stored procedure:     track_req_status_query")
                    If SQLReader4.Read() Then
                        Do While SQLReader4.Read()
                            Count += 1
                            'add this info to the empty reqnumber string
                            Console.WriteLine((SQLReader4.GetString(1).Trim.ToArray()))
                            Console.WriteLine(SQLReader4.GetString(0).Trim.ToArray())
                            Query_Results.UID = (SQLReader4.GetString(0).Trim.ToArray())
                            sGUID = Query_Results.UID
                            Check_Email_Box_for_GUID(Email_Address, Password, POP_Host, POP_Port, POP)

                        Loop
                        If Debug_Menu = True Then
                            Console.WriteLine()
                            Console.ForegroundColor = ConsoleColor.Green
                            Console.WriteLine("ALL CHECKED!")
                            Console.ResetColor()
                        End If
                    End If
                End If
                SQLReader4.Close()
            End Using
            Return Query_Results
        Catch ex As Exception
            Error_Msg.Errmsg += ex.Message
            If Debug_Menu = True Then
                Console.WriteLine(ex.Message)
            End If
        End Try
    End Function
    Public Function Check_If_Email_Exists(POP_Username As String, POP_Password As String, POP_Server As String, POP_Port As String, POP As MailBee.Pop3Mail.Pop3)
        Dim MSG As MailMessage
        Try
            POP.Connect(POP_Server, POP_Port)
            POP.Login(POP_Username, POP_Password)
        Catch xx As MailBee.MailBeeInvalidStateException
            Error_Msg.Errmsg = xx.Message
            If Debug_Menu = True Then
                Console.WriteLine(xx.Message)
            End If
        Catch ex As Exception
            Error_Msg.Errmsg = ex.Message
            If Debug_Menu = True Then
                Console.WriteLine(ex.Message)
            End If
        End Try

        'MsgBox(PopUsername + poppassword + PopPort)
        If POP.InboxMessageCount > 0 Then
            For j = 1 To POP.InboxMessageCount
                'msg = Pop3.QuickDownloadMessage(popserver, PopUsername, poppassword, 0)
                MSG = POP.DownloadEntireMessage(j)

                UID = ""
                UID = POP.GetMessageUidFromIndex(j).ToString()
                UID = CDbl(UID)
                UID = UID.Replace("'", " ")

                If Debug_Menu = True Then
                    Console.WriteLine(j)
                End If
                Unparsed_Variables.sMAIL_TO = POP_Username
                Unparsed_Variables.sRECIEVED = ""
                Dim MSG_Body As String = MSG.BodyHtmlText
                MSG_Body = MSG_Body.Replace("'", "")
                Dim htmlbody As String = MSG.BodyHtmlText
                MSG_Body = RemoveHTMLTags(MSG_Body.ToString())
                Unparsed_Variables.sRECIEVED += MSG.DateReceived
                'parsed_variables.UID += pop.GetMessageUidFromIndex(j).ToString()
                Dim MSG_Sub As String = MSG.Subject
                If (MSG_Sub.ToString.Contains(Search_Term1)) Then
                    Good_To_Go = True

                    If Good_To_Go = True Then
                        If Debug_Menu = True Then
                            Console.WriteLine()
                            Console.ForegroundColor = ConsoleColor.DarkMagenta
                            Console.WriteLine("MESSAGE DELETED!!!!")
                            Console.ResetColor()
                        End If
                        If Delete_Email_After_Read = True Then
                            Try
                                POP.DeleteMessage(j)
                            Catch ex As Exception
                                Error_Msg.Errmsg = ex.Message

                            End Try
                        Else
                            If Debug_Menu = True Then
                                Console.WriteLine()
                                Console.ForegroundColor = ConsoleColor.DarkMagenta
                                Console.WriteLine("MESSAGE NOT DELETED")
                                Console.ResetColor()
                            End If
                        End If

                    End If
                End If
            Next

        End If

        'If data_exists = "No" Then
        '    send_an_email(ReadSetting("Send_to_Support"), ReadSetting("SmtpServer.Credentials.Username1"), ReadSetting("SmtpServer.Credentials.Password1"), "Please Check on Ctera, it has sent no emails in the last 24 hours.", "CTERA: NO EMAILS SENT IN 24 HOURS", ReadSetting("Send_to_Support"))
        'End If
        POP.Disconnect()
    End Function
    Sub Send_An_Email_Autochoose(SendTo As String, From As String, From_Password As String, MailBody As String, MailSubject As String)
        GenerateGUID()
        Dim SMTPServer As New System.Net.Mail.SmtpClient
        Dim Mail As New System.Net.Mail.MailMessage()
        Dim SMTPUsername As String = From
        Dim SMTPPassword As String = From_Password
        SMTPServer.Credentials = New Net.NetworkCredential(SMTPUsername, SMTPPassword)
        Mail = New Mail.MailMessage()

        If From.EndsWith("@gmail.com") Then
            SMTPServer.Host = SMTP_GMail_Host1
            SMTPServer.Port = SMTP_Gmail_Port1
        End If



        Unparsed_Variables.sMAIL_TO = ReadSetting("send_to")

        If IsEmail(From) = False Then
            Error_Msg.Errmsg = "Not a real email address..."
            check_errmsg_object(Error_Msg.Errmsg)
        End If
        If IsEmail(SendTo) = False Then
            Error_Msg.Errmsg = "Not a real email address..."
            check_errmsg_object(Error_Msg.Errmsg)
        End If

        Mail.From = New System.Net.Mail.MailAddress(From)
        Unparsed_Variables.sMAIL_FROM = ""
        Unparsed_Variables.sMAIL_FROM += Mail.From.ToString()


        Mail.Priority = Net.Mail.MailPriority.High
        Mail.IsBodyHtml = True
        Mail.Body = MailBody
        Unparsed_Variables.sMAIL_BODY = ""
        Unparsed_Variables.sMAIL_BODY += Mail.Body


        Mail.Subject = MailSubject
        Unparsed_Variables.sMAIL_SUBJECT = ""
        Unparsed_Variables.sMAIL_SUBJECT += Mail.Subject

        'check_errmsg_object(error_msg.errmsg)

        '    'mail.Body = Page.RegisterClientScriptBlock("key", "<script>alert('Hello World');</script>")
        Dim All_Of_The_Email As String = ("New Email Logged at: " & DateTime.Now.ToString() & Environment.NewLine &
         "TO: " & SendTo & Environment.NewLine &
         "From: " & From & Environment.NewLine &
           "From Password: " & From_Password & Environment.NewLine &
            "Body: " & MailBody & Environment.NewLine &
           "Subject: " & MailSubject & Environment.NewLine)
        EmailLogEvent(All_Of_The_Email)
        Mail.To.Add(SendTo)
        Try
            SMTPServer.EnableSsl = True
            SMTPServer.Send(Mail)
            '------------------------
        Catch e As System.Net.Mail.SmtpException
            If Debug_Menu = True Then
                Error_Msg.Errmsg = e.Message
                Console.WriteLine("Exception message: " & e.Message)
                Console.WriteLine()
            End If
        Catch ex As MailBeeException
            If Debug_Menu = True Then
                Error_Msg.Errmsg += ex.Message
                Console.WriteLine(("Exception: " & ex.ToString()))
                Console.WriteLine()
                Console.WriteLine(("Exception message: " & ex.Message))
            End If

        End Try
        'POP.SSLMode = MailBee.Security.SslStartupMode.OnConnect

    End Sub
    Sub Send_An_Email(SendTo As String, From As String, From_Password As String, SMTP_Host As String, SMTP_Port As String, MailBody As String, MailSubject As String)
        GenerateGUID()
        Dim SMTPServer As New System.Net.Mail.SmtpClient
        Dim Mail As New System.Net.Mail.MailMessage()
        SMTPServer.Host = SMTP_Host
        SMTPServer.Port = SMTP_Port
        Dim SMTPUsername As String = From
        Dim SMTPPassword As String = From_Password
        SMTPServer.Credentials = New Net.NetworkCredential(SMTPUsername, SMTPPassword)
        Mail = New Mail.MailMessage()

        Unparsed_Variables.sMAIL_TO = ReadSetting("send_to")

        If IsEmail(From) = False Then
            Error_Msg.Errmsg = "Not a real email address..."
            check_errmsg_object(Error_Msg.Errmsg)
        End If
        If IsEmail(SendTo) = False Then
            Error_Msg.Errmsg = "Not a real email address..."
            check_errmsg_object(Error_Msg.Errmsg)
        End If
        Mail.From = New System.Net.Mail.MailAddress(From)
        Unparsed_Variables.sMAIL_FROM = ""
        Unparsed_Variables.sMAIL_FROM += Mail.From.ToString()


        Mail.Priority = Net.Mail.MailPriority.High
        Mail.IsBodyHtml = True
        Mail.Body = MailBody
        Unparsed_Variables.sMAIL_BODY = ""
        Unparsed_Variables.sMAIL_BODY += Mail.Body


        Mail.Subject = MailSubject
        Unparsed_Variables.sMAIL_SUBJECT = ""
        Unparsed_Variables.sMAIL_SUBJECT += Mail.Subject

        'check_errmsg_object(error_msg.errmsg)

        '    'mail.Body = Page.RegisterClientScriptBlock("key", "<script>alert('Hello World');</script>")
        Dim All_of_the_Email As String = ("New Email Logged at: " & DateTime.Now.ToString() & Environment.NewLine &
         "TO: " & SendTo & Environment.NewLine &
         "From: " & From & Environment.NewLine &
           "From Password: " & From_Password & Environment.NewLine &
            "Body: " & MailBody & Environment.NewLine &
           "Subject: " & MailSubject & Environment.NewLine)
        EmailLogEvent(All_of_the_Email)
        Mail.To.Add(SendTo)
        Try
            SMTPServer.EnableSsl = True
            SMTPServer.Send(Mail)
            '------------------------
        Catch e As System.Net.Mail.SmtpException
            If Debug_Menu = True Then
                Error_Msg.Errmsg = e.Message
                Console.WriteLine("Exception message: " & e.Message)
                Console.WriteLine()
            End If
        Catch ex As MailBeeException
            If Debug_Menu = True Then
                Error_Msg.Errmsg += ex.Message
                Console.WriteLine(("Exception: " & ex.ToString()))
                Console.WriteLine()
                Console.WriteLine(("Exception message: " & ex.Message))
            End If

        End Try
        Dim ReadText As String = File.ReadAllText(Email_Log_Path)
        If Email_Log = True Then
            '--------------------'CHECKS FOR ANY MATCHES IN LOG, IF NONE, WRITE TO A FILE----------------------------
            If ReadText.Contains(All_of_the_Email) Then

                Console.WriteLine("Already Exists " + "Message Date: " & Date_Now)

            Else
                Using SW As StreamWriter = File.AppendText(Email_Log_Path)
                    SW.WriteLine(All_of_the_Email)
                    Console.WriteLine(All_of_the_Email)
                    Console.WriteLine(SW.ToString())
                    For iCount = 1 To 1000
                    Next iCount
                    SW.WriteLine()
                    SW.Close()
                End Using
            End If
        End If


        '    'check_errmsg_object(error_msg.errmsg)
        '    'pop.SslMode = MailBee.Security.SslStartupMode.OnConnect

        '    'Dim path As String = "E:\Temp\email_log.txt"
        '    Dim path As String = ReadSetting("log_path")
    End Sub

    Public Sub Check_Email_Box_for_GUID(PopUsername As String, poppassword As String, popserver As String, PopPort As String, pop As MailBee.Pop3Mail.Pop3)
        Dim POP_Configuration As New POP_Settings
        Dim Parsed_Variables As New Parsed_Mail
        Dim Unparsed_Variables As New Unparsed_Mail
        Dim Error_Msg As New Err_Msgs
        Dim MSG As MailMessage
        Try
            pop.Connect(popserver, PopPort)
            pop.Login(PopUsername, poppassword)
        Catch xx As MailBee.MailBeeInvalidStateException
            Error_Msg.Errmsg = xx.Message
            If Debug_Menu = True Then
                Console.WriteLine(xx.Message)
            End If

        Catch ex As Exception
            Error_Msg.Errmsg = ex.Message
            If Debug_Menu = True Then
                Console.WriteLine(ex.Message)
            End If
        End Try
        'MsgBox(PopUsername + poppassword + PopPort)
        If pop.InboxMessageCount > 0 Then
            For j = 1 To pop.InboxMessageCount
                'msg = Pop3.QuickDownloadMessage(popserver, PopUsername, poppassword, 0)
                MSG = pop.DownloadEntireMessage(j)
                If Debug_Menu = True Then
                    Console.WriteLine(MSG.Subject + "vs." + sGUID)
                End If

                If (MSG.Subject = sGUID) Then
                    If Debug_Menu = True Then
                        Console.WriteLine()
                        Console.ForegroundColor = ConsoleColor.Red
                        Console.WriteLine("UNRESOLVED EMAIL FOUND!")
                        Console.ResetColor()
                    End If
                    Update_SQL_to_Recieved()
                    If Debug_Menu = True Then
                        Console.WriteLine()
                        Console.ForegroundColor = ConsoleColor.Blue
                        Console.WriteLine("UPDATED TO RECIEVED!")
                        Console.ResetColor()
                    End If

                    If Delete_Email_After_Read = True Then
                        pop.DeleteMessage(j)
                    End If

                    If Debug_Menu = True Then

                        Console.WriteLine()
                        Console.ForegroundColor = ConsoleColor.Red
                        Console.WriteLine("MESSAGE DELETED!!!!")
                        Console.ResetColor()
                    End If


                End If

            Next

        End If
        pop.Disconnect()
    End Sub
End Module