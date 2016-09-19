Imports System.Configuration
Imports System.IO
Imports System.Text.RegularExpressions

Module Parsing_Functions
    Dim error_msg As New err_msgs
    '--------------------------------------------------------EMPTY ALL EMAIL VARIABLES-----------------------------------------------
    'These variables can be used to empty all of your parsing variables. Should be used at the end of a loop if you want to make sure
    'all of your variables are empty.
    Public Function make_parsed_variables_empty() As parsed_mail
        Dim empty_parsed_variables As New parsed_mail
        empty_parsed_variables.parsed_MAIL_TO = ""
        empty_parsed_variables.parsed_MAIL_FROM = ""
        empty_parsed_variables.parsed_MAIL_SUBJECT = ""
        empty_parsed_variables.parsed_MAIL_BODY = ""
        empty_parsed_variables.parsed_DATE_RECIEVED = ""
        empty_parsed_variables.parsed_SENT = ""
        empty_parsed_variables.parsed_RECIEVED = ""
    End Function
    Public Function make_unparsed_variables_empty() As unparsed_mail
        Dim empty_unparsed_variables As New unparsed_mail
        empty_unparsed_variables.sMAIL_TO = ""
        empty_unparsed_variables.sMAIL_FROM = ""
        empty_unparsed_variables.sMAIL_SUBJECT = ""
        empty_unparsed_variables.sMAIL_BODY = ""
        empty_unparsed_variables.sDATE_RECIEVED = ""
        empty_unparsed_variables.sSENT = ""
        empty_unparsed_variables.sRECIEVED = ""
    End Function

    Public Function check_errmsg_object(error_message As String) As err_msgs
        If debug_menu = True Then
            Console.ForegroundColor = ConsoleColor.Yellow
            Console.WriteLine("checking for errors...")
            Console.ResetColor()
        End If

        If Len(error_message) > 0 Then
            If error_log = True Then
                ErrorLogEvent(error_message)
            End If

            If debug_menu = True Then
                Console.ForegroundColor = ConsoleColor.Red
                Console.WriteLine("Exception Error: " + error_message)
                Console.ResetColor()
            End If


            Return error_msg
            'send_an_email(ReadSetting("send_to_me"), ReadSetting("SmtpServer.Credentials.Username"), ReadSetting("SmtpServer.Credentials.Password"), errmsg, "CHECK_EMAIL_SERVERS APPLICATION: ", ReadSetting("Send_to_Support"))
            'error_msg.errmsg = ""
        Else
            If debug_menu = True Then
                Console.ForegroundColor = ConsoleColor.Green
                Console.WriteLine("NO ERRORS FOUND!")
                Console.ResetColor()
            End If
        End If
        Return error_msg
    End Function
    '----------------------------------------------------------------------------------------PARSES OUT ALL ILLEGAL CHARACTERS AND CODE-----------------------------------------------
    Public Function Parse_Data(MAIL_TO As String, MAIL_FROM As String, MAIL_SUBJECT As String, MAIL_BODY As String, [UID] As String, date_recieved As String) As parsed_mail
        Dim parsed_variables As New parsed_mail
        Dim unparsed_variables As New unparsed_mail
        MAIL_SUBJECT = unparsed_variables.sMAIL_SUBJECT
        MAIL_SUBJECT = MAIL_SUBJECT.Replace("'", " ")
        MAIL_SUBJECT = Left(MAIL_SUBJECT, 100)
        'Return MAIL_SUBJECT
        parsed_variables.parsed_MAIL_SUBJECT = ""
        parsed_variables.parsed_MAIL_SUBJECT += MAIL_SUBJECT

        MAIL_FROM = unparsed_variables.sMAIL_FROM
        MAIL_FROM = MAIL_FROM.Replace("'", " ")
        'Return MAIL_FROM
        parsed_variables.parsed_MAIL_FROM = ""
        parsed_variables.parsed_MAIL_FROM += MAIL_FROM

        MAIL_BODY = unparsed_variables.sMAIL_BODY
        MAIL_BODY = MAIL_BODY.Replace("'", "")
        MAIL_BODY = RemoveHTMLTags(MAIL_BODY.ToString()).Replace("&nbsp;", "")
        'Return MAIL_BODY
        parsed_variables.parsed_MAIL_BODY = ""
        parsed_variables.parsed_MAIL_BODY += MAIL_BODY

        'parsed_variables.parsed_DATE_RECIEVED = date_now
        'Return date_recieved




        'UID = CDbl(parsed_variables.UID)
        'UID = UID.Replace("'", " ")
        'Return UID
        parsed_variables.UID = UID
    End Function

    '----------------------------------------------------------------------------------------REMOVE ALL HTML AND STYLE TAGS-----------------------------------------------
    Public Function RemoveHTMLTags(ByVal HTMLCode As String) As String
        HTMLCode = Regex.Replace(HTMLCode, "style=()[^\1]*?>", "")
        HTMLCode = Regex.Replace(HTMLCode, "<style>()[^\1]*?>", "")
        Return Regex.Replace(HTMLCode, "<[^>]*>", "")
    End Function


    'removes all text upto a certain point
    Public Function ALLTextUpTo(ByRef WordtoFind As String, ByRef Sentence As String) As String
        Dim error_msg As New err_msgs
        If (Len(WordtoFind) = 0) Or (Len(Sentence) = 0) Then
            If debug_menu = True Then
                Console.WriteLine("no text")
            End If
        Else
            Try
                ALLTextUpTo = Sentence.Substring(0, Sentence.IndexOf(WordtoFind))
            Catch ex As Exception
                error_msg.errmsg += ex.Message
            End Try

        End If

    End Function
    '----------------------------------------------------------------------------------------READ EXTERNAL FILES-----------------------------------------------
    Sub ReadAllSettings()
        Dim error_msg As New err_msgs
        Try
            Dim appSettings = ConfigurationManager.AppSettings
            Dim all_settings As String = ""
            If appSettings.Count = 0 Then
                If debug_menu = True Then
                    Console.WriteLine("AppSettings is empty.")
                End If
            Else
                For Each key As String In appSettings.AllKeys
                    If debug_menu = True Then
                        'Console.WriteLine("")
                        'Console.WriteLine("Key: {0} Value: {1}", key, appSettings(key))
                    End If
                    all_settings += ("Key: {0} Value: {1}" & key & appSettings(key))
                Next
            End If
            'MsgBox(all_settings)

        Catch e As ConfigurationErrorsException
            If debug_menu = True Then
                Console.WriteLine("Error reading app settings")
            End If
            error_msg.errmsg += e.Message
        End Try
    End Sub


    Public Function ReadSetting(key As String)
        Dim error_msg As New err_msgs
        Try
            Dim appSettings = ConfigurationManager.AppSettings
            Dim result As String = appSettings(key)
            If IsNothing(result) Then
                If debug_menu = True Then
                    'Console.WriteLine()
                    'Console.WriteLine("this setting cant be read: " & appSettings(key).ToString())
                End If

                result = "Could not read " & key & ". Check your spelling, it must match the app.config file perfectly."
            End If
            If debug_menu = True Then
                Console.WriteLine("")
                Console.WriteLine(result)
            End If
            'MsgBox(result)
            Return result
        Catch e As ConfigurationErrorsException
            error_msg.errmsg += e.Message
            If debug_menu = True Then
                Console.WriteLine("Error reading app settings")
            End If
        End Try
    End Function

    Sub AddUpdateAppSettings(key As String, value As String)
        Dim error_msg As New err_msgs
        Try
            Dim configFile = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None)
            Dim settings = configFile.AppSettings.Settings
            If IsNothing(settings(key)) Then
                settings.Add(key, value)
            Else
                settings(key).Value = value
            End If
            configFile.Save(ConfigurationSaveMode.Modified)
            If debug_menu = True Then
                ConfigurationManager.RefreshSection(configFile.AppSettings.SectionInformation.Name)
            End If
        Catch e As ConfigurationErrorsException
            error_msg.errmsg += e.Message
            If debug_menu = True Then
                Console.WriteLine("Error writing app settings")
            End If

        End Try
    End Sub
    Function IsEmail(ByVal email As String) As Boolean
        Static emailExpression As New Regex("\A(?:[a-z0-9!#$%&'*+/=?^_`{|}~-]+(?:\.[a-z0-9!#$%&'*+/=?^_`{|}~-]+)* | (?:[\x01-\x08\x0b\x0c\x0e-\x1f\x21\x23-\x5b\x5d-\x7f] | \\[\x01-\x09\x0b\x0c\x0e-\x7f])*) @ (?:(?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?\.)+[a-z0-9](?:[a-z0-9-]*[a-z0-9])? | \[(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3} (?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?|[a-z0-9-]*[a-z0-9]: (?:[\x01-\x08\x0b\x0c\x0e-\x1f\x21-\x5a\x53-\x7f] | \\[\x01-\x09\x0b\x0c\x0e-\x7f])+) \])\z")

        Return emailExpression.IsMatch(email)
    End Function

    Public Sub OnEntryWritten(ByVal source As Object, ByVal e As EntryWrittenEventArgs)
        Console.WriteLine(("Written: " + e.Entry.Message))
    End Sub
End Module
