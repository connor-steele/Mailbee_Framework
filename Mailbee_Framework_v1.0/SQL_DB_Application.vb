Imports System.Data.SqlClient
Module SQL_DB_Application


    '/////////General Table Usage//////////////////////

    '    create table email_log(
    'ID bigint Not NULL IDENTITY(-2000000000000,1),
    'MAIL_TO varchar(255) Not NULL,
    'MAIL_FROM varchar(255),
    'MAIL_SUBJECT varchar(255),
    'MAIL_BODY varchar(max),
    'DATE_RECIEVED varchar(255),
    'SENT varchar(255),
    'RECIEVED varchar(255),
    'UID varchar(255),
    'PRIMARY KEY (ID));

    Public Function Update_SQL_to_Recieved()

        Dim Error_Msg As New err_msgs
        Dim Parsed_Variables As New parsed_mail
        Dim Unparsed_Variables As New unparsed_mail
        Dim SSSQL1 As String = "[database].[update email_log] set [recieved]=1 " & "where [MAIL_SUBJECT]=" & "'" & Parsed_Variables.parsed_MAIL_SUBJECT & "'" & "AND [UID]=" & "'" & UID & "';"
        If Debug_Menu = True Then
            Console.WriteLine("test_subject: " & sGUID)
        End If

        Try
            Dim cmd1 As SqlCommand
            cmd1 = New SqlCommand(SSSQL1)
            If Debug_Menu = True Then
                Console.WriteLine(SSSQL1)
            End If

            Return InsertUpdateData(cmd1)

        Catch ex As Exception
            Error_Msg.errmsg += ex.Message
        End Try
    End Function

    Public Function Insert_into_SQL_sent(MAIL_TO As String, MAIL_FROM As String, MAIL_SUBJECT As String, MAIL_BODY As String, UID As String, date_recieved As String)
        Dim error_msg As New err_msgs
        Dim parsed_variables As New parsed_mail
        Dim unparsed_variables As New unparsed_mail

        Dim checked As Boolean = False

        Try
            'MsgBox(date_recieved.ToString())
            'MsgBox(UID.ToString())
            'DATE_RECIEVED = "',GETDATE(),'"
            Dim sssql As String = "INSERT INTO [database].[update email_log]"
            sssql = sssql & "([MAIL_TO],[MAIL_FROM],[MAIL_SUBJECT],[MAIL_BODY],[DATE_RECIEVED],[SENT],[RECIEVED],[UID])"
            sssql = sssql & "VALUES"
            sssql = sssql & "('" & MAIL_TO & "','" & MAIL_FROM & "','" & MAIL_SUBJECT & "','" & MAIL_BODY & "',getdate(),'" & "1" & "',0,'" & sGUID & "')"
            'make_unparsed_variables_empty()
            'MsgBox(unparsed_variables.sMAIL_TO & "CHECK")'
            unparsed_variables.sMAIL_SUBJECT = sGUID
            unparsed_variables.sMAIL_TO = MAIL_TO
            unparsed_variables.sMAIL_FROM = MAIL_FROM
            unparsed_variables.sMAIL_BODY = MAIL_BODY
            unparsed_variables.sMAIL_TO = MAIL_BODY
            cmd = New SqlCommand(sssql)
            If Debug_Menu = True Then
                Console.WriteLine(sssql)
            End If
            Return InsertUpdateData(cmd)
        Catch ex As Exception
            If Debug_Menu = True Then
                'MsgBox(ex.Message)
                error_msg.errmsg += ex.Message
            End If

        End Try
    End Function

    Public Function InsertUpdateData(ByVal cmd As SqlCommand) As Boolean
        Dim error_msg As New err_msgs
        Dim con As New SqlConnection(sConnectionString)
        Try

            cmd.CommandType = CommandType.Text
            cmd.Connection = con

            con.Open()
            cmd.ExecuteNonQuery()
            Return True
        Catch ex As Exception
            If debug_menu = True Then
                Console.WriteLine(ex.Message)
            End If
            error_msg.errmsg += ex.Message

        Finally
            check_errmsg_object(error_msg.errmsg)
            con.Close()
            con.Dispose()
        End Try
    End Function
End Module

