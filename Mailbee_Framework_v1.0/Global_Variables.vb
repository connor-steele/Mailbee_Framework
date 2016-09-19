
Module Global_Variables
    '----------------------------------------------------VARIABLES-----------------------------------------------
    Dim SMTP_Configuration As New SMTP_Settings
    Dim POP_Configuration As New POP_Settings


    'Randomly Generated number created by the parsing_function called "GenerateGUID()"
    Public GUID As String = ""

    'Gets the Unique Idenitifier automatically assigned to each email. Every email
    'will have a different UID.
    Public UID As String = ""


    'Checks for nulls
    Public Data_Exists As String = ""
    Public sGUID As String
    'List of Server names
    Public Names As String = ""
    Public Good_To_Go As Boolean = False
    'No html body
    Public No_HTML_Text As String = ""


    Public Error_Log_Path As String = ReadSetting("Error_Log")
    Public Email_Log_Path As String = ReadSetting("Email_Log")
    Public My_Log_Path1 As String = ReadSetting("Log_Path3")
    Public My_Log_Path2 As String = ReadSetting("Log_Path4")
    Public My_Log_Path3 As String = ReadSetting("Log_Path5")

    'sets the liscence key
    Public Licsence_Key As String = ReadSetting("MailBee.Global.LicenseKey")

    Public Database_Name As String = ReadSetting("Database_Name")

    '***********************************************SEARCH TERM SETTINGS ****************************************************************
    '  ******************************** SET THESE IN THE EXTERNAL APP.CONFIG FILE ****************************************************************
    Public Search_Term1 As String = ReadSetting("search_term1")
    Public Search_Term2 As String = ReadSetting("search_term2")
    Public Search_Term3 As String = ReadSetting("search_term3")
    Public Search_Term4 As String = ReadSetting("search_term4")
    Public Search_Term5 As String = ReadSetting("search_term5")

    '***********************************************SMTP SETTINGS ****************************************************************
    '*********************************************** OFFICE365 SETTINGS ****************************************************************
    Public SMTP_Office365_Username1 As String = smtp_configuration.SMTP_Office365_Username1
    Public SMTP_Office365_Password1 As String = smtp_configuration.SMTP_Office365_Password1
    Public SMTP_Office365_Host1 As String = smtp_configuration.SMTP_Office365_Host1
    Public SMTP_Office365_Port1 As String = smtp_configuration.SMTP_Office365_Port1

    'Public Sub 2_office_user

    'End Sub

    Public SMTP_Office365_Username2 As String = SMTP_Configuration.SMTP_Office365_Username2
    Public SMTP_Office365_Password2 As String = SMTP_Configuration.SMTP_Office365_Password2
    Public SMTP_Office365_Host2 As String = SMTP_Configuration.SMTP_Office365_Host2
    Public SMTP_Office365_Port2 As String = SMTP_Configuration.SMTP_Office365_Port2

    Public SMTP_Office365_Username3 As String = SMTP_Configuration.SMTP_Office365_Username3
    Public SMTP_Office365_Password3 As String = SMTP_Configuration.SMTP_Office365_Password3
    Public SMTP_Office365_Host3 As String = SMTP_Configuration.SMTP_Office365_Host3
    Public SMTP_Office365_Port3 As String = SMTP_Configuration.SMTP_Office365_Port3

    Public SMTP_Office365_Username4 As String = SMTP_Configuration.SMTP_Office365_Username4
    Public SMTP_Office365_Password4 As String = SMTP_Configuration.SMTP_Office365_Password4
    Public SMTP_Office365_Host4 As String = SMTP_Configuration.SMTP_Office365_Host4
    Public SMTP_Office365_Port4 As String = SMTP_Configuration.SMTP_Office365_Port4


    Public SMTP_Office365_Username5 As String = SMTP_Configuration.SMTP_Office365_Username5
    Public SMTP_Office365_Password5 As String = SMTP_Configuration.SMTP_Office365_Password5
    Public SMTP_Office365_Host5 As String = SMTP_Configuration.SMTP_Office365_Host5
    Public SMTP_Office365_Port5 As String = SMTP_Configuration.SMTP_Office365_Port5

    '*********************************************** GMAIL SETTINGS ****************************************************************
    Public SMTP_GMail_Username1 As String = smtp_configuration.SMTP_Gmail_Username1
    Public SMTP_GMail_Password1 As String = smtp_configuration.SMTP_Gmail_Password1
    Public SMTP_GMail_Host1 As String = smtp_configuration.SMTP_Gmail_Host1
    Public SMTP_Gmail_Port1 As String = smtp_configuration.SMTP_Gmail_Port1

    'Public SMTP_GMail_Username2 As String = smtp_configuration.SMTP_Gmail_Username2
    'Public SMTP_GMail_Password2 As String = smtp_configuration.SMTP_Gmail_Password2
    'Public SMTP_GMail_Host2 As String = smtp_configuration.SMTP_Gmail_Host2
    'Public SMTP_Gmail_Port2 As String = smtp_configuration.SMTP_Gmail_Port2

    'Public SMTP_GMail_Username3 As String = smtp_configuration.SMTP_Gmail_Username3
    'Public SMTP_GMail_Password3 As String = smtp_configuration.SMTP_Gmail_Password3
    'Public SMTP_GMail_Host3 As String = smtp_configuration.SMTP_Gmail_Host3
    'Public SMTP_Gmail_Port3 As String = smtp_configuration.SMTP_Gmail_Port3

    'Public SMTP_GMail_Username4 As String = smtp_configuration.SMTP_Gmail_Username4
    'Public SMTP_GMail_Password4 As String = smtp_configuration.SMTP_Gmail_Password4
    'Public SMTP_GMail_Host4 As String = smtp_configuration.SMTP_Gmail_Host4
    'Public SMTP_Gmail_Port4 As String = smtp_configuration.SMTP_Gmail_Port4

    'Public SMTP_GMail_Username5 As String = smtp_configuration.SMTP_Gmail_Username5
    'Public SMTP_GMail_Password5 As String = smtp_configuration.SMTP_Gmail_Password5
    'Public SMTP_GMail_Host5 As String = smtp_configuration.SMTP_Gmail_Host5
    'Public SMTP_Gmail_Port5 As String = smtp_configuration.SMTP_Gmail_Port5


    '***********************************************POP SETTINGS ****************************************************************
    '*********************************************** GMAIL SETTINGS ****************************************************************

    Public POP_GMail_Username1 As String = pop_configuration.POP_Gmail_Username1
    Public POP_GMail_Password1 As String = POP_configuration.POP_Gmail_Password1
    Public POP_GMail_Host1 As String = POP_configuration.POP_Gmail_Host1
    Public POP_GMail_Port1 As String = POP_configuration.POP_Gmail_Port1


    Public POP_GMail_Username2 As String = pop_configuration.POP_Gmail_Username2
    Public POP_GMail_Password2 As String = pop_configuration.POP_Gmail_Password2
    Public POP_GMail_Host2 As String = pop_configuration.POP_Gmail_Host2
    Public POP_GMail_port2 As String = pop_configuration.POP_Gmail_Port2

    Public POP_GMail_Username3 As String = pop_configuration.POP_Gmail_Username3
    Public POP_GMail_Password3 As String = pop_configuration.POP_Gmail_Password3
    Public POP_GMail_Host3 As String = pop_configuration.POP_Gmail_Host3
    Public POP_GMail_port3 As String = pop_configuration.POP_Gmail_Port3

    Public POP_GMail_Username4 As String = pop_configuration.POP_Gmail_Username4
    Public POP_GMail_Password4 As String = pop_configuration.POP_Gmail_Password4
    Public POP_GMail_Host4 As String = pop_configuration.POP_Gmail_Host4
    Public POP_GMail_port4 As String = pop_configuration.POP_Gmail_Port4

    Public POP_GMail_Username5 As String = pop_configuration.POP_Gmail_Username5
    Public POP_GMail_Password5 As String = pop_configuration.POP_Gmail_Password5
    Public POP_GMail_Host5 As String = pop_configuration.POP_Gmail_Host5
    Public POP_GMail_port5 As String = pop_configuration.POP_Gmail_Port5




    '----------------------------------------------------ERROR MESSAGE OBJECT-----------------------------------------------
    '--------------------------------This class is used to hold all of the error messages in the filepath.------------------
    Class Err_Msgs
        Friend Errmsg As String = ""
        Friend Filepath As String = ""
    End Class
    '----------------------------------------------------PARSED MAIL OBJECT-----------------------------------------------
    '--------------------------------------This class is used to hold all of the parsed_mail_variables.-------------------
    Class Parsed_Mail
        Friend Parsed_MAIL_TO As String = ""
        Friend Parsed_MAIL_FROM As String = ""
        Friend Parsed_MAIL_SUBJECT As String = ""
        Friend Parsed_MAIL_BODY As String = ""
        Friend Parsed_DATE_RECIEVED As String = ""
        Friend Parsed_SENT As String = ""
        Friend Parsed_RECIEVED As String = ""
        Friend UID As String = ""
    End Class
    '----------------------------------------------------UNPARSED MAIL OBJECT-----------------------------------------------
    '--------------------------------------This class is used to hold all of the unparsed_mail_variables.----------------
    Class Unparsed_Mail
        Friend sMAIL_TO As String = ""
        Friend sMAIL_FROM As String = ""
        Friend sMAIL_SUBJECT As String = ""
        Friend sMAIL_BODY As String = ""
        Friend sDATE_RECIEVED As String = ""
        Friend sSENT As String = ""
        Friend sRECIEVED As String = ""
    End Class
    '----------------------------------------------------qQUERY UNRESOLVED EMAILS OBJECT-----------------------------------------------
    '--------------------------------------This class is used to hold all of THE guids or uid's of an email----------------
    Class Query_Unresolved_Emails_Results
        Friend UID As String
    End Class
    '----------------------------------------------------SMTP SETTINGS OBJECT-----------------------------------------------
    Class SMTP_Settings
        'Uses the ReadSetting function to grab the correct variable. This function can be found in the parseing functions file.
        Friend Send_to As String = ReadSetting("send_to")
        Friend Send_to_Support As String = ReadSetting("Send_to_Support")


        '--------------------------------------------SMTP SETTINGS VARIABLES-----------------------------------------------

        '*********************************************** OFFICE365 SETTINGS ****************************************************************
        Friend SMTP_Office365_Username1 As String = ReadSetting("Office365.SMTP.Username1")
        Friend SMTP_Office365_Password1 As String = ReadSetting("Office365.SMTP.Password1")
        Friend SMTP_Office365_Host1 As String = ReadSetting("Office365.SMTP.Host1")
        Friend SMTP_Office365_Port1 As String = ReadSetting("Office365.SMTP.Port1")

        Friend SMTP_Office365_Username2 As String = ReadSetting("Office365.SMTP.Username2")
        Friend SMTP_Office365_Password2 As String = ReadSetting("Office365.SMTP.Password2")
        Friend SMTP_Office365_Host2 As String = ReadSetting("Office365.SMTP.Host2")
        Friend SMTP_Office365_Port2 As String = ReadSetting("Office365.SMTP.Port2")

        Friend SMTP_Office365_Username3 As String = ReadSetting("Office365.SMTP.Username3")
        Friend SMTP_Office365_Password3 As String = ReadSetting("Office365.SMTP.Password3")
        Friend SMTP_Office365_Host3 As String = ReadSetting("Office365.SMTP.Host3")
        Friend SMTP_Office365_Port3 As String = ReadSetting("Office365.SMTP.Port3")

        Friend SMTP_Office365_Username4 As String = ReadSetting("Office365.SMTP.Username4")
        Friend SMTP_Office365_Password4 As String = ReadSetting("Office365.SMTP.Password4")
        Friend SMTP_Office365_Host4 As String = ReadSetting("Office365.SMTP.Host4")
        Friend SMTP_Office365_Port4 As String = ReadSetting("Office365.SMTP.Port4")

        Friend SMTP_Office365_Username5 As String = ReadSetting("Office365.SMTP.Username5")
        Friend SMTP_Office365_Password5 As String = ReadSetting("Office365.SMTP.Password5")
        Friend SMTP_Office365_Host5 As String = ReadSetting("Office365.SMTP.Host5")
        Friend SMTP_Office365_Port5 As String = ReadSetting("Office365.SMTP.Port5")


        '*********************************************** GMAIL SETTINGS ****************************************************************

        Friend SMTP_Gmail_Username1 As String = ReadSetting("Gmail.STMP.Username1")
        Friend SMTP_Gmail_Password1 As String = ReadSetting("Gmail.STMP.Password1")
        Friend SMTP_Gmail_Host1 As String = ReadSetting("Gmail.STMP.host1")
        Friend SMTP_Gmail_Port1 As String = ReadSetting("Gmail.STMP.port1")

        Friend SMTP_Gmail_Username2 As String = ReadSetting("Gmail.STMP.Username2")
        Friend SMTP_Gmail_Password2 As String = ReadSetting("Gmail.STMP.Password2")
        Friend SMTP_Gmail_Host2 As String = ReadSetting("Gmail.STMP.host2")
        Friend SMTP_Gmail_Port2 As String = ReadSetting("Gmail.STMP.port2")

        Friend SMTP_Gmail_Username3 As String = ReadSetting("Gmail.STMP.Username3")
        Friend SMTP_Gmail_Password3 As String = ReadSetting("Gmail.STMP.Password3")
        Friend SMTP_Gmail_Host3 As String = ReadSetting("Gmail.STMP.host3")
        Friend SMTP_Gmail_Port3 As String = ReadSetting("Gmail.STMP.port3")

        Friend SMTP_Gmail_Username4 As String = ReadSetting("Gmail.STMP.Username4")
        Friend SMTP_Gmail_Password4 As String = ReadSetting("Gmail.STMP.Password4")
        Friend SMTP_Gmail_Host4 As String = ReadSetting("Gmail.STMP.host4")
        Friend SMTP_Gmail_Port4 As String = ReadSetting("Gmail.STMP.port4")

        Friend SMTP_Gmail_Username5 As String = ReadSetting("Gmail.STMP.Username5")
        Friend SMTP_Gmail_Password5 As String = ReadSetting("Gmail.STMP.Password5")
        Friend SMTP_Gmail_Host5 As String = ReadSetting("Gmail.STMP.host5")
        Friend SMTP_Gmail_Port5 As String = ReadSetting("Gmail.STMP.port5")

    End Class

    Class POP_Settings
        '----------------------------------------------------POP SETTINGS OBJECT-----------------------------------------------

        '*********************************************** OFFICE365 SETTINGS ****************************************************************

        Friend POP_Office365_Username1 As String = ReadSetting("Office365.Pop.Username1")
        Friend POP_Office365_Password1 As String = ReadSetting("Office365.Pop.Password1")
        Friend POP_Office365_Host1 As String = ReadSetting("Office365.Pop.Host1")
        Friend POP_Office365_Port1 As String = ReadSetting("Office365.Pop.Port1")

        Friend POP_Office365_Username2 As String = ReadSetting("Office365.Pop.Username2")
        Friend POP_Office365_Password2 As String = ReadSetting("Office365.Pop.Password2")
        Friend POP_Office365_Host2 As String = ReadSetting("Office365.Pop.Host2")
        Friend POP_Office365_Port2 As String = ReadSetting("Office365.Pop.Port2")

        Friend POP_Office365_Username3 As String = ReadSetting("Office365.Pop.Username3")
        Friend POP_Office365_Password3 As String = ReadSetting("Office365.Pop.Password3")
        Friend POP_Office365_Host3 As String = ReadSetting("Office365.Pop.Host3")
        Friend POP_Office365_Port3 As String = ReadSetting("Office365.Pop.Port3")

        Friend POP_Office365_Username4 As String = ReadSetting("Office365.Pop.Username4")
        Friend POP_Office365_Password4 As String = ReadSetting("Office365.Pop.Password4")
        Friend POP_Office365_Host4 As String = ReadSetting("Office365.Pop.Host4")
        Friend POP_Office365_Port4 As String = ReadSetting("Office365.Pop.Port4")

        Friend POP_Office365_Username5 As String = ReadSetting("Office365.Pop.Username5")
        Friend POP_Office365_Password5 As String = ReadSetting("Office365.Pop.Password5")
        Friend POP_Office365_Host5 As String = ReadSetting("Office365.Pop.Host5")
        Friend POP_Office365_Port5 As String = ReadSetting("Office365.Pop.Port5")
        '<!--*********************************************** GMAIL SETTINGS ****************************************************************-->

        Friend POP_Gmail_Username1 As String = ReadSetting("GMail.Pop.Username1")
        Friend POP_Gmail_Password1 As String = ReadSetting("GMail.Pop.Password1")
        Friend POP_Gmail_Host1 As String = ReadSetting("GMail.Pop.Host1")
        Friend POP_Gmail_Port1 As String = ReadSetting("GMail.Pop.Port1")

        Friend POP_Gmail_Username2 As String = ReadSetting("GMail.Pop.Username2")
        Friend POP_Gmail_Password2 As String = ReadSetting("GMail.Pop.Password2")
        Friend POP_Gmail_Host2 As String = ReadSetting("GMail.Pop.Host2")
        Friend POP_Gmail_Port2 As String = ReadSetting("GMail.Pop.Port2")

        Friend POP_Gmail_Username3 As String = ReadSetting("GMail.Pop.Username3")
        Friend POP_Gmail_Password3 As String = ReadSetting("GMail.Pop.Password3")
        Friend POP_Gmail_Host3 As String = ReadSetting("GMail.Pop.Host3")
        Friend POP_Gmail_Port3 As String = ReadSetting("GMail.Pop.Port3")

        Friend POP_Gmail_Username4 As String = ReadSetting("GMail.Pop.Username4")
        Friend POP_Gmail_Password4 As String = ReadSetting("GMail.Pop.Password4")
        Friend POP_Gmail_Host4 As String = ReadSetting("GMail.Pop.Host4")
        Friend POP_Gmail_Port4 As String = ReadSetting("GMail.Pop.Port4")

        Friend POP_Gmail_Username5 As String = ReadSetting("GMail.Pop.Username5")
        Friend POP_Gmail_Password5 As String = ReadSetting("GMail.Pop.Password5")
        Friend POP_Gmail_Host5 As String = ReadSetting("GMail.Pop.Host5")
        Friend POP_Gmail_Port5 As String = ReadSetting("GMail.Pop.Port5")

    End Class
End Module
