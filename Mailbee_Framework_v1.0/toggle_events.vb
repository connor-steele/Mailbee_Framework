Module toggle_events

    Public SMTP_mail_Count As Boolean = False


    'turns all of the console messages on or off with true or false
    Public Debug_Menu As Boolean = True

    'Creates an error log in C:/error{0}.txt if true
    Public Error_Log As Boolean = True

    'Creates an error log in C:/error{0}.txt if true
    Public Email_Log As Boolean = True

    'deletes all logs older than 1 week
    Public Delete_Log_1Week As Boolean = False

    'deletes all logs older than 1 month
    Public Delete_Log_1Month As Boolean = True


    Public Delete_Email_After_Read = True



    'Experimental / Unused Features


    ''puts the application into parameterized application
    'Public take_in_arguement_mode As Boolean = False

End Module
