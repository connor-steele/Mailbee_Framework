Imports System.IO
Imports System.Text.RegularExpressions

Module File_Reading_Writing
    Dim error_msg As New err_msgs

    '----------------------------------------------------------------------------------------FILE WRITING/READING-----------------------------------------------
    Public Sub ReadSingleRow(ByVal record As IDataRecord)
        If debug_menu = True Then
            Console.WriteLine(String.Format("{0}, {1}", record(0), record(1)))
        End If
    End Sub


    Public Sub ErrorLogEvent(ByVal output As String)
        If error_log = True Then
            Dim strFile1 As String = String.Format("C:\logs\ErrorLog_{0}.txt", DateTime.Today.ToString("dd-MMM-yyyy"))
            Dim fileExists As Boolean = File.Exists(strFile1)
            Using sw As New StreamWriter(File.Open(strFile1, FileMode.OpenOrCreate))
                sw.WriteLine(("Error Message in  Occured at--" & output & Environment.NewLine), DateTime.Now, Environment.NewLine)
            End Using

        Else
            If debug_menu = True Then
                Console.WriteLine("Error Log is currently disabled. This can be toggled in the 'toggle_events.vb' module.")
            End If
        End If

        Dim directory As New IO.DirectoryInfo("C:\logs")


        If delete_log_1week = True Then
            For Each file As IO.FileInfo In directory.GetFiles
                If (Now - file.CreationTime).Days > 7 Then file.Delete()
            Next
        End If

        If delete_log_1month = True Then
            For Each file As IO.FileInfo In directory.GetFiles
                If (Now - file.CreationTime).Days > 30 Then file.Delete()
            Next
        End If
    End Sub


    ' SIMPLE TIME STAMPED LOG FILE

    Public Sub EmailLogEvent(ByVal output As String)
        If email_log = True Then
            Dim strFile2 As String = String.Format("C:\logs\EmailLog_{0}.txt", DateTime.Today.ToString("dd-MMM-yyyy"))
            Dim fileExists As Boolean = File.Exists(strFile2)
            Using sw As New StreamWriter(File.Open(strFile2, FileMode.OpenOrCreate))
                sw.WriteLine(output)
            End Using

        Else
            If debug_menu = True Then
                Console.WriteLine("Email Log is currently disabled. This can be toggled in the 'toggle_events.vb' module.")
            End If

        End If

        Dim directory As New IO.DirectoryInfo("C:\logs")


        If delete_log_1week = True Then
            For Each file As IO.FileInfo In directory.GetFiles
                If (Now - file.CreationTime).Days > 7 Then file.Delete()
            Next
        End If

        If delete_log_1month = True Then
            For Each file As IO.FileInfo In directory.GetFiles
                If (Now - file.CreationTime).Days > 30 Then file.Delete()
            Next
        End If


    End Sub

    ' GETS APPLICATION NAME
    Public Function ApplicationName() As String

        Return System.IO.Path.GetFileNameWithoutExtension(System.Reflection.Assembly.GetExecutingAssembly().Location)

    End Function

    ' GETS APPLICATION PATH
    Public Function ApplicationPath() As String

        Return System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location)

    End Function

    '--------------------------------------------------------------------------------READS ALL TEXT USING FILE PATH AS A PARAM-----------------------------------------------

    Public Function read_all_text(filepath As String)
        Dim objStreamReader As StreamReader
        Dim strLine As String
        Dim sanitized_file_path As String = filepath.Replace("'", "")
        'Pass the file path and the file name to the StreamReader constructor.
        objStreamReader = New StreamReader(sanitized_file_path)
        'Read the first line of text.
        strLine = objStreamReader.ReadLine
        'Continue to read until you reach the end of the file.
        Do While Not strLine Is Nothing
            'Write the line to the Console window.
            If debug_menu = True Then
                Console.WriteLine(strLine)
            End If
            'Read the next line.
            strLine = objStreamReader.ReadLine
            Dim regex As Regex = New Regex("(?i)^console.writeline()(?-i)")
            Dim match As Match = regex.Match(strLine)

            If match.Success Then
                If debug_menu = True Then
                    Console.WriteLine(match.Value)
                End If
            End If
        Loop
        'Close the file.
        objStreamReader.Close()
        Console.ReadLine()
    End Function
    Public Function get_my_file_path(filepath As String) As err_msgs
        Dim fullPath As String
        Dim error_msgs As New err_msgs
        fullPath = Path.GetFullPath(filepath)
        'MsgBox(fullPath)
        error_msgs.filepath = fullPath
        'MsgBox(error_msgs.filepath)
        'my_file_path1 = filepath
        If debug_menu = True Then
            Console.WriteLine("GetFullPath('{0}') returns '{1}'",
    filepath, fullPath)
        End If

        ' Output is based on your current directory, except
        ' in the last case, where it is based on the root drive
        ' GetFullPath('mydir') returns 'C:\temp\Demo\mydir'
        ' GetFullPath('myfile.ext') returns 'C:\temp\Demo\myfile.ext'
        ' GetFullPath('\mydir') returns 'C:\mydir'
    End Function
    Public Function FileNameWithoutExtension(ByVal FullPath _
        As String) As String
        MsgBox(System.IO.Path.GetFileNameWithoutExtension(FullPath))
        Return System.IO.Path.GetFileNameWithoutExtension(FullPath)

    End Function
    Public Sub read_windows_security_eventlogging()
        Dim aLog As EventLog
        Dim myLog As New EventLog
        Dim aEventLogList() As EventLog
        Dim aLogEntry As EventLogEntry
        Dim aLogEntries As EventLogEntryCollection
        Dim logtype As String = "Security"
        '        ' from which you want to read the logs 
        Dim evtLog As New EventLog(logtype, System.Environment.MachineName)
        ' Create a new log.
        '
        'If Not EventLog.SourceExists("Security") Then
        '    EventLog.CreateEventSource("MyNewSource", "MyNewLog")
        'End If
        '
        ' Add an event to fire when an entry is written.
        '
        AddHandler myLog.EntryWritten, AddressOf OnEntryWritten

        With evtLog
            .Source = "Security"
            aLogEntries = .Entries()
            For Each aLogEntry In aLogEntries
                With aLogEntry
                    Console.WriteLine(
              "Source: {0}" & vbCrLf &
              "Category: {1}" & vbCrLf &
              "Message: {2}" & vbCrLf &
              "EntryType: {3}" & vbCrLf &
              "EventID: {4}" & vbCrLf &
              "UserName: {5}",
              .Source, .Category, .Message, .EntryType, .EventID, .UserName)
                End With

                'Threading.Thread.Sleep(200)

            Next
            '
            ' Delete your new log.
            '
            '.Clear()
            '.Deete("MyNewLog")
            '
            ' Output the names of all logs on the system.
            '
            aEventLogList = .GetEventLogs()
            For Each aLog In aEventLogList
                Console.WriteLine("Log name: " & aLog.LogDisplayName)

            Next
        End With


    End Sub
End Module
