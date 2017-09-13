Imports System.Data.SQLite

Module Module1
    Public conn As SQLiteConnection = New SQLiteConnection

    Sub Execute(query As String)
        Dim cmd As SQLiteCommand = New SQLiteCommand(conn)
        cmd.CommandText = query
        My.Application.Log.WriteEntry("SQLite: " + cmd.CommandText, TraceEventType.Verbose)
        Try
            cmd.ExecuteNonQuery()
        Catch SQLiteExcep As SQLiteException
            My.Application.Log.WriteException(SQLiteExcep)
        End Try
    End Sub

    Sub Main()
        Dim HACProcesses As Diagnostics.Process() = Process.GetProcessesByName("HAController")
        For Each HACProcess As Process In HACProcesses
            HACProcess.CloseMainWindow()
        Next
        Threading.Thread.Sleep(5000)
        My.Application.Log.WriteEntry("Executing HACdb_Date_Conversion Script", TraceEventType.Warning)

        Dim Args As String() = Environment.GetCommandLineArgs()
        Dim DBFile As String = ""
        Dim LastEnvironmentIndex As Integer = 0
        Dim LastLocationIndex As Integer = 0
        For Each Arg As String In Args
            If Arg.Substring(0, 5) = "file=" Then
                DBFile = Arg.Substring(5, Arg.Length - 5)
            Else
                DBFile = "C:\HAC\Hacdb.sqlite"
            End If
            If Arg.Substring(0, 6) = "elast=" Then
                LastEnvironmentIndex = CInt(Arg.Substring(6, Arg.Length - 6))
                My.Application.Log.WriteEntry("Last environment database index to modify:" + CStr(LastEnvironmentIndex))
            End If
            If Arg.Substring(0, 6) = "llast=" Then
                LastLocationIndex = CInt(Arg.Substring(6, Arg.Length - 6))
                My.Application.Log.WriteEntry("Last location database index to modify:" + CStr(LastLocationIndex))
            End If
        Next

        Dim connstring As String = "URI=file:" + DBFile

        conn.ConnectionString = connstring
        Try
            My.Application.Log.WriteEntry("Connecting to database")
            conn.Open()
        Catch SQLiteExcep As SQLiteException
            My.Application.Log.WriteException(SQLiteExcep)
        End Try

        Dim strOldDate As String = ""
        Dim strNewDate As String = ""
        Dim intId As Integer = 0
        Dim strDateQuery As String = ""

        Dim strQuery As String = "SELECT * FROM ENVIRONMENT WHERE Id <= " + CStr(LastEnvironmentIndex)
        Dim esqlcommand As SQLiteCommand = New SQLiteCommand(strQuery, conn)
        Try
            Dim ereader As SQLiteDataReader = esqlcommand.ExecuteReader()

            Do While ereader.Read()
                strOldDate = ereader.GetString(1)
                strNewDate = Convert.ToDateTime(strOldDate).ToUniversalTime.ToString("u")
                intId = ereader.GetInt32(0)
                My.Application.Log.WriteEntry("Id: " + CStr(intId) + " / Old Date: " + strOldDate + " / New Date: " + strNewDate + " /" + CStr(ereader.StepCount))
                strDateQuery = "UPDATE ENVIRONMENT SET Date = """ + strNewDate + """ WHERE Id = " + CStr(intId)
                Execute(strDateQuery)
            Loop
        Catch SQLiteExcep As SQLiteException
            My.Application.Log.WriteException(SQLiteExcep)
        End Try


        strQuery = "SELECT * FROM LOCATION WHERE Id <= " + CStr(LastLocationIndex)
        Dim lsqlcommand As SQLiteCommand = New SQLiteCommand(strQuery, conn)
        Try
            Dim lreader As SQLiteDataReader = lsqlcommand.ExecuteReader()
            Do While lreader.Read()
                strOldDate = lreader.GetString(1)
                strNewDate = Convert.ToDateTime(strOldDate).ToUniversalTime.ToString("u")
                intId = lreader.GetInt32(0)
                My.Application.Log.WriteEntry("Id: " + CStr(intId) + " / Old Date: " + strOldDate + " / New Date: " + strNewDate + " /" + CStr(lreader.StepCount))
                strDateQuery = "UPDATE LOCATION SET Date = """ + strNewDate + """ WHERE Id = " + CStr(intId)
                Execute(strDateQuery)
            Loop
        Catch SQLiteExcep As SQLiteException
            My.Application.Log.WriteException(SQLiteExcep)
        End Try
    End Sub

End Module
