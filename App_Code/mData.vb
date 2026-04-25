Imports Microsoft.VisualBasic
Imports System.Data.SqlClient
Imports System.Data
Imports System.Data.OleDb
Imports System.IO
Imports System.Math
Imports System.Collections
Imports MySql.Data.MySqlClient
Imports Microsoft.VisualBasic.DateAndTime
Imports Oracle.ManagedDataAccess.Client
Imports Newtonsoft.Json
Imports System.Net
Imports Mysqlx
Imports System.Threading
Public Module mData
    Public userdbcase As String

    Public Function GetDataBase(userconstr As String, Optional ByVal userconprv As String = "") As String
        Dim ret As String = "DB"
        If userconstr.ToUpper.IndexOf("NAMESPACE") >= 0 Then
            Dim items As String() = userconstr.Split(";")
            For i As Integer = 0 To items.Count - 1
                If items(i).ToUpper.Contains("NAMESPACE") Then
                    ret = Piece(items(i), "=", 2).Trim
                    Exit For
                End If
            Next
        ElseIf userconstr.ToUpper.IndexOf("DATABASE") >= 0 Then
            Dim items As String() = userconstr.Split(";")
            For i As Integer = 0 To items.Count - 1
                If items(i).ToUpper.Contains("DATABASE") Then
                    ret = Piece(items(i), "=", 2).Trim
                    Exit For
                End If
            Next
        ElseIf userconstr.ToUpper.IndexOf("DATA SOURCE") >= 0 Then
            Dim items As String() = userconstr.Split(";")
            If userconprv = "Oracle.ManagedDataAccess.Client" Then
                ret = GetUserIDFromConnectionString(userconstr)  '.Replace("C##", "")
                Return ret
            End If

            ret = items(0)
            Return ret
            If userconstr.IndexOf("/") > 0 Then
                ret = userconstr.Substring(userconstr.LastIndexOf("/"))
                ret = ret.Replace(".", "_").Replace(";", "").Replace("/", "")
                Return ret
            End If
        ElseIf userconstr.ToUpper.IndexOf("DSN") >= 0 Then
            Dim items As String() = userconstr.Split(";")
            For i As Integer = 0 To items.Count - 1
                If items(i).ToUpper.Contains("DSN") Then
                    ret = Piece(items(i), "=", 2).Trim
                    Exit For
                End If
            Next
        End If
        If userconprv <> "Oracle.ManagedDataAccess.Client" Then
            ret = FixReservedWords(ret, userconprv)
        End If
        Return ret
    End Function
    Public Function CorrectSQLforSQLServer(ByVal sql As String) As String
        sql = sql.Replace("OUR.HelpDesk", "OURHelpDesk")
        sql = sql.Replace("OUR.AccessLog", "OURAccessLog")
        sql = sql.Replace("OUR.Permits", "OURPermits")
        sql = sql.Replace("OUR.Permissions", "OURPermissions")
        sql = sql.Replace("OUR.ReportInfo", "OURReportInfo")
        sql = sql.Replace("OUR.ReportShow", "OURReportShow")
        sql = sql.Replace("OUR.ReportSQLquery", "OURReportSQLquery")
        sql = sql.Replace("OUR.ReportGroups", "OURReportGroups")
        sql = sql.Replace("OUR.ReportFormat", "OURReportFormat")
        sql = sql.Replace("OUR.ReportLists", "OURReportLists")
        sql = sql.Replace("OUR.ReportChildren", "OURReportChildren")
        sql = sql.Replace("OUR.FriendlyNames", "OURFriendlyNames")
        sql = sql.Replace("OUR.Files", "OURFiles")
        sql = sql.Replace("OURFILES", "OURFiles")
        sql = sql.Replace("OURFRIENDLYNAMES", "OURFriendlyNames")
        sql = sql.Replace("OURHELPDESK", "OURHelpDesk")
        sql = sql.Replace("OURACCESSLOG", "OURAccessLog")
        sql = sql.Replace("OURPERMITS", "OURPermits")
        sql = sql.Replace("OURPERMISSIONS", "OURPermissions")
        sql = sql.Replace("OURREPORTINFO", "OURReportInfo")
        sql = sql.Replace("OURREPORTSHOW", "OURReportShow")
        sql = sql.Replace("OURREPORTSQLQUERY", "OURReportSQLquery")
        sql = sql.Replace("OURREPORTGROUPS", "OURReportGroups")
        sql = sql.Replace("OURREPORTFORMAT", "OURReportFormat")
        sql = sql.Replace("OURREPORTLISTS", "OURReportLists")
        sql = sql.Replace("OURREPORTCHILDREN", "OURReportChildren")
        sql = sql.Replace("ourfiles", "OURFiles")
        sql = sql.Replace("ourfriendlynames", "OURFriendlyNames")
        sql = sql.Replace("ourhelpdesk", "OURHelpDesk")
        sql = sql.Replace("ouraccesslog", "OURAccessLog")
        sql = sql.Replace("ourpermits", "OURPermits")
        sql = sql.Replace("ourpermissions", "OURPermissions")
        sql = sql.Replace("ourreportinfo", "OURReportInfo")
        sql = sql.Replace("ourreportshow", "OURReportShow")
        sql = sql.Replace("ourreportsqlquery", "OURReportSQLquery")
        sql = sql.Replace("ourreportgroups", "OURReportGroups")
        sql = sql.Replace("ourreportformat", "OURReportFormat")
        sql = sql.Replace("ourreportlists", "OURReportLists")
        sql = sql.Replace("ourreportchildren", "OURReportChildren")

        Return sql
    End Function
    Public Function CorrectSQLforCache(ByVal sql As String, Optional ByVal donotreplace As String = "") As String
        'sql = sql.Replace("OURHelpDesk", "OUR.HelpDesk")
        'sql = sql.Replace("OURAccessLog", "OUR.AccessLog")
        'sql = sql.Replace("OURPermits", "OUR.Permits")
        'sql = sql.Replace("OURPermissions", "OUR.Permissions")
        'sql = sql.Replace("OURReportInfo", "OUR.ReportInfo")
        'sql = sql.Replace("OURReportShow", "OUR.ReportShow")
        'sql = sql.Replace("OURReportSQLquery", "OUR.ReportSQLquery")
        'sql = sql.Replace("OURReportGroups", "OUR.ReportGroups")
        'sql = sql.Replace("OURReportFormat", "OUR.ReportFormat")
        'sql = sql.Replace("OURReportLists", "OUR.ReportLists")
        'sql = sql.Replace("OURReportChildren", "OUR.ReportChildren")
        'sql = sql.Replace("OURFriendlyNames", "OUR.FriendlyNames")
        'sql = sql.Replace("OURFiles", "OUR.Files")
        'sql = sql.Replace("OURUnits", "OUR.Units")
        'sql = sql.Replace("OURUserTables", "OUR.UserTables")
        'sql = sql.Replace("OURUSERTABLES", "OUR.UserTables")
        'sql = sql.Replace("OURUNITS", "OUR.Units")
        'sql = sql.Replace("OURFILES", "OUR.Files")
        'sql = sql.Replace("OURFRIENDLYNAMES", "OUR.FriendlyNames")
        'sql = sql.Replace("OURHELPDESK", "OUR.HelpDesk")
        'sql = sql.Replace("OURACCESSLOG", "OUR.AccessLog")
        'sql = sql.Replace("OURPERMITS", "OUR.Permits")
        'sql = sql.Replace("OURPERMISSIONS", "OUR.Permissions")
        'sql = sql.Replace("OURREPORTINFO", "OUR.ReportInfo")
        'sql = sql.Replace("OURREPORTSHOW", "OUR.ReportShow")
        'sql = sql.Replace("OURREPORTSQLQUERY", "OUR.ReportSQLquery")
        'sql = sql.Replace("OURREPORTGROUPS", "OUR.ReportGroups")
        'sql = sql.Replace("OURREPORTFORMAT", "OUR.ReportFormat")
        'sql = sql.Replace("OURREPORTLISTS", "OUR.ReportLists")
        'sql = sql.Replace("OURREPORTCHILDREN", "OUR.ReportChildren")
        'sql = sql.Replace("OURAgents", "OUR.Agents")
        'sql = sql.Replace("OURAGENTS", "OUR.Agents")
        'sql = sql.Replace("OURActivity", "OUR.Activity")
        'sql = sql.Replace("OURACTIVITY", "OUR.Activity")
        'sql = sql.Replace("ourunits", "OUR.Units")
        'sql = sql.Replace("ourfiles", "OUR.Files")
        'sql = sql.Replace("ourfriendlynames", "OUR.FriendlyNames")
        'sql = sql.Replace("ourhelpdesk", "OUR.HelpDesk")
        'sql = sql.Replace("ouraccesslog", "OUR.AccessLog")
        'sql = sql.Replace("ourpermits", "OUR.Permits")
        'sql = sql.Replace("ourpermissions", "OUR.Permissions")
        'sql = sql.Replace("ourreportinfo", "OUR.ReportInfo")
        'sql = sql.Replace("ourreportshow", "OUR.ReportShow")
        'sql = sql.Replace("ourreportsqlquery", "OUR.ReportSQLquery")
        'sql = sql.Replace("ourreportgroups", "OUR.ReportGroups")
        'sql = sql.Replace("ourreportformat", "OUR.ReportFormat")
        'sql = sql.Replace("ourreportlists", "OUR.ReportLists")
        'sql = sql.Replace("ourreportchildren", "OUR.ReportChildren")
        'If sql.ToUpper.Contains(" AFTER [") Then
        '    sql = sql.Substring(0, sql.ToUpper.IndexOf("AFTER")).Trim
        'End If
        ''sql = sql.Substring(0, sql.ToUpper.IndexOf("AFTER")).Trim
        'sql = sql.Replace("[Order]", """Order""").Replace("[Group]", """Group""").Replace("[User]", """User""").Replace("[Count]", """Count""")
        'sql = sql.Replace("`Order`", """Order""").Replace("`Group`", """Group""").Replace("`User`", """User""").Replace("`Count`", """Count""")
        'If donotreplace = "" Then sql = sql.Replace("[", "").Replace("]", "")
        'Return sql

        sql = sql.Replace("OURHelpDesk", "OUR.HelpDesk")
        sql = sql.Replace("OURAccessLog", "OUR.AccessLog")
        sql = sql.Replace("OURDashboards", "OUR.Dashboards")
        sql = sql.Replace("OURPermits", "OUR.Permits")
        sql = sql.Replace("OURPermissions", "OUR.Permissions")
        sql = sql.Replace("OURReportInfo", "OUR.ReportInfo")
        sql = sql.Replace("OURReportShow", "OUR.ReportShow")
        sql = sql.Replace("OURReportSQLquery", "OUR.ReportSQLquery")
        sql = sql.Replace("OURReportGroups", "OUR.ReportGroups")
        sql = sql.Replace("OURReportFormat", "OUR.ReportFormat")
        sql = sql.Replace("OURReportLists", "OUR.ReportLists")
        sql = sql.Replace("OURReportChildren", "OUR.ReportChildren")
        sql = sql.Replace("OURFriendlyNames", "OUR.FriendlyNames")
        sql = sql.Replace("OURFiles", "OUR.Files")
        sql = sql.Replace("OURUnits", "OUR.Units")
        sql = sql.Replace("OURUserTables", "OUR.UserTables")
        sql = sql.Replace("OURReportView", "OUR.ReportView")
        sql = sql.Replace("OURReportItems", "OUR.ReportItems")
        sql = sql.Replace("OurHelpDesk", "OUR.HelpDesk")
        sql = sql.Replace("OurAccessLog", "OUR.AccessLog")
        sql = sql.Replace("OurDashboards", "OUR.Dashboards")
        sql = sql.Replace("OurPermits", "OUR.Permits")
        sql = sql.Replace("OurPermissions", "OUR.Permissions")
        sql = sql.Replace("OurReportInfo", "OUR.ReportInfo")
        sql = sql.Replace("OurReportShow", "OUR.ReportShow")
        sql = sql.Replace("OurReportSQLquery", "OUR.ReportSQLquery")
        sql = sql.Replace("OurReportGroups", "OUR.ReportGroups")
        sql = sql.Replace("OurReportFormat", "OUR.ReportFormat")
        sql = sql.Replace("OurReportLists", "OUR.ReportLists")
        sql = sql.Replace("OurReportView", "OUR.ReportView")
        sql = sql.Replace("OurReportItems", "OUR.ReportItems")
        sql = sql.Replace("OurReportChildren", "OUR.ReportChildren")
        sql = sql.Replace("OurFriendlyNames", "OUR.FriendlyNames")
        sql = sql.Replace("OurFiles", "OUR.Files")
        sql = sql.Replace("OurUnits", "OUR.Units")
        sql = sql.Replace("OurUserTables", "OUR.UserTables")
        sql = sql.Replace("OURUSERTABLES", "OUR.UserTables")
        sql = sql.Replace("OURUNITS", "OUR.Units")
        sql = sql.Replace("OURFILES", "OUR.Files")
        sql = sql.Replace("OURFRIENDLYNAMES", "OUR.FriendlyNames")
        sql = sql.Replace("OURHELPDESK", "OUR.HelpDesk")
        sql = sql.Replace("OURACCESSLOG", "OUR.AccessLog")
        sql = sql.Replace("OURDASHBOARDS", "OUR.Dashboards")
        sql = sql.Replace("OURPERMITS", "OUR.Permits")
        sql = sql.Replace("OURPERMISSIONS", "OUR.Permissions")
        sql = sql.Replace("OURREPORTINFO", "OUR.ReportInfo")
        sql = sql.Replace("OURREPORTSHOW", "OUR.ReportShow")
        sql = sql.Replace("OURREPORTSQLQUERY", "OUR.ReportSQLquery")
        sql = sql.Replace("OURREPORTGROUPS", "OUR.ReportGroups")
        sql = sql.Replace("OURREPORTFORMAT", "OUR.ReportFormat")
        sql = sql.Replace("OURREPORTLISTS", "OUR.ReportLists")
        sql = sql.Replace("OURREPORTCHILDREN", "OUR.ReportChildren")
        sql = sql.Replace("OURREPORTVIEW", "OUR.ReportView")
        sql = sql.Replace("OURREPORTITEMS", "OUR.ReportItems")
        sql = sql.Replace("OURAgents", "OUR.Agents")
        sql = sql.Replace("OURAGENTS", "OUR.Agents")
        sql = sql.Replace("OURActivity", "OUR.Activity")
        sql = sql.Replace("OURACTIVITY", "OUR.Activity")
        sql = sql.Replace("ourunits", "OUR.Units")
        sql = sql.Replace("ourfiles", "OUR.Files")
        sql = sql.Replace("ourfriendlynames", "OUR.FriendlyNames")
        sql = sql.Replace("ourhelpdesk", "OUR.HelpDesk")
        sql = sql.Replace("ouraccesslog", "OUR.AccessLog")
        sql = sql.Replace("ourdashboards", "OUR.Dashboards")
        sql = sql.Replace("ourpermits", "OUR.Permits")
        sql = sql.Replace("ourpermissions", "OUR.Permissions")
        sql = sql.Replace("ourreportinfo", "OUR.ReportInfo")
        sql = sql.Replace("ourreportshow", "OUR.ReportShow")
        sql = sql.Replace("ourreportsqlquery", "OUR.ReportSQLquery")
        sql = sql.Replace("ourreportgroups", "OUR.ReportGroups")
        sql = sql.Replace("ourreportformat", "OUR.ReportFormat")
        sql = sql.Replace("ourreportlists", "OUR.ReportLists")
        sql = sql.Replace("ourreportchildren", "OUR.ReportChildren")
        sql = sql.Replace("ourreportview", "OUR.ReportView")
        sql = sql.Replace("ourreportitems", "OUR.ReportItems")

        sql = sql.Replace("ourusertables", "OUR.UserTables")

        sql = sql.Replace("ourtasklistsetting", "OUR.TaskListSetting")
        sql = sql.Replace("OURTASKLISTSETTING", "OUR.TaskListSetting")
        sql = sql.Replace("OurTaskListSetting", "OUR.TaskListSetting")
        sql = sql.Replace("OURTaskListSetting", "OUR.TaskListSetting")

        sql = sql.Replace("ourkmlhistory", "OUR.KMLHistory")
        sql = sql.Replace("OURKMLHISTORY", "OUR.KMLHistory")
        sql = sql.Replace("OURKMLHistory", "OUR.KMLHistory")
        sql = sql.Replace("OurKMLHistory", "OUR.KMLHistory")
        sql = sql.Replace("OurKmlHistory", "OUR.KMLHistory")
        sql = sql.Replace("OURKmlHistory", "OUR.KMLHistory")

        sql = sql.Replace("ourscheduledreports", "OUR.ScheduledReports")
        sql = sql.Replace("OURSCHEDULEDREPORTS", "OUR.ScheduledReports")
        sql = sql.Replace("OurScheduledReports", "OUR.ScheduledReports")
        sql = sql.Replace("OURScheduledReports", "OUR.ScheduledReports")

        sql = sql.Replace("ourscheduleddownloads", "OUR.ScheduledDownloads")
        sql = sql.Replace("OURSCHEDULEDDOWNLOADS", "OUR.ScheduledDownloads")
        sql = sql.Replace("OurScheduledDownloads", "OUR.ScheduledDownloads")
        sql = sql.Replace("OURScheduledDownloads", "OUR.ScheduledDownloads")

        sql = sql.Replace("ourscheduledimports", "OUR.ScheduledImports")
        sql = sql.Replace("OURSCHEDULEDIMPORTS", "OUR.ScheduledImports")
        sql = sql.Replace("OurScheduledImports", "OUR.ScheduledImports")
        sql = sql.Replace("OURScheduledImports", "OUR.ScheduledImports")
        sql = sql.Replace("ourscheduledImports", "OUR.ScheduledImports")

        'OURComparison
        sql = sql.Replace("OURComparison", "OUR.Comparison")
        sql = sql.Replace("OurComparison", "OUR.Comparison")
        sql = sql.Replace("OURCOMPARISON", "OUR.Comparison")
        sql = sql.Replace("OURcomparison", "OUR.Comparison")
        sql = sql.Replace("ourcomparison", "OUR.Comparison")
        sql = sql.Replace("User.", "")

        sql = sql.Replace("[[", "[").Replace("]]", "]").Replace("``", "`")
        '.Replace(";", "") - no, cut last ; in the end
        If sql.EndsWith(";") Then
            sql = sql.Substring(0, sql.Length - 1)
        End If
        If sql.ToUpper.Contains(" AFTER [") Then
            sql = sql.Substring(0, sql.ToUpper.IndexOf("AFTER")).Trim
        End If
        'TODO make function to  replace [] aroung reserved word to ""
        sql = sql.Replace("[Order]", """Order""").Replace("[Group]", """Group""").Replace("[User]", """User""").Replace("[Count]", """Count""").Replace("[Status]", """Status""").Replace("[Section]", """Section""")
        sql = sql.Replace("`Order`", """Order""").Replace("`Group`", """Group""").Replace("`User`", """User""").Replace("`Count`", """Count""").Replace("`Status`", """Status""").Replace("`Section`", """Section""")

        If donotreplace = "" Then sql = sql.Replace("[", "").Replace("]", "").Replace("`", "")
        Return sql
    End Function
    Public Function CorrectSQLforMySql(ByVal sql As String, ByRef myconstring As String, Optional ByVal donotreplace As String = "") As String
        sql = sql.Replace("OURHelpDesk", "ourhelpdesk")
        sql = sql.Replace("OURAccessLog", "ouraccesslog")
        sql = sql.Replace("OURPermits", "ourpermits")
        sql = sql.Replace("OURPermissions", "ourpermissions")
        sql = sql.Replace("OURReportInfo", "ourreportinfo")
        sql = sql.Replace("OURReportShow", "ourreportshow")
        sql = sql.Replace("OURReportSQLquery", "ourreportsqlquery")
        sql = sql.Replace("OURReportGroups", "ourreportgroups")
        sql = sql.Replace("OURReportFormat", "ourreportformat")
        sql = sql.Replace("OURReportLists", "ourreportlists")
        sql = sql.Replace("OURReportChildren", "ourreportchildren")
        sql = sql.Replace("OURFriendlyNames", "ourfriendlynames")
        sql = sql.Replace("OURFiles", "ourfiles")
        sql = sql.Replace("OURUnits", "ourunits")
        sql = sql.Replace("OURUserTables", "ourusertables")
        sql = sql.Replace("OURUSERTABLES", "ourusertables")
        sql = sql.Replace("OURAgents", "ouragents")
        sql = sql.Replace("OURAGENTS", "ouragents")
        sql = sql.Replace("OURActivity", "ouractivity")
        sql = sql.Replace("OURACTIVITY", "ouractivity")
        sql = sql.Replace("OURUNITS", "ourunits")
        sql = sql.Replace("OURFILES", "ourfiles")
        sql = sql.Replace("OURFRIENDLYNAMES", "ourfriendlynames")
        sql = sql.Replace("OURHELPDESK", "ourhelpdesk")
        sql = sql.Replace("OURACCESSLOG", "ouraccessLog")
        sql = sql.Replace("OURPERMITS", "ourpermits")
        sql = sql.Replace("OURPERMISSIONS", "ourpermissions")
        sql = sql.Replace("OURREPORTINFO", "ourreportinfo")
        sql = sql.Replace("OURREPORTSHOW", "ourreportshow")
        sql = sql.Replace("OURREPORTSQLQUERY", "ourreportsqlquery")
        sql = sql.Replace("OURREPORTGROUPS", "ourreportgroups")
        sql = sql.Replace("OURREPORTFORMAT", "ourreportformat")
        sql = sql.Replace("OURREPORTLISTS", "ourreportlists")
        sql = sql.Replace("OURREPORTCHILDREN", "ourreportchildren")
        If donotreplace = "" Then sql = sql.Replace("[", "`").Replace("]", "`")
        'remove TOP N
        Dim t As Integer = sql.ToUpper.IndexOf(" TOP ")
        If t > 0 AndAlso sql.Trim.ToUpper.StartsWith("SELECT ") Then
            Dim temp As String = sql.Substring(t + 5).Trim
            t = temp.IndexOf(" ")
            If t > 0 Then
                sql = "SELECT " & temp.Substring(t + 1)
            End If
        End If
        'correct myconstring
        'myconstring = (myconstring & "; Convert Zero Datetime=True; Allow Zero Datetime=True;").Replace(";;", ";")
        Return sql
    End Function
    Public Function HasRecords(ByVal mySQL As String, Optional ByVal myconstr As String = "", Optional ByVal myconprv As String = "", Optional ByRef er As String = "") As Boolean
        mySQL = mySQL.Replace("""", "'")
        Dim hasrec As Boolean
        mySQL = cleanSQL(mySQL)
        mySQL = Replace(mySQL, Chr(34), Chr(39))
        Dim myRecords As DataTable = New DataTable
        Dim myconstring, myprovider As String
        Dim dbcase As String = String.Empty
        Try
            If myconstr.Trim = "" Then
                myconstring = System.Configuration.ConfigurationManager.ConnectionStrings.Item("mySQLconnection").ToString
                myprovider = System.Configuration.ConfigurationManager.ConnectionStrings.Item("mySQLconnection").ProviderName.ToString
                If myprovider = "Npgsql" Then dbcase = ConfigurationManager.AppSettings("ourdbcase").ToString
            Else
                If myconprv = "Npgsql" Then
                    If System.Configuration.ConfigurationManager.ConnectionStrings.Item("CSVconnection") IsNot Nothing AndAlso myconstr.Trim.ToUpper = System.Configuration.ConfigurationManager.ConnectionStrings.Item("CSVconnection").ToString.Trim.ToUpper Then
                        dbcase = ConfigurationManager.AppSettings("csvdbcase").ToString
                    ElseIf myconstr.Trim.ToUpper = System.Configuration.ConfigurationManager.ConnectionStrings.Item("mySQLconnection").ToString.Trim.ToUpper Then
                        dbcase = ConfigurationManager.AppSettings("ourdbcase").ToString
                    ElseIf myconstr.Trim.ToUpper = System.Configuration.ConfigurationManager.ConnectionStrings.Item("UserSqlConnection").ToString.Trim.ToUpper Then
                        dbcase = ConfigurationManager.AppSettings("userdbcase").ToString
                    Else 'postgres, etc...
                        dbcase = ConfigurationManager.AppSettings("ourdbcase").ToString
                    End If
                End If
                myconstring = myconstr
                myprovider = myconprv
            End If

            If myprovider = "InterSystems.Data.IRISClient" Then
                mySQL = CorrectSQLforCache(mySQL)
                Dim propResult As String = HasRecords_IRIS(mySQL, myconstring, myRecords)
                If propResult = "RETURN_FALSE" Then Return False

            ElseIf myprovider = "InterSystems.Data.CacheClient" Then
                mySQL = CorrectSQLforCache(mySQL)
                Dim propResult As String = HasRecords_Cache(mySQL, myconstring, myRecords)
                If propResult = "RETURN_FALSE" Then Return False


            ElseIf myprovider = "Npgsql" Then  'PostgreSQL  Npgsql
                myconstring = CorrectConnstringForPostgres(myconstring, dbcase)
                mySQL = ConvertFromSqlServerToPostgres(mySQL, dbcase, myconstring, myprovider)
                Dim myConnection As Npgsql.NpgsqlConnection
                Dim myCommand As New Npgsql.NpgsqlCommand
                myConnection = New Npgsql.NpgsqlConnection(myconstring)
                myCommand.Connection = myConnection
                myCommand.CommandType = CommandType.Text
                myCommand.CommandTimeout = 300000
                myCommand.CommandText = mySQL
                If myCommand.Connection.State = ConnectionState.Closed Then myCommand.Connection.Open()
                Dim myAdapter As New Npgsql.NpgsqlDataAdapter(myCommand)
                myAdapter.Fill(myRecords)
                myAdapter.Dispose()
                myCommand.Connection.Close()
                myCommand.Dispose()

            ElseIf myprovider = "System.Data.SqlClient" Then
                Dim myConnection As SqlConnection
                Dim myCommand As New SqlClient.SqlCommand
                Dim myAdapter As SqlClient.SqlDataAdapter
                myConnection = New SqlConnection(myconstring)
                myCommand.Connection = myConnection
                myCommand.CommandType = CommandType.Text
                myCommand.CommandTimeout = 300000
                myCommand.CommandText = mySQL
                If myCommand.Connection.State = ConnectionState.Closed Then myCommand.Connection.Open()
                myAdapter = New SqlClient.SqlDataAdapter(myCommand)
                myAdapter.Fill(myRecords)
                myAdapter.Dispose()
                myCommand.Connection.Close()
                myCommand.Dispose()
            ElseIf myprovider = "MySql.Data.MySqlClient" Then
                mySQL = CorrectSQLforMySql(mySQL, myconstring)
                Dim myConnection As MySqlConnection
                Dim myCommand As New MySqlCommand
                Dim myAdapter As MySqlDataAdapter
                myConnection = New MySqlConnection(myconstring)
                myCommand.Connection = myConnection
                myCommand.CommandType = CommandType.Text
                myCommand.CommandTimeout = 300000
                myCommand.CommandText = mySQL
                If myCommand.Connection.State = ConnectionState.Closed Then myCommand.Connection.Open()
                myAdapter = New MySqlDataAdapter(myCommand)
                myAdapter.Fill(myRecords)
                myAdapter.Dispose()
                myCommand.Connection.Close()
                myCommand.Dispose()
            ElseIf myprovider = "Oracle.ManagedDataAccess.Client" Then
                mySQL = CorrectSQLforOracle(mySQL, myconstring)
                HasRecords_Oracle(mySQL, myconstring, myRecords)

            End If
            If myRecords.Rows.Count > 0 Then
                hasrec = True
            Else
                hasrec = False
            End If
        Catch ex As Exception
            er = ex.Message
            Return False
        End Try
        Return hasrec
    End Function
    Public Function CountOfRecords(ByVal mySQL As String, Optional ByVal myconstr As String = "", Optional ByVal myconprv As String = "") As String
        mySQL = mySQL.Replace("""", "'")
        mySQL = cleanSQL(mySQL)
        mySQL = Replace(mySQL, Chr(34), Chr(39))
        Dim myRecords As DataTable = New DataTable
        Dim myconstring, myprovider As String
        Dim dbcase As String = String.Empty
        Try
            If myconstr.Trim = "" Then
                myconstring = System.Configuration.ConfigurationManager.ConnectionStrings.Item("mySQLconnection").ToString
                myprovider = System.Configuration.ConfigurationManager.ConnectionStrings.Item("mySQLconnection").ProviderName.ToString
                If myprovider = "Npgsql" Then dbcase = ConfigurationManager.AppSettings("ourdbcase").ToString
            Else
                If myconprv = "Npgsql" Then
                    If System.Configuration.ConfigurationManager.ConnectionStrings.Item("CSVconnection") IsNot Nothing AndAlso myconstr.Trim.ToUpper = System.Configuration.ConfigurationManager.ConnectionStrings.Item("CSVconnection").ToString.Trim.ToUpper Then
                        dbcase = ConfigurationManager.AppSettings("csvdbcase").ToString
                    ElseIf myconstr.Trim.ToUpper = System.Configuration.ConfigurationManager.ConnectionStrings.Item("mySQLconnection").ToString.Trim.ToUpper Then
                        dbcase = ConfigurationManager.AppSettings("ourdbcase").ToString
                    ElseIf myconstr.Trim.ToUpper = System.Configuration.ConfigurationManager.ConnectionStrings.Item("UserSqlConnection").ToString.Trim.ToUpper Then
                        dbcase = ConfigurationManager.AppSettings("userdbcase").ToString
                    Else 'postgres, etc...
                        dbcase = ConfigurationManager.AppSettings("ourdbcase").ToString
                    End If
                End If
                myconstring = myconstr
                myprovider = myconprv
            End If

            If myprovider = "InterSystems.Data.IRISClient" Then
                mySQL = CorrectSQLforCache(mySQL)
                Dim propResult As String = CountOfRecords_IRIS(mySQL, myconstring, myRecords)
                If propResult <> String.Empty Then Return propResult

            ElseIf myprovider = "InterSystems.Data.CacheClient" Then
                mySQL = CorrectSQLforCache(mySQL)
                Dim propResult As String = CountOfRecords_Cache(mySQL, myconstring, myRecords)
                If propResult <> String.Empty Then Return propResult

            ElseIf myprovider = "Npgsql" Then  'PostgreSQL  Npgsql
                myconstring = CorrectConnstringForPostgres(myconstring, dbcase)
                mySQL = ConvertFromSqlServerToPostgres(mySQL, dbcase, myconstring, myprovider)
                Dim myConnection As Npgsql.NpgsqlConnection
                Dim myCommand As New Npgsql.NpgsqlCommand
                myConnection = New Npgsql.NpgsqlConnection(myconstring)
                myCommand.Connection = myConnection
                myCommand.CommandType = CommandType.Text
                myCommand.CommandTimeout = 300000
                myCommand.CommandText = mySQL
                If myCommand.Connection.State = ConnectionState.Closed Then myCommand.Connection.Open()
                Dim myAdapter As New Npgsql.NpgsqlDataAdapter(myCommand)
                myAdapter.Fill(myRecords)
                myAdapter.Dispose()
                myCommand.Connection.Close()
                myCommand.Dispose()
            ElseIf myprovider = "System.Data.SqlClient" Then
                Dim myConnection As SqlConnection
                Dim myCommand As New SqlClient.SqlCommand
                Dim myAdapter As SqlClient.SqlDataAdapter
                myConnection = New SqlConnection(myconstring)
                myCommand.Connection = myConnection
                myCommand.CommandType = CommandType.Text
                myCommand.CommandTimeout = 300000
                myCommand.CommandText = mySQL
                If myCommand.Connection.State = ConnectionState.Closed Then myCommand.Connection.Open()
                myAdapter = New SqlClient.SqlDataAdapter(myCommand)
                myAdapter.Fill(myRecords)
                myAdapter.Dispose()
                myCommand.Connection.Close()
                myCommand.Dispose()
            ElseIf myprovider = "MySql.Data.MySqlClient" Then
                mySQL = CorrectSQLforMySql(mySQL, myconstring)
                Dim myConnection As MySqlConnection
                Dim myCommand As New MySqlCommand
                Dim myAdapter As MySqlDataAdapter
                myConnection = New MySqlConnection(myconstring)
                myCommand.Connection = myConnection
                myCommand.CommandType = CommandType.Text
                myCommand.CommandTimeout = 300000
                myCommand.CommandText = mySQL
                If myCommand.Connection.State = ConnectionState.Closed Then myCommand.Connection.Open()
                myAdapter = New MySqlDataAdapter(myCommand)
                myAdapter.Fill(myRecords)
                myAdapter.Dispose()
                myCommand.Connection.Close()
                myCommand.Dispose()
            ElseIf myprovider = "Oracle.ManagedDataAccess.Client" Then
                mySQL = CorrectSQLforOracle(mySQL, myconstring)
                Dim propResult As String = CountOfRecords_Oracle(mySQL, myconstring, myRecords)
                If myRecords.Rows.Count > 0 Then
                    Return myRecords.Rows.Count.ToString
                End If
            Else
                Return 0.ToString
            End If
        Catch ex As Exception
            Return "ERROR!! " & ex.Message
        End Try
        Return 0.ToString
    End Function
    Public Function CorrectSQLforOracle(sql As String, ByRef myconstring As String, Optional ByVal donotreplace As String = "") As String
        sql = sql.Replace("OUR.HelpDesk", "OURHELPDESK")
        sql = sql.Replace("OUR.AccessLog", "OURACCESSLOG")
        sql = sql.Replace("OUR.Dashboards", "OURDASHBOARDS")
        sql = sql.Replace("OUR.KMLHistory", "OURKMLHISTORY")
        sql = sql.Replace("OUR.kmlhistory", "OURKMLHISTORY")
        sql = sql.Replace("OUR.Permits", "OURPERMITS")
        sql = sql.Replace("OUR.Permissions", "OURPERMISSIONS")
        sql = sql.Replace("OUR.ReportInfo", "OURREPORTINFO")
        sql = sql.Replace("OUR.ReportShow", "OURREPORTSHOW")
        sql = sql.Replace("OUR.ReportSQLquery", "OURREPORTSQLQUERY")
        sql = sql.Replace("OUR.ReportGroups", "OURREPORTGROUPS")
        sql = sql.Replace("OUR.ReportFormat", "OURREPORTFORMAT")
        sql = sql.Replace("OUR.ReportLists", "OURREPORTLISTS")
        sql = sql.Replace("OUR.ReportChildren", "OURREPORTCHILDREN")
        sql = sql.Replace("OUR.ReportView", "OURREPORTVIEW")
        sql = sql.Replace("OUR.ReportItems", "OURREPORTITEMS")
        sql = sql.Replace("OUR.FriendlyNames", "OURFRIENDLYNAMES")
        sql = sql.Replace("OUR.Files", "OURFILES")
        sql = sql.Replace("ourfiles", "OURFILES")
        sql = sql.Replace("ourfriendlynames", "OURFRIENDLYNAMES")
        sql = sql.Replace("ourhelpdesk", "OURHELPDESK")
        sql = sql.Replace("ouraccesslog", "OURACCESSLOG")
        sql = sql.Replace("ourdashboards", "OURDASHBOARDS")
        sql = sql.Replace("ourkmlhistory", "OURKMLHISTORY")
        sql = sql.Replace("ourpermits", "OURPERMITS")
        sql = sql.Replace("ourpermissions", "OURPERMISSIONS")
        sql = sql.Replace("ourreportinfo", "OURREPORTINFO")
        sql = sql.Replace("ourreportshow", "OURREPORTSHOW")
        sql = sql.Replace("ourreportsqlquery", "OURREPORTSQLQUERY")
        sql = sql.Replace("ourreportgroups", "OURREPORTGROUPS")
        sql = sql.Replace("ourreportformat", "OURREPORTFORMAT")
        sql = sql.Replace("ourreportlists", "OURREPORTLISTS")
        sql = sql.Replace("ourreportchildren", "OURREPORTCHILDREN")
        sql = sql.Replace("ourreportview", "OURREPORTVIEW")
        sql = sql.Replace("ourreportitems", "OURREPORTITEMS")
        sql = sql.Replace("OURHelpDesk", "OURHELPDESK")
        sql = sql.Replace("OURAccessLog", "OURACCESSLOG")
        sql = sql.Replace("OURDashboards", "OURDASHBOARDS")
        sql = sql.Replace("OURKMLHistory", "OURKMLHISTORY")
        sql = sql.Replace("OURPermits", "OURPERMITS")
        sql = sql.Replace("OURPermissions", "OURPERMISSIONS")
        sql = sql.Replace("OURReportInfo", "OURREPORTINFO")
        sql = sql.Replace("OURReportShow", "OURREPORTSHOW")
        sql = sql.Replace("OURReportSQLquery", "OURREPORTSQLQUERY")
        sql = sql.Replace("OURReportGroups", "OURREPORTGROUPS")
        sql = sql.Replace("OURReportFormat", "OURREPORTFORMAT")
        sql = sql.Replace("OURReportLists", "OURREPORTLISTS")
        sql = sql.Replace("OURReportChildren", "OURREPORTCHILDREN")
        sql = sql.Replace("OURReportView", "OURREPORTVIEW")
        sql = sql.Replace("OURReportItems", "OURREPORTITEMS")
        sql = sql.Replace("OURFriendlyNames", "OURFRIENDLYNAMES")
        sql = sql.Replace("OURFiles", "OURFILES")
        sql = sql.Replace("ourtasklistsetting", "OURTASKLISTSETTING")
        sql = sql.Replace("OurTaskListSetting", "OURTASKLISTSETTING")
        sql = sql.Replace("OUR.TaskListSetting", "OURTASKLISTSETTING")
        'ScheduledReports
        sql = sql.Replace("ourscheduledreports", "OURSCHEDULEDREPORTS")
        sql = sql.Replace("OurScheduledReports", "OURSCHEDULEDREPORTS")
        sql = sql.Replace("OUR.ScheduledReports", "OURSCHEDULEDREPORTS")
        'OURComparison
        sql = sql.Replace("OURComparison", "OURCOMPARISON")
        sql = sql.Replace("OurComparison", "OURCOMPARISON")
        sql = sql.Replace("OURcomparison", "OURCOMPARISON")
        sql = sql.Replace("ourcomparison", "OURCOMPARISON")
        sql = sql.Replace("[[", "[").Replace("]]", "]").Replace("``", "`")
        sql = sql.Replace("`", "")

        sql = sql.Replace("[Order]", """Order""").Replace("[Group]", """Group""").Replace("[User]", """User""").Replace("[Count]", """Count""").Replace("[Access]", """Access""")
        sql = sql.Replace("`Order`", """Order""").Replace("`Group`", """Group""").Replace("`User`", """User""").Replace("`Count`", """Count""").Replace("`Access`", """Access""")
        'replace [] aroung reserved word to ""
        'If donotreplace = "" Then sql = sql.Replace("[", """").Replace("]", """")
        If donotreplace = "" Then sql = sql.Replace("[", "").Replace("]", "")
        Return sql
    End Function

    Public Function mRecords(ByVal mySQL As String, Optional ByRef er As String = "", Optional ByVal myconstr As String = "", Optional ByVal myconprv As String = "") As DataView
        Dim myView As DataView = New DataView
        Dim myRecords As DataTable = New DataTable
        Dim myconstring, myprovider As String
        Dim dbcase As String = String.Empty
        Try
            If myconstr.Trim = "" Then
                myconstring = System.Configuration.ConfigurationManager.ConnectionStrings.Item("mySQLconnection").ToString
                myprovider = System.Configuration.ConfigurationManager.ConnectionStrings.Item("mySQLconnection").ProviderName.ToString
                If myprovider = "Npgsql" Then dbcase = ConfigurationManager.AppSettings("ourdbcase").ToString
            Else
                If myconprv = "Npgsql" Then
                    If System.Configuration.ConfigurationManager.ConnectionStrings.Item("CSVconnection") IsNot Nothing AndAlso myconstr.Trim.ToUpper = System.Configuration.ConfigurationManager.ConnectionStrings.Item("CSVconnection").ToString.Trim.ToUpper Then
                        dbcase = ConfigurationManager.AppSettings("csvdbcase").ToString
                    ElseIf myconstr.Trim.ToUpper = System.Configuration.ConfigurationManager.ConnectionStrings.Item("mySQLconnection").ToString.Trim.ToUpper Then
                        dbcase = ConfigurationManager.AppSettings("ourdbcase").ToString
                    ElseIf myconstr.Trim.ToUpper = System.Configuration.ConfigurationManager.ConnectionStrings.Item("UserSqlConnection").ToString.Trim.ToUpper Then
                        dbcase = ConfigurationManager.AppSettings("userdbcase").ToString
                    Else 'postgres, etc...
                        dbcase = ConfigurationManager.AppSettings("ourdbcase").ToString
                    End If
                End If
                myconstring = myconstr
                myprovider = myconprv
            End If

            If myprovider = "InterSystems.Data.IRISClient" Then
                mySQL = CorrectSQLforCache(mySQL)
                Dim propResult As String = mRecords_IRIS(mySQL, er, myconstring, myRecords)
                If propResult = "USE_DISTINCT" Then
                    er = GetDistinctDataViewWithProperCase(mySQL, myView, myconstring, myprovider)
                    Return myView
                ElseIf propResult = "RETURN_VIEW" Then
                    Return myView
                End If

            ElseIf myprovider = "InterSystems.Data.CacheClient" Then
                mySQL = CorrectSQLforCache(mySQL)
                Dim propResult As String = mRecords_Cache(mySQL, er, myconstring, myRecords)
                If propResult = "USE_DISTINCT" Then
                    er = GetDistinctDataViewWithProperCase(mySQL, myView, myconstring, myprovider)
                    Return myView
                ElseIf propResult = "RETURN_VIEW" Then
                    Return myView
                End If

            ElseIf myprovider = "MySql.Data.MySqlClient" Then
                Dim myConnection As MySqlConnection
                Dim myCommand As New MySqlCommand
                Dim myAdapter As MySqlDataAdapter
                Try
                    mySQL = CorrectSQLforMySql(mySQL, myconstring)
                    myConnection = New MySqlConnection(myconstring)
                    myCommand.Connection = myConnection
                    myCommand.CommandType = CommandType.Text
                    myCommand.CommandTimeout = 300000
                    myCommand.CommandText = mySQL
                    If myCommand.Connection.State = ConnectionState.Closed Then myCommand.Connection.Open()
                    myAdapter = New MySqlDataAdapter(myCommand)
                    myAdapter.Fill(myRecords)
                    myAdapter.Dispose()
                Catch ex As Exception
                    'assign error er
                    er = "ERROR!! " & ex.Message
                End Try
                myCommand.Connection.Close()
                myCommand.Dispose()

            ElseIf myprovider = "Oracle.ManagedDataAccess.Client" Then
                mySQL = CorrectSQLforOracle(mySQL, myconstring)
                mRecords_Oracle(mySQL, er, myconstring, myRecords)

            ElseIf myprovider = "Npgsql" Then  'PostgreSQL  Npgsql
                Dim myConnection As Npgsql.NpgsqlConnection
                Dim myCommand As New Npgsql.NpgsqlCommand
                Dim myAdapter As New Npgsql.NpgsqlDataAdapter
                Try
                    myconstring = CorrectConnstringForPostgres(myconstring, dbcase)
                    mySQL = ConvertFromSqlServerToPostgres(mySQL, dbcase, myconstring, myprovider)
                    myConnection = New Npgsql.NpgsqlConnection(myconstring)
                    myCommand.Connection = myConnection
                    myCommand.CommandType = CommandType.Text
                    myCommand.CommandTimeout = 300000
                    myCommand.CommandText = mySQL
                    If myCommand.Connection.State = ConnectionState.Closed Then myCommand.Connection.Open()
                    myAdapter = New Npgsql.NpgsqlDataAdapter(myCommand)
                    myAdapter.Fill(myRecords)
                Catch ex As Exception
                    'assign error er
                    er = "ERROR!! " & ex.Message
                End Try
                'myAdapter.Dispose()
                myCommand.Connection.Close()
                myCommand.Dispose()

            Else
                Dim myConnection As SqlConnection
                Dim myCommand As New SqlClient.SqlCommand
                Dim myAdapter As SqlClient.SqlDataAdapter
                Try
                    mySQL = CorrectSQLforSQLServer(mySQL)
                    mySQL = mySQL.Replace("""", "'")
                    mySQL = mySQL.Replace("`", "")
                    myConnection = New SqlConnection(myconstring)
                    myCommand.Connection = myConnection
                    myCommand.CommandType = CommandType.Text
                    myCommand.CommandTimeout = 300000
                    myCommand.CommandText = mySQL
                    If myCommand.Connection.State = ConnectionState.Closed Then myCommand.Connection.Open()
                    myAdapter = New SqlClient.SqlDataAdapter(myCommand)
                    myAdapter.Fill(myRecords)
                    myAdapter.Dispose()
                Catch ex As Exception
                    'assign error er
                    er = "ERROR!! " & ex.Message
                End Try
                myCommand.Connection.Close()
                myCommand.Dispose()
            End If
            myView = myRecords.DefaultView

            'Dim dt As DataTable = MakeDTColumnsNamesCLScompliant(myRecords, myprovider, er)
            'myView = dt.DefaultView
            ''fix upper case in InterSystems
            'If myprovider.StartsWith("InterSystems.Data.") AndAlso mySQL.ToUpper.IndexOf(" DISTINCT ") > 0 Then
            '    err = GetDistinctDataViewWithProperCase(mySQL, myView, myconstring, myprovider)
            'End If

        Catch ex As Exception
            'return error somehow
            er = "ERROR!! " & ex.Message
        End Try
        Return myView
    End Function
    Public Function GetDistinctDataViewWithProperCase(ByVal sqlDistinct As String, ByRef dtdist As DataView, Optional myconnstring As String = "", Optional myprovider As String = "") As String
        Dim er As String = String.Empty
        If sqlDistinct.ToUpper.IndexOf(" DISTINCT ") < 0 Then
            Return er
        End If
        Dim col As String = String.Empty
        Dim sqls As String = String.Empty
        Dim FLD As String = String.Empty
        Dim fild As String = String.Empty
        Dim dt As DataView
        If myprovider.Trim = "" Then
            myprovider = System.Configuration.ConfigurationManager.ConnectionStrings.Item("mySQLconnection").ProviderName.ToString
            myconnstring = System.Configuration.ConfigurationManager.ConnectionStrings.Item("mySQLconnection").ToString
        End If
        Try
            sqls = sqlDistinct.Replace("DISTINCT ", "").Replace("distinct ", "")
            If sqls.ToUpper.IndexOf(" DISTINCT ") > 0 Then
                sqls = sqls.Substring(0, sqls.ToUpper.IndexOf(" DISTINCT ")) & " " & sqls.Substring(sqls.ToUpper.IndexOf(" DISTINCT ") + 9)
            End If
            dt = mRecords(sqls, er, myconnstring, myprovider)
            Dim dattbl As DataTable = dt.ToTable(True)
            dtdist = dattbl.DefaultView
            Return er
            'dt.Table.CaseSensitive = False
            'If dtdist IsNot Nothing AndAlso dtdist.Table.Rows.Count > 0 Then
            '    If myprovider.StartsWith("InterSystems.Data.") Then
            '        'DISTINCT returns upper case in string fields, fixing dtdist
            '        For j = 0 To dtdist.Table.Columns.Count - 1
            '            If Not ColumnTypeIsNumeric(dtdist.Table.Columns(j)) Then
            '                For i = 0 To dtdist.Table.Rows.Count - 1
            '                    col = dtdist.Table.Columns(j).Caption
            '                    FLD = dtdist.Table.Rows(i)(j)  'upper case
            '                    'dt.RowFilter = "(UCASE(er))=" & "'" & FLD & "'"    'cannot recognize UCASE in InterSystems
            '                    'If Not dt.ToTable Is Nothing AndAlso dt.ToTable.Rows.Count > 0 Then
            '                    '    dtdist.Table.Rows(i)(j) = dt.ToTable.Rows(0)(j)
            '                    'End If
            '                    'dt.RowFilter = ""
            '                    sqls = sqlDistinct.Replace(" DISTINCT ", " TOP 1 ").Replace(" distinct ", " TOP 1 ")
            '                    If sqls.ToUpper.IndexOf(" DISTINCT ") > 0 Then 'still has Distinct in mix case
            '                        sqls = sqlDistinct.Substring(0, sqlDistinct.ToUpper.IndexOf(" DISTINCT ")) & " TOP 1 " & sqlDistinct.Substring(sqlDistinct.ToUpper.IndexOf(" DISTINCT ") + 9)
            '                    End If
            '                    If sqls.ToUpper.IndexOf(" ORDER BY ") > 0 Then
            '                        sqls = sqls.Substring(0, sqls.ToUpper.IndexOf(" ORDER BY "))
            '                    End If
            '                    If sqls.ToUpper.IndexOf(" GROUP BY ") > 0 Then
            '                        sqls = sqls.Substring(0, sqls.ToUpper.IndexOf(" GROUP BY "))
            '                    End If
            '                    If sqls.ToUpper.IndexOf(" HAVING") > 0 Then
            '                        sqls = sqls.Substring(0, sqls.ToUpper.IndexOf(" HAVING"))
            '                    End If
            '                    If sqls.ToUpper.IndexOf(" WHERE ") > 0 Then
            '                        sqls = sqls.Substring(0, sqls.ToUpper.IndexOf(" WHERE ") + 7) & " UCASE(" & col & ")='" & FLD.ToUpper & "' AND " & sqls.Substring(sqls.ToUpper.IndexOf(" WHERE ") + 7)
            '                    Else
            '                        sqls = sqls & " WHERE (UCASE(" & col & ")='" & FLD.ToUpper & "')"
            '                    End If
            '                    dt = mRecords(sqls, er, myconnstring, myprovider)
            '                    If er.Trim = "" AndAlso Not dt.ToTable Is Nothing AndAlso dt.ToTable.Rows.Count > 0 Then
            '                        dtdist.Table.Rows(i)(j) = dt.Table.Rows(0)(j)
            '                    End If
            '                Next
            '            End If
            '        Next
            '    End If
            '    Return er
            'End If
        Catch ex As Exception
            er = "ERROR!! " & ex.Message
            Return Nothing
        End Try
        Return Nothing
    End Function
    Public Function ConvertFromSqlServerToPostgres(ByVal mSql As String, ByVal dbcase As String, Optional ByVal myconstring As String = "", Optional ByVal myprovider As String = "") As String
        Dim er As String = ""
        Dim Sql As String = mSql.Replace("[[", "[").Replace("]]", "]").Replace("``", "`")
        Try
            Sql = Sql.Replace("[Order]", """Order""").Replace("[Group]", """Group""").Replace("[User]", """User""").Replace("[Count]", """Count""").Replace("[Status]", """Status""")
            Sql = Sql.Replace("`Order`", """Order""").Replace("`Group`", """Group""").Replace("`User`", """User""").Replace("`Count`", """Count""").Replace("`Status`", """Status""")
            If dbcase = "doublequoted" Then
                Sql = Sql.Replace("[", """").Replace("]", """").Replace("`", """")

                Sql = Sql.Replace("OUR.HelpDesk", "OURHelpDesk")
                Sql = Sql.Replace("OUR.AccessLog", "OURAccessLog")
                Sql = Sql.Replace("OUR.Dashboards", "OURDashboards")
                Sql = Sql.Replace("OUR.Permits", "OURPermits")
                Sql = Sql.Replace("OUR.Permissions", "OURPermissions")
                Sql = Sql.Replace("OUR.ReportInfo", "OURReportInfo")
                Sql = Sql.Replace("OUR.ReportShow", "OURReportShow")
                Sql = Sql.Replace("OUR.ReportSQLquery", "OURReportSQLquery")
                Sql = Sql.Replace("OUR.ReportGroups", "OURReportGroups")
                Sql = Sql.Replace("OUR.ReportFormat", "OURReportFormat")
                Sql = Sql.Replace("OUR.ReportLists", "OURReportLists")
                Sql = Sql.Replace("OUR.ReportChildren", "OURReportChildren")
                Sql = Sql.Replace("OUR.FriendlyNames", "OURFriendlyNames")
                Sql = Sql.Replace("OUR.Files", "OURFiles")
                Sql = Sql.Replace("OURFILES", "OURFiles")
                Sql = Sql.Replace("OURFRIENDLYNAMES", "OURFriendlyNames")
                Sql = Sql.Replace("OURHELPDESK", "OURHelpDesk")
                Sql = Sql.Replace("OURACCESSLOG", "OURAccessLog")
                Sql = Sql.Replace("OURDASHBOARDS", "OURDashboards")
                Sql = Sql.Replace("OURPERMITS", "OURPermits")
                Sql = Sql.Replace("OURPERMISSIONS", "OURPermissions")
                Sql = Sql.Replace("OURREPORTINFO", "OURReportInfo")
                Sql = Sql.Replace("OURREPORTSHOW", "OURReportShow")
                Sql = Sql.Replace("OURREPORTSQLQUERY", "OURReportSQLquery")
                Sql = Sql.Replace("OURREPORTGROUPS", "OURReportGroups")
                Sql = Sql.Replace("OURREPORTFORMAT", "OURReportFormat")
                Sql = Sql.Replace("OURREPORTLISTS", "OURReportLists")
                Sql = Sql.Replace("OURREPORTCHILDREN", "OURReportChildren")
                Sql = Sql.Replace("ourfiles", "OURFiles")
                Sql = Sql.Replace("ourfriendlynames", "OURFriendlyNames")
                Sql = Sql.Replace("ourhelpdesk", "OURHelpDesk")
                Sql = Sql.Replace("ouraccesslog", "OURAccessLog")
                Sql = Sql.Replace("ourdashboards", "OURDashboards")
                Sql = Sql.Replace("ourpermits", "OURPermits")
                Sql = Sql.Replace("ourpermissions", "OURPermissions")
                Sql = Sql.Replace("ourreportinfo", "OURReportInfo")
                Sql = Sql.Replace("ourreportshow", "OURReportShow")
                Sql = Sql.Replace("ourreportsqlquery", "OURReportSQLquery")
                Sql = Sql.Replace("ourreportgroups", "OURReportGroups")
                Sql = Sql.Replace("ourreportformat", "OURReportFormat")
                Sql = Sql.Replace("ourreportlists", "OURReportLists")
                Sql = Sql.Replace("ourreportchildren", "OURReportChildren")
                Sql = Sql.Replace("OUR.ScheduledReports", "OURScheduledReports")
                Sql = Sql.Replace("ourscheduledreports", "OURScheduledReports")
                Sql = Sql.Replace("OURSCHEDULEDREPORTS", "OURScheduledReports")
                Sql = Sql.Replace("ourtasklistsetting", "OurTaskListSetting")
                Sql = Sql.Replace("OURTASKLISTSETTING", "OurTaskListSetting")
                Sql = Sql.Replace("OUR.TaskListSetting", "OurTaskListSetting")
                Sql = Sql.Replace("ourkmlhistory", "OURKMLHistory")
                Sql = Sql.Replace("OURKMLHISTORY", "OURKMLHistory")
                Sql = Sql.Replace("OUR.KMLHistory", "OURKMLHistory")

                'OURComparison
                Sql = Sql.Replace("OurComparison", "OURComparison")
                Sql = Sql.Replace("OURCOMPARISON", "OURComparison")
                Sql = Sql.Replace("OURcomparison", "OURComparison")
                Sql = Sql.Replace("ourcomparison", "OURComparison")

                Sql = Sql.Replace("""""", """").Replace("""""", """")
                If Not Sql.ToUpper.Contains("FROM INFORMATION_SCHEMA") Then
                    'correct tables and fields in Sql with double quotes
                    Sql = CorrectSQLforPostgreWithDoublequotes(Sql, dbcase, myconstring, "Npgsql")
                End If
                Sql = Sql.Replace("""""", """").Replace("""""", """")
            Else
                Sql = Sql.Replace("[", "").Replace("]", "").Replace("`", "")


                Sql = Sql.Replace("OURHelpDesk", "ourhelpdesk")
                Sql = Sql.Replace("OURAccessLog", "ouraccesslog")
                Sql = Sql.Replace("OURDashboards", "ourdashboards")
                Sql = Sql.Replace("OURKMLHistory", "ourkmlhistory")
                Sql = Sql.Replace("OURPermits", "ourpermits")
                Sql = Sql.Replace("OURPermissions", "ourpermissions")
                Sql = Sql.Replace("OURReportInfo", "ourreportinfo")
                Sql = Sql.Replace("OURReportShow", "ourreportshow")
                Sql = Sql.Replace("OURReportSQLquery", "ourreportsqlquery")
                Sql = Sql.Replace("OURReportGroups", "ourreportgroups")
                Sql = Sql.Replace("OURReportFormat", "ourreportformat")
                Sql = Sql.Replace("OURReportLists", "ourreportlists")
                Sql = Sql.Replace("OURReportChildren", "ourreportchildren")
                Sql = Sql.Replace("OURFriendlyNames", "ourfriendlynames")
                Sql = Sql.Replace("OURFiles", "ourfiles")
                Sql = Sql.Replace("OURUnits", "ourunits")
                Sql = Sql.Replace("OURUserTables", "ourusertables")
                Sql = Sql.Replace("OURUSERTABLES", "ourusertables")
                Sql = Sql.Replace("OURAgents", "ouragents")
                Sql = Sql.Replace("OURAGENTS", "ouragents")
                Sql = Sql.Replace("OURActivity", "ouractivity")
                Sql = Sql.Replace("OURACTIVITY", "ouractivity")
                Sql = Sql.Replace("OURUNITS", "ourunits")
                Sql = Sql.Replace("OURFILES", "ourfiles")
                Sql = Sql.Replace("OURFRIENDLYNAMES", "ourfriendlynames")
                Sql = Sql.Replace("OURHELPDESK", "ourhelpdesk")
                Sql = Sql.Replace("OURACCESSLOG", "ouraccessLog")
                Sql = Sql.Replace("OURDASHBOARDS", "ourdashboards")
                Sql = Sql.Replace("OURKMLHISTORY", "ourkmlhistory")
                Sql = Sql.Replace("OURPERMITS", "ourpermits")
                Sql = Sql.Replace("OURPERMISSIONS", "ourpermissions")
                Sql = Sql.Replace("OURREPORTINFO", "ourreportinfo")
                Sql = Sql.Replace("OURREPORTSHOW", "ourreportshow")
                Sql = Sql.Replace("OURREPORTSQLQUERY", "ourreportsqlquery")
                Sql = Sql.Replace("OURREPORTGROUPS", "ourreportgroups")
                Sql = Sql.Replace("OURREPORTFORMAT", "ourreportformat")
                Sql = Sql.Replace("OURREPORTLISTS", "ourreportlists")
                Sql = Sql.Replace("OURREPORTCHILDREN", "ourreportchildren")
                Sql = Sql.Replace("OURTASKLISTSETTING", "ourtasklistsetting")
                Sql = Sql.Replace("OurTaskListSetting", "ourtasklistsetting")
                Sql = Sql.Replace("OurScheduledReports", "ourscheduledreports")
                Sql = Sql.Replace("OURScheduledReports", "ourscheduledreports")
                Sql = Sql.Replace("OURSCHEDULEDREPORTS", "ourscheduledreports")
                'OURComparison
                Sql = Sql.Replace("OURComparison", "ourcomparison")
                Sql = Sql.Replace("OurComparison", "ourcomparison")
                Sql = Sql.Replace("OURcomparison", "ourcomparison")
                Sql = Sql.Replace("OURCOMPARISON", "ourcomparison")
            End If


        Catch exc As Exception
            er = exc.Message
            Sql = ""
        End Try
        Return Sql
    End Function
    Public Function CorrectSQLforPostgreWithDoublequotes(ByVal mSql As String, ByVal dbcase As String, Optional ByVal myconstring As String = "", Optional ByVal myprovider As String = "") As String
        Dim ret As String = String.Empty
        Dim er As String = String.Empty
        Dim fromstmt As String = String.Empty
        Dim fieldsarr As String = String.Empty
        Dim fldname As String = String.Empty
        Dim vals As String = String.Empty
        Dim wherestmt As String = String.Empty
        Dim i, k As Integer

        Try
            If dbcase.Trim <> "doublequoted" Then
                Return mSql
            End If
            If myconstring.Trim = "" Then
                myconstring = System.Configuration.ConfigurationManager.ConnectionStrings.Item("mySQLconnection").ToString
            End If
            If myprovider.Trim = "" Then
                myprovider = System.Configuration.ConfigurationManager.ConnectionStrings.Item("mySQLconnection").ProviderName.ToString
            End If
            If myprovider.Trim <> "Npgsql" Then
                Return mSql
            End If
            'CorrectWhereHavingFromSQLquery, CorrectFieldsArrayFromSQLquery, CorrectTablesArrayFromSQLquery
            If mSql.TrimStart.ToUpper.StartsWith("INSERT ") Then
                'has only array of fields msql = "INSERT INTO " & Tbl & " (" & mfields & ") VALUES ("
                fromstmt = mSql.Substring(11, mSql.IndexOf("(") - 11).Trim
                fieldsarr = mSql.Substring(mSql.IndexOf("("), mSql.IndexOf(")") - mSql.IndexOf("(")).Replace("(", "").Replace(")", "").Trim
                vals = mSql.Substring(mSql.IndexOf(")"))
                fieldsarr = CorrectFieldsArrayFromSQLquery("SELECT " & fieldsarr & " FROM " & fromstmt, myconstring, myprovider, er)
                ret = "INSERT INTO """ & fromstmt & """ (" & fieldsarr & vals
            ElseIf mSql.ToUpper.TrimStart.StartsWith("UPDATE ") Then
                'has sets and where
                fromstmt = mSql.Substring(7, mSql.IndexOf(" SET ")).Trim  'table name from UPDATE query
                Dim dtb As New DataTable
                dtb.TableName = "Table1"
                dtb.Columns.Add("Tbl1")
                Dim myRow As DataRow
                myRow = dtb.NewRow
                myRow.Item(0) = fromstmt
                dtb.Rows.Add(myRow)

                wherestmt = String.Empty
                If mSql.IndexOf(" WHERE ") > 0 OrElse mSql.ToUpper.IndexOf(" HAVING ") > 0 Then
                    If mSql.ToUpper.IndexOf(" WHERE ") > 0 Then
                        wherestmt = mSql.Substring(mSql.IndexOf(" WHERE ") + 6)
                    ElseIf mSql.ToUpper.IndexOf(" HAVING ") > 0 Then
                        wherestmt = mSql.Substring(mSql.IndexOf(" HAVING ") + 7)
                    End If
                    'wherestmt = CorrectWhereHavingFromSQLquery(wherestmt, dtb, myconstring, myprovider, er)
                End If
                fieldsarr = mSql.Substring(mSql.IndexOf(" SET ") + 4)
                Dim fieldsvals() As String = fieldsarr.Split(",")
                For i = 0 To fieldsvals.Length - 1
                    fldname = """" & Piece(fieldsvals(i), "=", 1).Trim & """"
                    vals = Piece(fieldsvals(i), "=", 2)
                    fieldsvals(i) = fldname & vals
                Next
                fieldsarr = ""
                For i = 0 To fieldsvals.Length - 1
                    fieldsarr = fieldsarr & fieldsvals(i)
                    If i < fieldsvals.Length - 1 Then
                        fieldsarr = fieldsarr & ", "
                    End If
                Next
                ret = "UPDATE """ & fromstmt & """ SET " & fieldsarr
                If wherestmt.Trim <> "" Then
                    ret = ret & " WHERE " & wherestmt
                End If

            ElseIf mSql.ToUpper.TrimStart.StartsWith("SELECT ") Then
                'has arrays of fields for selected fields, order by, group by and where or having statement
                'Dim dt As DataTable = GetListOfTablesFromSQLquery(mSql, myconstring, myprovider, er)
                'If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                '    For i = 0 To dt.Rows.Count - 1
                '        If Not dt.Rows(i)("Tbl1").ToString.Trim.StartsWith("""") OrElse Not dt.Rows(i)("Tbl1").ToString.Trim.EndsWith("""") Then
                '            mSql = ReplaceWholeWord(mSql, dt.Rows(i)("Tbl1"), """" & dt.Rows(i)("Tbl1").trim & """")
                '            mSql = mSql.Replace("""""", """")
                '        End If
                '    Next
                'End If

                'Fix JOINS ON statements
                mSql = CorrectTablesJoinsFromSQLquery(mSql, myconstring, myprovider, er)
                fieldsarr = CorrectFieldsArrayFromSQLquery(mSql, myconstring, myprovider, er)
                wherestmt = String.Empty
                If mSql.IndexOf(" WHERE ") > 0 OrElse mSql.ToUpper.IndexOf(" HAVING ") > 0 Then
                    If mSql.ToUpper.IndexOf(" WHERE ") > 0 Then
                        wherestmt = mSql.Substring(mSql.IndexOf(" WHERE ") + 7)
                    ElseIf mSql.ToUpper.IndexOf(" HAVING ") > 0 Then
                        wherestmt = mSql.Substring(mSql.IndexOf(" HAVING ") + 8)
                    End If
                    'wherestmt = CorrectWhereHavingFromSQLquery(wherestmt, dt, myconstring, myprovider, er)
                End If
                'group by fields
                Dim groupbyflds As String = String.Empty
                If mSql.ToUpper.IndexOf("GROUP BY") > 0 Then
                    groupbyflds = mSql.Substring(mSql.ToUpper.IndexOf("GROUP BY") + 8)

                    k = groupbyflds.ToUpper.IndexOf(" ORDER BY")
                    If k > 0 Then
                        groupbyflds = groupbyflds.Substring(0, k)
                    End If
                    k = groupbyflds.ToUpper.IndexOf(" WHERE ")
                    If k > 0 Then
                        groupbyflds = groupbyflds.Substring(0, k)
                    End If
                    k = groupbyflds.ToUpper.IndexOf(" HAVING ")
                    If k > 0 Then
                        groupbyflds = groupbyflds.Substring(0, k)
                    End If
                    groupbyflds = CorrectFieldsArrayFromSQLquery(groupbyflds, myconstring, myprovider, er)
                End If
                'order by fields
                Dim orderbyflds As String = String.Empty
                If mSql.ToUpper.IndexOf("ORDER BY") > 0 Then
                    orderbyflds = mSql.Substring(mSql.ToUpper.IndexOf("ORDER BY") + 8).Trim
                    Dim ord As String = orderbyflds.Substring(orderbyflds.LastIndexOf(" "))
                    If ord.Trim.ToUpper = "ASC" OrElse ord.Trim.ToUpper = "DESC" Then
                        orderbyflds = orderbyflds.Substring(0, orderbyflds.LastIndexOf(" ")).Trim
                        orderbyflds = CorrectFieldsArrayFromSQLquery(orderbyflds, myconstring, myprovider, er)
                        orderbyflds = orderbyflds & " " & ord
                    Else
                        orderbyflds = CorrectFieldsArrayFromSQLquery(orderbyflds, myconstring, myprovider, er)
                    End If
                End If
                fromstmt = mSql.Substring(mSql.IndexOf(" FROM ") + 6).Trim
                k = fromstmt.ToUpper.IndexOf(" WHERE ")
                If k > 0 Then
                    fromstmt = fromstmt.Substring(0, k)
                End If
                k = fromstmt.ToUpper.IndexOf(" HAVING ")
                If k > 0 Then
                    fromstmt = fromstmt.Substring(0, k)
                End If
                k = fromstmt.ToUpper.IndexOf(" GROUP BY ")
                If k > 0 Then
                    fromstmt = fromstmt.Substring(0, k)
                End If
                k = fromstmt.ToUpper.IndexOf(" ORDER BY ")
                If k > 0 Then
                    fromstmt = fromstmt.Substring(0, k)
                End If

                If mSql.ToUpper.StartsWith("SELECT ") Then
                    ret = "SELECT "
                End If
                If mSql.ToUpper.StartsWith("SELECT DISTINCT ") Then
                    ret = "SELECT DISTINCT "
                End If
                ret = ret & fieldsarr & " FROM " & fromstmt

                If groupbyflds.Trim <> "" Then
                    ret = ret & " GROUP BY " & groupbyflds
                End If

                If wherestmt.Trim <> "" Then
                    ret = ret & " WHERE " & wherestmt
                End If

                If orderbyflds.Trim <> "" Then
                    ret = ret & " ORDER BY " & orderbyflds
                End If
            End If

        Catch ex As Exception
            ret = "ERROR!! " & ex.Message
        End Try
        Return ret
    End Function
    Public Function FixTableAndField(ByVal dtb As DataTable, ByRef tbl As String, ByRef fld As String, Optional ByVal userconstr As String = "", Optional ByVal userconprv As String = "", Optional ByVal er As String = "") As String
        'dtb is list of report tables
        Dim ret As String = "Field Not found: "
        If fld = "" Then
            Return ""
        End If
        tbl = tbl.Replace("`", "")
        tbl = tbl.Replace("[", "").Replace("]", "").Trim
        tbl = tbl.Replace("""", "")
        tbl = tbl.Replace("(", "").Replace(")", "").Trim
        fld = fld.Replace("=", "")
        fld = fld.Replace("`", "")
        fld = fld.Replace("[", "").Replace("]", "").Trim
        fld = fld.Replace("""", "")
        fld = fld.Replace("(", "").Replace(")", "").Trim

        Dim sTable As String = tbl.ToUpper
        If sTable = "DATE1" OrElse sTable = "DATE2" OrElse sTable = "VALUE1" OrElse sTable = "VALUE2" Then _
            Return ""

        Dim j, n As Integer
        n = fld.LastIndexOf(".")
        If n > 0 Then
            'table and field
            tbl = fld.Substring(0, n)
            fld = fld.Substring(n + 1)
            ret = "Field found: "
        ElseIf n < 0 AndAlso tbl <> "" Then
            If Not IsColumnFromTable(tbl, fld, userconstr, userconprv, er) Then
                ' to find the table
                For j = 0 To dtb.Rows.Count - 1
                    tbl = dtb.Rows(j)("Tbl1")
                    If IsColumnFromTable(tbl, fld, userconstr, userconprv, er) Then
                        ret = "Field found: "
                        Exit For
                    End If
                Next
            End If
        Else
            'fld itself, to find the table
            For j = 0 To dtb.Rows.Count - 1
                tbl = dtb.Rows(j)("Tbl1")
                If IsColumnFromTable(tbl, fld, userconstr, userconprv, er) Then
                    ret = "Field found: "
                    Exit For
                End If
            Next
        End If
        If ret = "Field found: " Then
            'table and field names fix and put them back together separated with dot
            tbl = FixTableName(tbl, userconstr, userconprv, er)
            fld = FixFieldName(dtb, fld, tbl, userconstr, userconprv, er)
        Else
            Return ""
        End If
        Return tbl & "." & fld
    End Function
    Public Function IsColumnFromTable(ByVal tbl As String, ByVal fld As String, Optional ByVal userconstr As String = "", Optional ByVal userconprv As String = "", Optional ByRef err As String = "") As Boolean
        Dim re As Boolean = False
        Dim i As Integer
        Dim dv As DataView
        Try
            dv = GetListOfTableColumns(tbl, userconstr, userconprv, err)
            If dv Is Nothing OrElse dv.Count = 0 OrElse dv.Table.Rows.Count = 0 Then
                Return False
            Else
                For i = 0 To dv.Table.Rows.Count - 1
                    If dv.Table.Rows(i)("COLUMN_NAME").ToString.ToUpper = fld.ToUpper Then
                        Return True
                    End If
                Next
            End If
        Catch ex As Exception
            err = ex.Message
            Return False
        End Try
        Return False
    End Function
    'Public Function CorrectWhereHavingFromSQLquery(ByVal wheresql As String, ByVal dtb As DataTable, Optional ByVal myconstr As String = "", Optional ByVal myconprv As String = "", Optional ByRef er As String = "") As String
    '    Dim ret As String = String.Empty
    '    Dim i, j, n, m, k, l As Integer
    '    Dim g As Integer = 0 'group number
    '    If wheresql = "" OrElse dtb Is Nothing OrElse dtb.Rows.Count = 0 Then
    '        Return wheresql
    '    End If
    '    Dim sql As String = wheresql
    '    Dim dt As New DataTable
    '    If sql.ToUpper.Trim = "WHERE" Then
    '        k = sql.ToUpper.IndexOf(" WHERE ")
    '        sql = sql.Substring(k + 6).Trim
    '        k = sql.ToUpper.IndexOf("HAVING")
    '        If k > 0 Then
    '            sql = sql.Substring(0, k)
    '        End If
    '    ElseIf sql.ToUpper.Trim = "HAVING" Then
    '        k = sql.ToUpper.IndexOf(" HAVING ")
    '        sql = sql.Substring(k + 7).Trim
    '        k = sql.ToUpper.IndexOf("WHERE")
    '        If k > 0 Then
    '            sql = sql.Substring(0, k)
    '        End If
    '    End If
    '    k = sql.ToUpper.IndexOf(" ORDER BY ")
    '    If k > 0 Then
    '        sql = sql.Substring(0, k).Trim
    '    End If
    '    k = sql.ToUpper.IndexOf("GROUP BY")
    '    If k > 0 Then
    '        sql = sql.Substring(0, k)
    '    End If

    '    If sql.ToUpper.Contains("BETWEEN") Then
    '        sql = AdjustBetweens(sql)
    '    End If

    '    'at this point sql is set of conditions connected with AND or OR
    '    Dim fldfullname As String = String.Empty
    '    Dim tbl1 As String = String.Empty
    '    Dim fld1 As String = String.Empty
    '    Dim tbl2 As String = String.Empty
    '    Dim fld2 As String = String.Empty
    '    Dim tbl3 As String = String.Empty
    '    Dim fld3 As String = String.Empty
    '    Dim logical As String = "And"
    '    Dim group As String = String.Empty
    '    Dim cond As String = String.Empty
    '    Dim opr As String = String.Empty
    '    Dim notopr As String = String.Empty
    '    Dim typ As String = String.Empty
    '    Dim sta As String = String.Empty
    '    Dim lside As String = String.Empty
    '    Dim rside As String = String.Empty
    '    Dim opers(), notopers() As String
    '    Dim oper As String = String.Empty
    '    'opr = ",=,<>,<,>,<=,>=, not between , between, StartsWith , Not StartsWith , Contains , Not Contains , In , Not In , EndsWith , Not EndsWith ,"
    '    opr = ",=,<>,<,>,<=,>=,Not between,between,not like,like,Not In, In ,"
    '    notopr = ",<>,=,>=,<=,>,<, Not between, between , Not StartsWith , StartsWith , Not Contains , Contains , Not In , In , Not EndsWith , EndsWith ,"  'Not oper
    '    opers = opr.Split(",")
    '    n = opers.Length
    '    notopers = notopr.Split(",")  'length is the same
    '    opr = ""
    '    notopr = ""
    '    Try
    '        dt.TableName = "Table1"
    '        dt.Columns.Add("Tbl1")
    '        dt.Columns.Add("Tbl1Fld1")
    '        dt.Columns.Add("Tbl2")
    '        dt.Columns.Add("Tbl2Fld2")
    '        dt.Columns.Add("Tbl3")
    '        dt.Columns.Add("Tbl3Fld3")
    '        dt.Columns.Add("Oper")
    '        dt.Columns.Add("Type")
    '        dt.Columns.Add("comments")
    '        dt.Columns.Add("RecOrder")
    '        dt.Columns.Add("Group")
    '        dt.Columns.Add("Logical")

    '        Dim myRow As DataRow
    '        m = 0
    '        Dim conds() As String

    '        Dim AndConds() As String
    '        Dim OrConds() As String
    '        Dim ok As Boolean = False
    '        Dim insSQL As String = String.Empty
    '        Dim sctSQL As String = String.Empty
    '        Dim ProcessConds As ConditionToProcess()
    '        Dim CondToProcess As ConditionToProcess
    '        ReDim ProcessConds(-1)

    '        sql = sql.Replace(" And ", " AND ").Replace(" and ", " AND ").Replace(" Or ", " OR ").Replace(" or ", " OR ")

    '        If sql.ToUpper.Contains(" AND ") AndAlso sql.ToUpper.Contains(" OR ") Then
    '            AndConds = Split(sql, " AND ")
    '            OrConds = Split(sql, " OR ")
    '            If IsArrayBalanced(AndConds) Then
    '                For i = 0 To AndConds.Length - 1
    '                    cond = AndConds(i).Replace("(", "").Replace(")", "").Trim
    '                    If cond.ToUpper.Contains(" OR ") Then
    '                        g += 1
    '                        group = "Group " & g.ToString

    '                        'make group condition to process
    '                        CondToProcess = New ConditionToProcess
    '                        CondToProcess.LogicalOperator = "And"
    '                        CondToProcess.GroupName = group
    '                        CondToProcess.IsGroup = True
    '                        l = ProcessConds.Length
    '                        ReDim Preserve ProcessConds(l)
    '                        ProcessConds(l) = CondToProcess

    '                        logical = "Or"
    '                        OrConds = Split(cond, " OR ")
    '                        For j = 0 To OrConds.Length - 1
    '                            cond = OrConds(j)
    '                            If cond.ToUpper.Contains(" BETWEEN ") Then
    '                                cond = cond.Replace(" & ", " AND ")
    '                            End If
    '                            l = ProcessConds.Length
    '                            ReDim Preserve ProcessConds(l)
    '                            CondToProcess = New ConditionToProcess
    '                            CondToProcess.Condition = cond
    '                            CondToProcess.LogicalOperator = logical
    '                            CondToProcess.GroupName = group
    '                            ProcessConds(l) = CondToProcess
    '                        Next
    '                    Else
    '                        If cond.ToUpper.Contains(" BETWEEN ") Then
    '                            cond = cond.Replace(" & ", " AND ")
    '                        End If

    '                        l = ProcessConds.Length
    '                        ReDim Preserve ProcessConds(l)
    '                        CondToProcess = New ConditionToProcess
    '                        CondToProcess.Condition = cond
    '                        CondToProcess.LogicalOperator = "And"
    '                        CondToProcess.GroupName = String.Empty
    '                        ProcessConds(l) = CondToProcess
    '                    End If
    '                Next
    '            ElseIf IsArrayBalanced(OrConds) Then
    '                For i = 0 To OrConds.Length - 1
    '                    cond = OrConds(i).Replace("(", "").Replace(")", "").Trim
    '                    If cond.ToUpper.Contains(" AND ") Then
    '                        g += 1
    '                        group = "Group " & g.ToString

    '                        'make group condition to process
    '                        CondToProcess = New ConditionToProcess
    '                        CondToProcess.LogicalOperator = "Or"
    '                        CondToProcess.GroupName = group
    '                        CondToProcess.IsGroup = True
    '                        l = ProcessConds.Length
    '                        ReDim Preserve ProcessConds(l)
    '                        ProcessConds(l) = CondToProcess

    '                        logical = "And"
    '                        AndConds = Split(cond, " AND ")
    '                        For j = 0 To AndConds.Length - 1
    '                            cond = AndConds(j)
    '                            If cond.ToUpper.Contains(" BETWEEN ") Then
    '                                cond = cond.Replace(" & ", " AND ")
    '                            End If
    '                            l = ProcessConds.Length
    '                            ReDim Preserve ProcessConds(l)
    '                            CondToProcess = New ConditionToProcess
    '                            CondToProcess.Condition = cond
    '                            CondToProcess.LogicalOperator = logical
    '                            CondToProcess.GroupName = group
    '                            ProcessConds(l) = CondToProcess
    '                        Next
    '                    Else
    '                        If cond.ToUpper.Contains(" BETWEEN ") Then
    '                            cond = cond.Replace(" & ", " AND ")
    '                        End If

    '                        l = ProcessConds.Length
    '                        ReDim Preserve ProcessConds(l)
    '                        CondToProcess = New ConditionToProcess
    '                        CondToProcess.Condition = cond
    '                        CondToProcess.LogicalOperator = "Or"
    '                        CondToProcess.GroupName = String.Empty
    '                        ProcessConds(l) = CondToProcess
    '                    End If
    '                Next
    '            End If
    '        Else
    '            If sql.ToUpper.Contains(" OR ") Then logical = "Or"
    '            If logical = "And" Then
    '                conds = Split(sql, " AND ")
    '            Else
    '                conds = Split(sql, " OR ")
    '            End If

    '            For i = 0 To conds.Length - 1
    '                cond = conds(i)
    '                If cond.ToUpper.Contains(" BETWEEN ") Then
    '                    cond = cond.Replace(" & ", " AND ")
    '                End If
    '                l = ProcessConds.Length
    '                ReDim Preserve ProcessConds(l)
    '                CondToProcess = New ConditionToProcess
    '                CondToProcess.Condition = cond
    '                CondToProcess.LogicalOperator = logical
    '                CondToProcess.GroupName = String.Empty
    '                ProcessConds(l) = CondToProcess
    '            Next
    '        End If

    '        For i = 0 To ProcessConds.Length - 1

    '            'handle group
    '            If ProcessConds(i).IsGroup Then
    '                logical = ProcessConds(i).LogicalOperator
    '                group = ProcessConds(i).GroupName
    '                m = m + 1
    '                myRow = dt.NewRow
    '                myRow.Item("Tbl1") = String.Empty
    '                myRow.Item("Tbl1Fld1") = String.Empty
    '                myRow.Item("Tbl2") = String.Empty
    '                myRow.Item("Tbl2Fld2") = String.Empty
    '                myRow.Item("Tbl3") = String.Empty
    '                myRow.Item("Tbl3Fld3") = String.Empty
    '                myRow.Item("Oper") = String.Empty
    '                myRow.Item("Type") = String.Empty
    '                myRow.Item("comments") = String.Empty
    '                myRow.Item("RecOrder") = m
    '                myRow.Item("Logical") = logical
    '                myRow.Item("Group") = group
    '                dt.Rows.Add(myRow)

    '                Continue For
    '            End If

    '            cond = ProcessConds(i).Condition
    '            If cond.Trim.Length = 0 Then
    '                Continue For
    '            End If
    '            logical = ProcessConds(i).LogicalOperator
    '            If logical = String.Empty Then logical = "And"
    '            group = ProcessConds(i).GroupName
    '            m = m + 1
    '            lside = ""
    '            rside = ""
    '            typ = ""
    '            sta = ""
    '            opr = ""
    '            oper = ""

    '            For j = 0 To n - 1
    '                If opers(j).Trim = "" Then Continue For
    '                l = cond.ToUpper.IndexOf(opers(j).ToUpper)
    '                If cond.IndexOf("'") < l Then
    '                    Continue For   'oper(j) is in the right side of condition
    '                End If
    '                If l > 0 Then
    '                    opr = opers(j).Trim
    '                    oper = opr
    '                    If opr.ToUpper.Contains("BETWEEN") Then _
    '                        cond = cond.Replace(" & ", " AND ")
    '                    lside = cond.Substring(0, l).Trim.Replace("(", "").Replace(")", "")

    '                    rside = cond.Substring(l + opers(j).Length).Trim.Replace("(", "").Replace(")", "").Replace("""", "")
    '                    If opr.ToUpper.Contains("LIKE") Then
    '                        Dim res As String = rside.Replace("'", "")
    '                        Dim first As String = res.Substring(0, 1)
    '                        Dim last As String = res.Substring(res.Length - 1, 1)
    '                        If first = "%" And last = "%" Then
    '                            If opr.ToUpper.Contains("NOT") Then
    '                                oper = "Not Contains"
    '                            Else
    '                                oper = "Contains"
    '                            End If
    '                        ElseIf first = "%" Then
    '                            If opr.ToUpper.Contains("NOT") Then
    '                                oper = "Not EndsWith"
    '                            Else
    '                                oper = "EndsWith"
    '                            End If
    '                        ElseIf last = "%" Then
    '                            If opr.ToUpper.Contains("NOT") Then
    '                                oper = "Not StartsWith"
    '                            Else
    '                                oper = "StartsWith"
    '                            End If
    '                        End If
    '                        rside = rside.Replace("%", "")
    '                    End If
    '                    opr = oper
    '                    Exit For
    '                End If
    '            Next
    '            If l = -1 Then Continue For

    '            tbl1 = ""
    '            fld1 = lside.Trim
    '            fldfullname = FixTableAndField(dtb, tbl1, fld1, myconstr, myconprv, er)
    '            If fldfullname = "" Then
    '                'lside is static
    '                tbl2 = ""
    '                fld2 = ""
    '                sta = lside.Trim
    '                'opr = oper
    '                typ = "Static"
    '                tbl1 = ""
    '                fld1 = rside.Trim
    '                fldfullname = FixTableAndField(dtb, tbl1, fld1, myconstr, myconprv, er)
    '                If fldfullname = "" Then
    '                    ret = ret & "ERROR!! converting the condition: " & cond
    '                    Continue For
    '                Else
    '                    tbl3 = ""
    '                    fld3 = ""
    '                End If
    '            Else
    '                If rside.ToUpper.Contains(" AND ") Then  'between
    '                    Dim b As Boolean = IsDateTimeField(fldfullname, myconstr, myconprv, er)
    '                    If b Then
    '                        tbl2 = "Date1"
    '                        tbl3 = "Date2"
    '                    Else
    '                        tbl2 = "Value1"
    '                        tbl3 = "Value2"
    '                    End If
    '                    typ = "Btw"
    '                    k = rside.ToUpper.IndexOf(" AND ")
    '                    fld2 = rside.Substring(0, k).Trim
    '                    If fld2.ToUpper.StartsWith("DATE") AndAlso fld2.Contains("(") AndAlso fld2.Contains(")") Then
    '                        'tbl2 = "Date1"
    '                        'relative
    '                        typ = typ & "RD1"
    '                    Else 'not relative date
    '                        fldfullname = FixTableAndField(dtb, "", fld2, myconstr, myconprv, er)
    '                        If fldfullname = "" Then  'value
    '                            fld2 = rside.Substring(0, k).Trim
    '                            If b Then
    '                                'tbl2 = "Date1"
    '                                typ = typ & "DT1"
    '                            Else
    '                                'tbl2 = ""
    '                                typ = typ & "ST1"
    '                            End If
    '                        Else 'field
    '                            typ = typ & "FD1"
    '                        End If
    '                    End If
    '                    'tbl3 = ""
    '                    fld3 = rside.Substring(k + 5).Trim
    '                    If fld3.ToUpper.StartsWith("DATE") AndAlso fld3.Contains("(") AndAlso fld3.Contains(")") Then
    '                        'tbl3 = "Date2"
    '                        'relative
    '                        typ = typ & "RD2"
    '                    Else 'not relative date
    '                        fldfullname = FixTableAndField(dtb, tbl3, fld3, myconstr, myconprv, er)
    '                        If fldfullname = "" Then  'value
    '                            fld3 = rside.Substring(k + 5).Trim
    '                            If b Then
    '                                'tbl3 = "Date1"
    '                                typ = typ & "DT2"
    '                            Else
    '                                'tbl2 = ""
    '                                typ = typ & "ST2"
    '                            End If
    '                        Else 'field
    '                            typ = typ & "FD2"
    '                        End If
    '                    End If
    '                Else  'not between
    '                    fldfullname = FixTableAndField(dtb, tbl2, fld2, myconstr, myconprv, er)
    '                    If fldfullname = "" Then
    '                        typ = "Static"
    '                        sta = rside.Trim
    '                        tbl2 = ""
    '                        fld2 = ""
    '                        tbl3 = ""
    '                        fld3 = ""
    '                    Else
    '                        typ = "Field"
    '                        tbl3 = ""
    '                        fld3 = ""
    '                        sta = fldfullname
    '                    End If
    '                End If
    '            End If

    '            myRow = dt.NewRow
    '            myRow.Item("Tbl1") = tbl1
    '            myRow.Item("Tbl1Fld1") = fld1
    '            myRow.Item("Tbl2") = tbl2.Replace("'", "").Replace("""", "")
    '            myRow.Item("Tbl2Fld2") = fld2.Replace("'", "").Replace("""", "")
    '            myRow.Item("Tbl3") = tbl3.Replace("'", "").Replace("""", "")
    '            myRow.Item("Tbl3Fld3") = fld3.Replace("'", "").Replace("""", "")
    '            myRow.Item("Oper") = opr
    '            myRow.Item("Type") = typ
    '            myRow.Item("comments") = sta.Replace("'", "").Replace("""", "")
    '            myRow.Item("RecOrder") = m
    '            myRow.Item("Logical") = logical
    '            myRow.Item("Group") = group
    '            dt.Rows.Add(myRow)
    '            'insert 
    '            Dim fieldtype As String = GetFieldDataType(tbl1, fld1, myconstr, myconprv)
    '            Dim IsDate As Boolean = ((fieldtype.ToUpper <> "TIME") AndAlso (fieldtype.ToUpper.Contains("DATE") OrElse fieldtype.ToUpper.Contains("TIME")))

    '            Dim sorttype As String = typ
    '        Next

    '        'make WHERE statement from dt
    '        'WHEREs
    '        Dim val As String = String.Empty
    '        'Dim typ As String = String.Empty

    '        Dim RecOrder As String = String.Empty
    '        'Dim Logical As String = String.Empty
    '        'Dim Group As String = String.Empty

    '        Dim HasNot As Boolean = False
    '        Dim qt As String = """"
    '        Dim dblqt = qt & qt
    '        Dim dv As New DataView
    '        dv = dt.DefaultView
    '        Dim fltr As String = String.Empty
    '        If Not dt Is Nothing AndAlso Not dt.Rows Is Nothing AndAlso dt.Rows.Count > 0 Then
    '            'Dim myprovider As String = userconprv
    '            Dim qfield As String = String.Empty
    '            ret = ret & "  WHERE "

    '            Dim htGroup As New Hashtable
    '            Dim PrevGroup As String = String.Empty
    '            For i = 0 To dt.Rows.Count - 1
    '                'ret = ret & " ( "
    '                oper = dt.Rows(i)("Oper").ToString
    '                tbl1 = FixReservedWords(CorrectTableNameWithDots(dt.Rows(i)("Tbl1").ToString, myconprv), myconprv, myconstr)
    '                fld1 = FixReservedWords(CorrectFieldNameWithDots(dt.Rows(i)("Tbl1Fld1").ToString), myconprv, myconstr)
    '                tbl2 = FixReservedWords(CorrectTableNameWithDots(dt.Rows(i)("Tbl2").ToString, myconprv), myconprv, myconstr)
    '                fld2 = FixReservedWords(CorrectFieldNameWithDots(dt.Rows(i)("Tbl2Fld2").ToString), myconprv, myconstr)
    '                tbl3 = FixReservedWords(CorrectTableNameWithDots(dt.Rows(i)("Tbl3").ToString, myconprv), myconprv, myconstr)
    '                fld3 = FixReservedWords(CorrectFieldNameWithDots(dt.Rows(i)("Tbl3Fld3").ToString), myconprv, myconstr)
    '                fltr = ""

    '                val = dt.Rows(i)("Comments").ToString
    '                logical = dt.Rows(i)("Logical").ToString
    '                group = dt.Rows(i)("Group").ToString

    '                If (group <> String.Empty) Then
    '                    If htGroup(group) Is Nothing Then
    '                        If PrevGroup <> String.Empty Then
    '                            ret &= " )"
    '                        End If
    '                        PrevGroup = group
    '                        htGroup(group) = 1

    '                        'logical = GetConditionGroupLogical(rep, group)  '!!!!!!
    '                        'Dim sql = "SELECT Logical FROM OURReportSQLquery WHERE ReportId = '" & rep & "' AND DOING = 'GROUP' AND " & FixReservedWords("Group") & " = '" & GroupName & "'"
    '                        logical = ""
    '                        'filter dt for Group=group
    '                        fltr = "DOING = 'GROUP' AND Group='" & group & "'"
    '                        dv.RowFilter = fltr
    '                        If dv.ToTable.DefaultView.Table.Rows.Count = 1 Then
    '                            logical = dv.ToTable.Rows(0)("Logical").ToString
    '                        End If
    '                        fltr = ""

    '                        If i > 0 Then
    '                            logical = logical & " ("
    '                        Else
    '                            logical = " ("
    '                        End If
    '                    End If
    '                Else
    '                    If PrevGroup <> String.Empty Then
    '                        ret &= " )"
    '                        PrevGroup = String.Empty
    '                    End If

    '                End If
    '                If logical = String.Empty Then logical = "And"

    '                If logical = " (" OrElse i > 0 Then _
    '                    ret = ret & " " & logical.ToUpper & " "

    '                If val = dblqt Then val = ""

    '                typ = dt.Rows(i)("Type").ToString.Trim

    '                HasNot = oper.ToUpper.Contains("NOT")
    '                If typ = "Field" Then   'to field
    '                    qfield = tbl2 & "." & fld2
    '                    If oper.ToUpper.Contains("STARTSWITH") Then
    '                        If HasNot Then
    '                            oper = "Not Like"
    '                        Else
    '                            oper = "Like"
    '                        End If
    '                        If myconprv.StartsWith("InterSystems.Data.") Then
    '                            ret = ret & " ( " & tbl1 & "." & fld1 & " " & oper & " STRING(" & qfield & ",'%')) "
    '                        Else
    '                            ret = ret & " ( " & tbl1 & "." & fld1 & " " & oper & " CONCAT(" & qfield & ",'%')) "
    '                        End If
    '                    ElseIf oper.ToUpper.Contains("CONTAINS") Then
    '                        If HasNot Then
    '                            oper = "Not Like"
    '                        Else
    '                            oper = "Like"
    '                        End If
    '                        If myconprv.StartsWith("InterSystems.Data.") Then
    '                            ret = ret & " ( " & tbl1 & "." & fld1 & " " & oper & " STRING('%'," & qfield & ",'%')) "
    '                        Else
    '                            ret = ret & " ( " & tbl1 & "." & fld1 & " " & oper & " CONCAT('%'," & qfield & ",'%')) "
    '                        End If
    '                    ElseIf oper.ToUpper.Contains("ENDSWITH") Then
    '                        If HasNot Then
    '                            oper = "Not Like"
    '                        Else
    '                            oper = "Like"
    '                        End If
    '                        If myconprv.StartsWith("InterSystems.Data.") Then
    '                            ret = ret & " ( " & tbl1 & "." & fld1 & " " & oper & " STRING('%'," & qfield & ")) "
    '                        Else
    '                            ret = ret & " ( " & tbl1 & "." & fld1 & " " & oper & " CONCAT('%'," & qfield & ")) "
    '                        End If
    '                    Else
    '                        ret = ret & " ( " & tbl1 & "." & fld1 & " " & oper & " " & qfield & ") "
    '                    End If
    '                ElseIf typ = "RelDate" Then
    '                    If oper.ToUpper.Contains("STARTSWITH") Then
    '                        If HasNot Then
    '                            oper = "Not Like"
    '                        Else
    '                            oper = "Like"
    '                        End If
    '                        If myconprv.StartsWith("InterSystems.Data.") Then
    '                            ret = ret & " ( " & tbl1 & "." & fld1 & " " & oper & " STRING(" & val & ",'%')) "
    '                        Else
    '                            ret = ret & " ( " & tbl1 & "." & fld1 & " " & oper & " CONCAT(" & val & ",'%')) "
    '                        End If
    '                    ElseIf oper.ToUpper.Contains("CONTAINS") Then
    '                        If HasNot Then
    '                            oper = "Not Like"
    '                        Else
    '                            oper = "Like"
    '                        End If
    '                        If myconprv.StartsWith("InterSystems.Data.") Then
    '                            ret = ret & " ( " & tbl1 & "." & fld1 & " " & oper & " STRING('%'," & val & ",'%')) "
    '                        Else
    '                            ret = ret & " ( " & tbl1 & "." & fld1 & " " & oper & " CONCAT('%'," & val & ",'%')) "
    '                        End If
    '                    ElseIf oper.ToUpper.Contains("ENDSWITH") Then
    '                        If HasNot Then
    '                            oper = "Not Like"
    '                        Else
    '                            oper = "Like"
    '                        End If
    '                        If myconprv.StartsWith("InterSystems.Data.") Then
    '                            ret = ret & " ( " & tbl1 & "." & fld1 & " " & oper & " STRING('%'," & val & ")) "
    '                        Else
    '                            ret = ret & " ( " & tbl1 & "." & fld1 & " " & oper & " CONCAT('%'," & val & ")) "
    '                        End If
    '                    Else
    '                        ret = ret & " ( " & tbl1 & "." & fld1 & " " & oper & " " & val & ") "
    '                    End If
    '                ElseIf typ = "Static" Then   'to static
    '                    If TblFieldIsNumeric(tbl1, fld1, myconstr, myconprv) Then  'numeric
    '                        If oper.ToUpper = "IN" OrElse oper.ToUpper = "NOT IN" Then
    '                            ret = ret & " ( " & tbl1 & "." & fld1 & " " & oper & " (" & val & ")) "
    '                        ElseIf oper.ToUpper.Contains("STARTSWITH") Then
    '                            If HasNot Then
    '                                oper = "Not Like"
    '                            Else
    '                                oper = "Like"
    '                            End If
    '                            If myconprv.StartsWith("InterSystems.Data.") Then
    '                                ret = ret & " ( " & tbl1 & "." & fld1 & " " & oper & " STRING(" & val & ",'%')) "
    '                            Else
    '                                ret = ret & " ( " & tbl1 & "." & fld1 & " " & oper & " CONCAT(" & val & ",'%')) "
    '                            End If
    '                        ElseIf oper.ToUpper.Contains("CONTAINS") Then
    '                            If HasNot Then
    '                                oper = "Not Like"
    '                            Else
    '                                oper = "Like"
    '                            End If
    '                            If myconprv.StartsWith("InterSystems.Data.") Then
    '                                ret = ret & " ( " & tbl1 & "." & fld1 & " " & oper & " STRING('%'," & val & ",'%')) "
    '                            Else
    '                                ret = ret & " ( " & tbl1 & "." & fld1 & " " & oper & " CONCAT('%'," & val & ",'%')) "
    '                            End If
    '                        ElseIf oper.ToUpper.Contains("ENDSWITH") Then
    '                            If HasNot Then
    '                                oper = "Not Like"
    '                            Else
    '                                oper = "Like"
    '                            End If
    '                            If myconprv.StartsWith("InterSystems.Data.") Then
    '                                ret = ret & " ( " & tbl1 & "." & fld1 & " " & oper & " STRING('%'," & val & ")) "
    '                            Else
    '                                ret = ret & " ( " & tbl1 & "." & fld1 & " " & oper & " CONCAT('%'," & val & ")) "
    '                            End If
    '                        ElseIf oper.ToUpper.Contains("IS NULL") Then
    '                            If Not HasNot Then
    '                                ret = ret & " ( " & tbl1 & "." & fld1 & " " & oper & ") "
    '                            Else
    '                                ret = ret & " ( Not " & tbl1 & "." & fld1 & " IS NULL) "
    '                            End If
    '                        Else
    '                            ret = ret & " ( " & tbl1 & "." & fld1 & " " & oper & " " & val & ") "
    '                        End If
    '                    Else                                                                                       'string, date, time, or else
    '                        If oper.ToUpper = "IN" OrElse oper.ToUpper = "NOT IN" Then
    '                            ret = ret & " ( " & tbl1 & "." & fld1 & " " & oper & " ('" & val.Replace(",", "','") & "')) "
    '                        ElseIf oper.ToUpper.Contains("STARTSWITH") Then
    '                            If HasNot Then
    '                                oper = "Not Like"
    '                            Else
    '                                oper = "Like"
    '                            End If
    '                            ret = ret & " ( " & tbl1 & "." & fld1 & " " & oper & " '" & val & "%') "
    '                        ElseIf oper.ToUpper.Contains("CONTAINS") Then
    '                            If HasNot Then
    '                                oper = "Not Like"
    '                            Else
    '                                oper = "Like"
    '                            End If
    '                            ret = ret & " ( " & tbl1 & "." & fld1 & " " & oper & " '%" & val & "%') "
    '                        ElseIf oper.ToUpper.Contains("ENDSWITH") Then
    '                            If HasNot Then
    '                                oper = "Not Like"
    '                            Else
    '                                oper = "Like"
    '                            End If
    '                            ret = ret & " ( " & tbl1 & "." & fld1 & " " & oper & " '%" & val & "') "
    '                        ElseIf oper.ToUpper.Contains("IS NULL") Then
    '                            If Not HasNot Then
    '                                ret = ret & " ( " & tbl1 & "." & fld1 & " " & oper & ") "
    '                            Else
    '                                ret = ret & " ( Not " & tbl1 & "." & fld1 & " IS NULL) "
    '                            End If

    '                        Else
    '                            ret = ret & " ( " & tbl1 & "." & fld1 & " " & oper & " '" & val & "') "
    '                        End If
    '                    End If
    '                ElseIf typ = "DateTime" Then
    '                    If oper.ToUpper.Contains("STARTSWITH") Then
    '                        If HasNot Then
    '                            oper = "Not Like"
    '                        Else
    '                            oper = "Like"
    '                        End If
    '                        ret = ret & " ( " & tbl1 & "." & fld1 & " " & oper & " '" & DateToString(CDate(val), myconprv) & "%') "
    '                    ElseIf oper.ToUpper.Contains("CONTAINS") Then
    '                        If HasNot Then
    '                            oper = "Not Like"
    '                        Else
    '                            oper = "Like"
    '                        End If
    '                        ret = ret & " ( " & tbl1 & "." & fld1 & " " & oper & " '%" & DateToString(CDate(val), myconprv) & "%') "

    '                    ElseIf oper.ToUpper.Contains("ENDSWITH") Then
    '                        If HasNot Then
    '                            oper = "Not Like"
    '                        Else
    '                            oper = "Like"
    '                        End If
    '                        ret = ret & " ( " & tbl1 & "." & fld1 & " " & oper & " '%" & DateToString(CDate(val), myconprv) & "') "
    '                    Else
    '                        ret = ret & " ( " & tbl1 & "." & fld1 & " " & oper & " '" & DateToString(CDate(val), myconprv) & "') "
    '                    End If
    '                ElseIf typ = "BtwFields" OrElse typ = "BtwFD1FD2" Then   'between fields  BtwFD1FD2
    '                    ret = ret & " ( " & tbl1 & "." & fld1 & " " & oper & " " & tbl2 & "." & fld2 & " AND " & tbl3 & "." & fld3 & ") "
    '                ElseIf typ = "BtwDates" OrElse typ = "BtwDT1DT2" Then   'between dates   BtwDT1DT2
    '                    ret = ret & " ( " & tbl1 & "." & fld1 & " " & oper & " '" & DateToString(CDate(fld2), myconprv) & "' AND '" & DateToString(CDate(fld3), myconprv) & "') "
    '                ElseIf typ = "BwRD1Date2" OrElse typ = "BtwRD1DT2" Then   ' between relative date and date   BtwRD1DT2
    '                    ret = ret & " ( " & tbl1 & "." & fld1 & " " & oper & " " & fld2 & " AND '" & DateToString(CDate(fld3), myconprv) & "') "
    '                ElseIf typ = "BwDate1RD2" OrElse typ = "BtwDT1RD2" Then   'between date and relative date    BtwDT1RD2
    '                    ret = ret & " ( " & tbl1 & "." & fld1 & " " & oper & " '" & DateToString(CDate(fld2), myconprv) & "' AND " & fld3 & ") "
    '                ElseIf typ = "BtwRD1RD2" Then   'between relative dates   BtwRD1RD2
    '                    ret = ret & " ( " & tbl1 & "." & fld1 & " " & oper & " " & fld2 & " AND " & fld3 & ") "
    '                ElseIf typ = "BtwFldDate" OrElse typ = "BtwFD1DT2" Then   'between field and date  BtwFD1DT2
    '                    ret = ret & " ( " & tbl1 & "." & fld1 & " " & oper & " " & tbl2 & "." & fld2 & " And '" & DateToString(CDate(fld3), myconprv) & "') "
    '                ElseIf typ = "BtwFldRD2" OrElse typ = "BtwFD1RD2" Then   'between field and relative date   BtwFD1RD2
    '                    ret = ret & " ( " & tbl1 & "." & fld1 & " " & oper & " " & tbl2 & "." & fld2 & " AND " & fld3 & ") "
    '                ElseIf typ = "BtwDateFld" OrElse typ = "BtwDT1FD2" Then   'between date and field   BtwDT1FD2
    '                    ret = ret & " ( " & tbl1 & "." & fld1 & " " & oper & " '" & DateToString(CDate(fld2), myconprv) & "' AND " & tbl3 & "." & fld3 & ") "
    '                ElseIf typ = "BtwRD1Fld" OrElse typ = "BtwRD1FD2" Then   'between relative date and field  BtwRD1FD2
    '                    ret = ret & " ( " & tbl1 & "." & fld1 & " " & oper & " " & fld2 & " AND " & tbl3 & "." & fld3 & ") "
    '                ElseIf typ = "BtwValues" OrElse typ = "BtwST1ST2" Then   'between static values    BtwST1ST2
    '                    If TblFieldIsNumeric(tbl1, fld1, myconstr, myconprv) Then
    '                        ret = ret & " ( " & tbl1 & "." & fld1 & " " & oper & " " & fld2 & " AND " & fld3 & ") "
    '                    Else
    '                        ret = ret & " ( " & tbl1 & "." & fld1 & " " & oper & " '" & fld2 & "' AND '" & fld3 & "') "
    '                    End If
    '                ElseIf typ = "BtwValFld" OrElse typ = "BtwST1FD2" Then   'between static value and field BtwST1FD2
    '                    If TblFieldIsNumeric(tbl1, fld1, myconstr, myconprv) Then
    '                        ret = ret & " ( " & tbl1 & "." & fld1 & " " & oper & " " & fld2 & " AND " & tbl3 & "." & fld3 & ") "
    '                    Else
    '                        ret = ret & " ( " & tbl1 & "." & fld1 & " " & oper & " '" & fld2 & "' AND " & tbl3 & "." & fld3 & ") "
    '                    End If
    '                ElseIf typ = "BtwFldVal" OrElse typ = "BtwFD1ST2" Then   'between field and static value  BtwFD1ST2
    '                    If TblFieldIsNumeric(tbl1, fld1, myconstr, myconprv) Then
    '                        ret = ret & " ( " & tbl1 & "." & fld1 & " " & oper & " " & tbl2 & "." & fld2 & " And " & fld3 & ") "
    '                    Else
    '                        ret = ret & " ( " & tbl1 & "." & fld1 & " " & oper & " " & tbl2 & "." & fld2 & " And '" & fld3 & "') "
    '                    End If

    '                End If

    '                If i = dt.Rows.Count - 1 AndAlso PrevGroup <> "" Then ret &= ")"

    '            Next
    '        End If
    '        Return ret
    '    Catch ex As Exception
    '        er = ex.Message
    '        Return "ERROR!! " & er
    '    End Try
    '    Return ret
    'End Function
    Public Function GetListOfTablesFromSQLquery(ByVal sql As String, Optional ByVal myconstr As String = "", Optional ByVal myconprv As String = "", Optional ByRef er As String = "") As DataTable
        Dim ret As String = String.Empty
        Dim dt As New DataTable
        If sql = "" Then
            Return dt
        End If
        Dim i, j, n, k, l As Integer
        Dim tbl As String = String.Empty
        Dim tblspace As String = String.Empty
        Dim sqlj As String = String.Empty
        Try
            i = sql.ToUpper.IndexOf(" FROM ")
            sql = sql.Substring(i + 6).Trim
            k = sql.ToUpper.IndexOf("WHERE")
            If k > 0 Then
                sql = sql.Substring(0, k)
            End If
            k = sql.ToUpper.IndexOf("GROUP BY")
            If k > 0 Then
                sql = sql.Substring(0, k)
            End If
            k = sql.ToUpper.IndexOf("ORDER BY")
            If k > 0 Then
                sql = sql.Substring(0, k)
            End If
            k = sql.ToUpper.IndexOf("HAVING")
            If k > 0 Then
                sql = sql.Substring(0, k)
            End If
            'sql = sql.Trim & " "
            'at this point sql is list of tables separated with commas or JOINs
            dt.TableName = "Table1"
            dt.Columns.Add("Tbl1")
            Dim myRow As DataRow
            k = sql.ToUpper.IndexOf(" JOIN ")
            If k < 0 Then
                Dim ar = sql.Split(",")
                For j = 0 To ar.Length - 1
                    tbl = ar(j).ToLower
                    If tbl.Trim.Length = 0 Then
                        Exit For
                    End If
                    tbl = tbl.Replace("`", "")
                    tbl = tbl.Replace("[", "").Replace("]", "").Trim
                    tbl = tbl.Replace("""", "")
                    'check if tbl is a table
                    If myconstr.Trim <> "" AndAlso Not TableExists(tbl.Trim, myconstr, myconprv, er) Then
                        Continue For
                    Else
                        myRow = dt.NewRow
                        myRow.Item(0) = tbl
                        dt.Rows.Add(myRow)
                    End If
                Next

            Else
                sqlj = sql

                Dim htTable As New Hashtable

                If sqlj.Contains(" INNER ") Then sqlj = sqlj.Replace(" INNER ", "")
                If sqlj.Contains(" LEFT ") Then sqlj = sqlj.Replace(" LEFT ", "")
                If sqlj.Contains(" RIGHT ") Then sqlj = sqlj.Replace(" RIGHT ", "")
                If sqlj.Contains(" OUTER ") Then sqlj = sqlj.Replace(" OUTER ", "")
                sqlj = sqlj.Replace("JOIN", "|")

                'sqlj = sqlj.Replace(" INNER JOIN ", " JOIN ")
                'sqlj = sqlj.Replace(" LEFT JOIN ", " JOIN ")
                'sqlj = sqlj.Replace(" RIGHT JOIN ", " JOIN ")
                'sqlj = sqlj.Replace(" OUTER JOIN ", " JOIN ")
                'sqlj = sqlj.Replace(" JOIN ", "|")

                'after JOIN
                Dim ar = sqlj.Split("|")
                For j = 0 To ar.Length - 1
                    tbl = ar(j).Trim
                    If tbl.Length = 0 Then
                        Exit For
                    End If
                    If j = 0 AndAlso tbl.IndexOf(",") > 0 Then
                        Dim tb = tbl.Split(",")
                        For n = 0 To tb.Length - 1
                            If tb(n).Trim.Length = 0 Then
                                Exit For
                            End If
                            'table name tb(n) put in Row
                            tb(n) = tb(n).Replace("`", "")
                            tb(n) = tb(n).Replace("[", "").Replace("]", "")
                            tb(n) = tb(n).Replace("""", "")
                            tb(n) = tb(n).ToLower.Trim


                            'tbl = tbl.Replace("""", "")
                            'check if tbl is a table
                            If myconstr.Trim <> "" AndAlso Not TableExists(tb(n).Trim, myconstr, myconprv, er) Then
                                Continue For
                            Else
                                If htTable(tb(n)) = String.Empty Then
                                    htTable.Add(tb(n), "1")
                                    myRow = dt.NewRow
                                    myRow.Item(0) = tb(n).ToLower.Trim
                                    dt.Rows.Add(myRow)
                                End If
                            End If
                        Next
                        Continue For
                    End If
                    'remove ON statement
                    l = tbl.ToUpper.IndexOf(" ON ")
                    If l > 0 Then
                        tbl = tbl.Substring(0, l).Trim
                    End If
                    'table name tbl put in Row
                    tbl = tbl.Replace("`", "")
                    tbl = tbl.Replace("[", "").Replace("]", "")
                    tbl = tbl.Replace("""", "")
                    tbl = tbl.ToLower.Trim

                    'check if tbl is a table
                    If myconstr.Trim <> "" AndAlso Not TableExists(tbl.Trim, myconstr, myconprv, er) Then
                        Continue For
                    Else
                        If htTable(tbl) = String.Empty Then
                            myRow = dt.NewRow
                            myRow.Item(0) = tbl.ToLower
                            dt.Rows.Add(myRow)
                            htTable.Add(tbl, "1")
                        End If
                    End If
                Next
            End If
            Dim dv As New DataView
            dv = dt.DefaultView
            dv.Sort = "Tbl1"
            dt = dv.Table
            Return dt
        Catch ex As Exception
            er = ex.Message
            Return Nothing
        End Try
        Return dt
    End Function

    Public Function GetListOfTableColumnsFromSchema(ByVal tbl As String, Optional ByVal userconstr As String = "", Optional ByVal userconprv As String = "", Optional ByRef err As String = "") As DataView
        'NOT IN USE
        Dim dv As DataView = Nothing
        Dim myprovider, sqls, strConnect As String
        sqls = String.Empty
        If userconstr = "" Then
            strConnect = System.Configuration.ConfigurationManager.ConnectionStrings.Item("mySQLconnection").ToString
            myprovider = System.Configuration.ConfigurationManager.ConnectionStrings.Item("mySQLconnection").ProviderName.ToString
        Else
            strConnect = userconstr
            myprovider = userconprv
        End If
        If myprovider.StartsWith("InterSystems.Data.") Then
            sqls = "Select Name As COLUMN_NAME FROM %Dictionary.PropertyDefinition WHERE UCASE(parent) = UCASE('" & tbl & "') AND Cardinality Is NULL"

        ElseIf myprovider = "MySql.Data.MySqlClient" Then
            Dim db As String = GetDataBase(strConnect, myprovider)
            sqls = "SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_SCHEMA ='" & db.ToLower & "' AND TABLE_NAME = '" & tbl & "'"

        ElseIf myprovider = "Oracle.ManagedDataAccess.Client" Then
            'Dim db As String = GetDataBase(strConnect)
            sqls = "SELECT * FROM all_tab_cols WHERE UPPER(TABLE_NAME) = UPPER('" & tbl & "')"

        ElseIf myprovider = "Npgsql" Then  'PostgreSQL  Npgsql
            Dim db As String = GetDataBase(strConnect, myprovider)
            sqls = "SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_SCHEMA ='public' AND LOWER(TABLE_CATALOG) ='" & db.ToLower & "' AND TABLE_NAME = '" & tbl & "'"

        Else
            sqls = "SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = '" & tbl & "'"
        End If
        If userconprv <> "System.Data.Odbc" AndAlso userconprv <> "System.Data.OleDb" Then
            dv = mRecords(sqls, err, userconstr, userconprv)
        End If

        If myprovider.StartsWith("InterSystems.Data") AndAlso err = "" Then
            Dim dt As DataTable = dv.Table
            Dim NewRow As Object() = New Object(0) {"ID"}
            dt.BeginLoadData()
            dt.LoadDataRow(NewRow, True)
            dt.EndLoadData()
        End If
        Return dv
    End Function
    Public Function CorrectTablesJoinsFromSQLquery(ByVal msql As String, Optional ByVal myconstr As String = "", Optional ByVal myconprv As String = "", Optional ByRef er As String = "") As String
        Dim ret As String = String.Empty
        Dim sql As String = msql.Trim 'dri.Rows(0)("SQLquerytext").ToString
        If sql = "" Then
            Return msql
            Exit Function
        End If
        sql = sql.Replace("`", "").Replace("[", "").Replace("]", "").Replace("""", "")
        Dim ddt As DataTable = GetListOfTablesFromSQLquery(msql, myconstr, myconprv, er)
        Dim i, j, n, m, k, l, q As Integer
        Dim tbl As String = String.Empty
        Dim tbls As String = String.Empty
        Dim fld2 As String = String.Empty
        Dim tbl1 As String = String.Empty
        Dim tbl2 As String = String.Empty
        Dim fld1 As String = String.Empty
        Dim typ As String = String.Empty
        Dim sqlu As String = String.Empty
        Dim sqlj As String = String.Empty
        Dim sqlon As String = String.Empty
        Try
            i = sql.ToUpper.IndexOf(" FROM ")
            sql = sql.Substring(i + 6).Trim
            k = sql.ToUpper.IndexOf("WHERE")
            If k > 0 Then
                sql = sql.Substring(0, k)
            End If
            k = sql.ToUpper.IndexOf("GROUP BY")
            If k > 0 Then
                sql = sql.Substring(0, k)
            End If
            k = sql.ToUpper.IndexOf("ORDER BY")
            If k > 0 Then
                sql = sql.Substring(0, k)
            End If
            k = sql.ToUpper.IndexOf("HAVING")
            If k > 0 Then
                sql = sql.Substring(0, k)
            End If

            'at this point here can be sets of Joins and single tables separated by commas
            Dim sqls() As String
            sqls = sql.Split(",")
            For q = 0 To sqls.Length - 1
                sql = sqls(q).Trim
                If sql.ToUpper.IndexOf(" ON ") < 0 Then 'add single tables to tbls separated by commas
                    'If tbls = "" Then
                    '    tbls = Sql
                    'Else
                    '    tbls = tbls & "," & Sql
                    'End If
                    msql = ReplaceWholeWord(msql, sql, """" & sql & """")
                    msql = msql.Replace("""""", """")
                    Continue For
                End If

                Dim oper As String = "1"
                'at this point sql is list of tables separated with JOINs, no commas
                sqlu = sql.ToUpper & " "
                k = sqlu.IndexOf(" JOIN ")
                If k > 0 Then
                    Dim tbar() As String = Split(sqlu, " JOIN ")
                    m = tbar.Length
                    For j = 0 To m - 1
                        l = tbar(j).Length
                        tbl = tbar(j).Trim
                        If tbl.Trim.Length = 0 Then
                            Exit For
                        End If
                        If j = 0 Then
                            n = tbl.IndexOf(" ")  'tbl has " " in the end
                            If n > 0 Then
                                tbl1 = tbl.Substring(0, n)
                                tbl1 = FixTableName(tbl1, myconstr, myconprv, er)
                                typ = tbl.Substring(n).Trim.ToUpper  'LEFT, RIGHT, INNER, OUTER or empty
                                If typ <> "LEFT" AndAlso typ <> "RIGHT" AndAlso typ <> "OUTER" Then
                                    typ = "INNER"
                                End If
                                typ = typ & " JOIN"
                            End If
                            Continue For
                        End If
                        sqlj = tbl.Trim.Substring(tbl.LastIndexOf(" ")).Trim.ToUpper
                        If sqlj = "LEFT" OrElse sqlj = "RIGHT" OrElse sqlj = "OUTER" OrElse sqlj = "INNER" Then
                            tbl = tbl.Trim.Substring(0, tbl.LastIndexOf(" ")).Trim.ToUpper
                        End If
                        If tbl.IndexOf(" ON ") > 0 Then  'it is not the very first table
                            tbl2 = tbl.Substring(0, tbl.ToUpper.IndexOf(" ON ")).Trim
                            tbl2 = FixTableName(tbl2, myconstr, myconprv, er)
                            sqlj = tbl.Substring(tbl.ToUpper.IndexOf(" ON ") + 4)
                            'there could be few ONs
                            Dim fdar() As String = Split(sqlj, " AND ")
                            For i = 0 To fdar.Length - 1
                                sqlon = fdar(i)
                                fld1 = sqlon.Substring(0, sqlon.IndexOf("="))
                                fld2 = sqlon.Substring(sqlon.IndexOf("="))
                                fld1 = FixFieldName(ddt, fld1, tbl1, myconstr, myconprv, er)
                                fld2 = FixFieldName(ddt, fld2, tbl2, myconstr, myconprv, er)

                                'TODO fix ON statement
                                msql = ReplaceWholeWord(msql, tbl1, """" & tbl1.Trim & """")
                                msql = msql.Replace("""""", """")
                                msql = ReplaceWholeWord(msql, tbl2, """" & tbl2.Trim & """")
                                msql = msql.Replace("""""", """")
                                msql = ReplaceWholeWord(msql, fld1, """" & fld1.Trim & """")
                                msql = msql.Replace("""""", """")
                                msql = ReplaceWholeWord(msql, fld2, """" & fld2.Trim & """")
                                msql = msql.Replace("""""", """")

                            Next
                            tbl1 = tbl2
                        End If
                    Next
                End If
            Next
            er = "Single Tables: " & tbls
            Return msql
        Catch ex As Exception
            er = "ERROR !! " & ex.Message
            Return Nothing
        End Try
        Return msql
    End Function
    Public Function FixTableName(ByVal tbl As String, Optional ByVal userconstr As String = "", Optional ByVal userconprv As String = "", Optional ByRef err As String = "") As String
        tbl = tbl.Replace("`", "")
        tbl = tbl.Replace("[", "").Replace("]", "").Trim
        tbl = tbl.Replace("""", "")
        tbl = tbl.Replace("(", "").Replace(")", "").Trim

        tbl = FixReservedWords(tbl, userconprv, userconstr)
        Return tbl
    End Function
    Public Function FixFieldName(ByVal ddt As DataTable, ByVal fld As String, ByRef tabl As String, Optional ByVal userconstr As String = "", Optional ByVal userconprv As String = "", Optional ByRef err As String = "") As String
        'If ddt Is Nothing OrElse ddt.Rows.Count = 0 Then
        '    ddt = GetReportTablesFromSQLqueryText(repid)
        'End If
        fld = fld.Replace("=", "")
        fld = fld.Replace("`", "")
        fld = fld.Replace("[", "").Replace("]", "").Trim
        fld = fld.Replace("""", "")
        fld = fld.Replace("(", "").Replace(")", "").Trim
        'get real field name from the list of report tables and their columns
        Dim ddf As DataTable
        Dim tbl As String = String.Empty
        If fld.IndexOf(".") > 0 Then
            tbl = fld.Substring(0, fld.LastIndexOf("."))
            fld = fld.Substring(fld.LastIndexOf(".") + 1)
            For i = 0 To ddt.Rows.Count - 1
                If ddt.Rows(i)("Tbl1").ToUpper = tbl.ToUpper Then
                    tabl = ddt.Rows(i)("Tbl1")
                    Exit For
                End If
            Next

            ddf = GetListOfTableColumns(tabl, userconstr, userconprv).Table
            For i = 0 To ddf.Rows.Count - 1
                If ddf.Rows(i)("COLUMN_NAME").ToUpper = fld.ToUpper Then
                    fld = ddf.Rows(i)("COLUMN_NAME")
                    Exit For
                End If
            Next
        Else
            'table name is missing, find table to the field
            For i = 0 To ddt.Rows.Count - 1
                ddf = GetListOfTableColumns(ddt.Rows(i)("Tbl1").ToString, userconstr, userconprv).Table
                For j = 0 To ddf.Rows.Count - 1
                    If ddf.Rows(j)("COLUMN_NAME").ToUpper = fld.ToUpper Then
                        fld = ddf.Rows(j)("COLUMN_NAME").ToString
                        tabl = ddt.Rows(i)("Tbl1").ToString
                        Exit For
                    End If
                Next
            Next
        End If
        'fld = tbl & "." & fld
        fld = FixReservedWords(fld, userconprv, userconstr)
        Return fld
    End Function
    Function ReplaceWholeWord(Inp As String, Fnd As String, Repl As String) As String
        Dim LettrsNum As String = "ABCDEFGHIJKLMNOPQRSTUVWYXZabcdefghijklmnopqrstuvwyxz0123456789_@"
        Dim FndIndx As Integer = 0
        Dim len As Integer = Fnd.Length
        'Dim ret As String = Inp
        Do Until False
            FndIndx = Inp.IndexOf(Fnd, FndIndx)
            If FndIndx < 0 Then Exit Do
            If FndIndx = 0 AndAlso LettrsNum.Contains(Inp.Substring(FndIndx + len, 1)) = False Then  'Fnd Is in the beginning of Inp
                Inp = Inp.Substring(0, FndIndx) & Repl & Inp.Substring(FndIndx + len)

            ElseIf FndIndx > 0 AndAlso (FndIndx + len = Inp.Length OrElse LettrsNum.Contains(Inp.Substring(FndIndx + len, 1)) = False) Then
                Inp = Inp.Substring(0, FndIndx) & Repl & Inp.Substring(FndIndx + len)
            End If
            FndIndx = FndIndx + Repl.Length
        Loop
        Return Inp
    End Function
    Public Function CorrectFieldsArrayFromSQLquery(ByVal Sql As String, Optional ByVal myconstr As String = "", Optional ByVal myconprv As String = "", Optional ByRef er As String = "") As String
        Dim selflgs As String = String.Empty
        If Sql = "" Then
            Return ""
        End If
        Dim i, j As Integer
        Dim tbl As String = String.Empty
        Try
            i = Sql.ToUpper.IndexOf(" FROM ")
            If i >= 0 Then Sql = Sql.Substring(0, i).Trim
            If Sql.ToUpper.StartsWith("SELECT ") Then
                Sql = Sql.Substring(6).Trim
            End If
            If Sql.ToUpper.StartsWith("DISTINCT ") Then
                Sql = Sql.Substring(8).Trim
            End If
            If Sql = "*" Then
                Return Sql
            End If
            Dim flds() As String
            flds = Sql.Split(",")
            For i = 0 To flds.Length - 1
                j = flds(i).ToUpper.IndexOf(" AS ")
                If j > 0 Then flds(i) = flds(i).Substring(0, j).Trim
                j = flds(i).ToUpper.IndexOf("(")
                If j > 0 Then flds(i) = flds(i).Substring(j).Trim
                flds(i) = flds(i).Replace(")", "")
                flds(i) = flds(i).Replace("(", "")
                'no need to add table name
                If selflgs = String.Empty Then
                    selflgs = FixReservedWords(flds(i).Trim, myconprv, myconstr)
                Else
                    selflgs = selflgs & ", " & FixReservedWords(flds(i).Trim, myconprv, myconstr)
                End If
            Next
            Return selflgs
        Catch ex As Exception
            Return "" '"ERROR!! " & ex.Message.ToString
        End Try
    End Function
    Public Function RunSP(ByVal mySP As String, ByVal Nparameters As Integer, ByVal ParamName As Array, ByVal ParamType As Array, ByVal ParamValue As Array, Optional ByVal myconstr As String = "", Optional ByVal myconprv As String = "") As String
        Dim r As String
        r = ""
        Dim myconstring, myprovider As String
        Try
            If myconstr.Trim = "" Then
                myconstring = System.Configuration.ConfigurationManager.ConnectionStrings.Item("mySQLconnection").ToString
                myprovider = System.Configuration.ConfigurationManager.ConnectionStrings.Item("mySQLconnection").ProviderName.ToString
            Else
                myconstring = myconstr
                myprovider = myconprv
            End If
            If myprovider = "InterSystems.Data.IRISClient" Then
                r = RunSP_IRIS(mySP, Nparameters, ParamName, ParamType, ParamValue, myconstring)
            ElseIf myprovider = "InterSystems.Data.CacheClient" Then
                r = RunSP_Cache(mySP, Nparameters, ParamName, ParamType, ParamValue, myconstring)
            ElseIf myprovider = "MySql.Data.MySqlClient" Then
                Dim ret As String = String.Empty
                ret = CorrectSQLforMySql(ret, myconstring)
                Dim myConnection As MySqlConnection
                Dim myCommand As New MySqlCommand
                myConnection = New MySqlConnection(myconstring)
                myCommand.Connection = myConnection
                myCommand.CommandType = CommandType.Text
                myCommand.CommandTimeout = 300000
                myCommand.CommandText = mySP
                Dim param(Nparameters) As MySqlParameter
                For i = 0 To Nparameters - 1
                    If ParamType(i) = "nvarchar" Then
                        param(i) = New MySqlParameter("@" + ParamName(i), MySqlDbType.VarChar, 255, ParameterDirection.Input)
                    ElseIf ParamType(i) = "datetime" Then
                        param(i) = New MySqlParameter("@" + ParamName(i), MySqlDbType.DateTime, 255, ParameterDirection.Input)
                    Else
                        param(i) = New MySqlParameter("@" + ParamName(i), MySqlDbType.Int16, 255, ParameterDirection.Input)
                    End If
                    param(i).Value = ParamValue(i)
                    myCommand.Parameters.Add(param(i))
                Next
                If myCommand.Connection.State = ConnectionState.Closed Then myCommand.Connection.Open()
                myCommand.ExecuteNonQuery()
                myCommand.Connection.Close()
                myCommand.Dispose()
            ElseIf myprovider = "Oracle.ManagedDataAccess.Client" Then
                r = RunSP_Oracle(mySP, Nparameters, ParamName, ParamType, ParamValue, myconstring)
            Else
                Dim myConnection As SqlConnection
                Dim myCommand As New SqlClient.SqlCommand
                myConnection = New SqlConnection(myconstring)
                myCommand.Connection = myConnection
                myCommand.CommandType = CommandType.Text
                myCommand.CommandTimeout = 300000
                myCommand.CommandType = CommandType.StoredProcedure
                myCommand.CommandText = mySP
                Dim param(Nparameters) As SqlParameter
                For i = 0 To Nparameters - 1
                    If ParamType(i) = "nvarchar" Then
                        param(i) = New SqlParameter("@" + ParamName(i), SqlDbType.NVarChar, 255, ParameterDirection.Input)
                    ElseIf ParamType(i) = "datetime" Then
                        param(i) = New SqlParameter("@" + ParamName(i), SqlDbType.DateTime, 255, ParameterDirection.Input)
                    Else
                        param(i) = New SqlParameter("@" + ParamName(i), SqlDbType.Int, 255, ParameterDirection.Input)
                    End If
                    param(i).Value = ParamValue(i)
                    myCommand.Parameters.Add(param(i))
                Next
                If myCommand.Connection.State = ConnectionState.Closed Then myCommand.Connection.Open()
                myCommand.ExecuteNonQuery()
                myCommand.Connection.Close()
                myCommand.Dispose()
            End If
        Catch exc As Exception
            r = exc.Message
            Return r
        End Try
        Return r
    End Function
    Public Function DateToString(dt As DateTime, Optional ByVal userconprv As String = "") As String
        Dim myprovider As String = userconprv
        If myprovider = String.Empty Then _
            myprovider = System.Configuration.ConfigurationManager.ConnectionStrings.Item("mySQLconnection").ProviderName.ToString
        Try
            If myprovider = "MySql.Data.MySqlClient" Then
                Return Format(dt, "yyyy-MM-dd HH:mm:00")
            ElseIf myprovider = "Oracle.ManagedDataAccess.Client" Then

                Return Format(dt, "yyyy-MM-dd")  '"dd-MMM-yy"
            Else
                'Convert.ToDateTime(Your date variable, CultureInfo.CurrentCulture).ToString("yyyy-MM-dd HH:mm:ss.fff")
                Return CDate(dt).ToString("yyyy-MM-dd HH:mm:00")
            End If
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function
    Public Function DateToStringFormat(dt As DateTime, Optional ByVal myprovider As String = "", Optional ByVal datetimeformat As String = "yyyy-MM-dd HH:mm:00") As String
        If myprovider = String.Empty Then _
            myprovider = System.Configuration.ConfigurationManager.ConnectionStrings.Item("mySQLconnection").ProviderName.ToString
        Try
            If myprovider = "MySql.Data.MySqlClient" Then
                Return Format(dt, datetimeformat)
            ElseIf myprovider = "Oracle.ManagedDataAccess.Client" Then
                Return Format(dt, datetimeformat)
            Else
                'Convert.ToDateTime(Your date variable, CultureInfo.CurrentCulture).ToString("yyyy-MM-dd HH:mm:ss.fff")
                Return CDate(dt).ToString(datetimeformat)
            End If
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function
    Public Function DateCurrentFunction() As String
        Dim myprovider As String = System.Configuration.ConfigurationManager.ConnectionStrings.Item("mySQLconnection").ProviderName.ToString
        If myprovider = "MySql.Data.MySqlClient" Then
            Return "CURDATE()"
        ElseIf myprovider = "Oracle.ManagedDataAccess.Client" Then
            Return "CURRENT_DATE"
        Else
            Return "GETDATE()"
        End If
    End Function

    Public Function ConvertFromCacheToMySql(ByVal mSql As String) As String
        Dim er As String = ""
        Dim rSql As String = ""
        Try
            rSql = mSql.Replace("""", "`")
        Catch exc As Exception
            er = exc.Message
            rSql = ""
        End Try
        Return rSql
    End Function
    Public Function ConvertFromCacheToSqlServer(ByVal mSql As String) As String
        Dim er As String = ""
        Dim rSql As String = ""
        Try
            Dim sqlparts() As String = mSql.Split("""")
            If sqlparts.Length > 0 Then
                For i = 0 To sqlparts.Length - 1
                    rSql = rSql & " " & sqlparts(i)
                    If i < sqlparts.Length - 1 Then
                        rSql = rSql & "[" & sqlparts(i + 1) & "]"
                        i = i + 1
                    End If
                Next
            End If
        Catch exc As Exception
            er = exc.Message
            rSql = ""
        End Try
        Return rSql
    End Function
    Public Function ConvertFromMySqlToCache(ByVal mSql As String) As String
        Dim er As String = ""
        Dim rSql As String = ""
        Try
            rSql = mSql.Replace("`", """")
        Catch exc As Exception
            er = exc.Message
            rSql = ""
        End Try

        'TODO fix syntax for classes names and field names (dots)

        Return rSql
    End Function
    Public Function ConvertFromMySqlToSqlServer(ByVal mSql As String) As String
        Dim er As String = ""
        Dim rSql As String = ""
        Try
            Dim sqlparts() As String = mSql.Split("`")
            If sqlparts.Length > 0 Then
                For i = 0 To sqlparts.Length - 1
                    rSql = rSql & " " & sqlparts(i)
                    If i < sqlparts.Length - 1 Then
                        rSql = rSql & "[" & sqlparts(i + 1) & "]"
                        i = i + 1
                    End If
                Next
            End If
        Catch exc As Exception
            er = exc.Message
            rSql = mSql
        End Try
        Return rSql.Trim
    End Function
    Public Function ConvertFromSqlServerToCache(ByVal mSql As String) As String
        Dim er As String = ""
        Dim rSql As String = ""
        Try
            rSql = mSql.Replace("[", """").Replace("]", """")
        Catch exc As Exception
            er = exc.Message
            rSql = ""
        End Try
        Return rSql
    End Function
    Public Function ConvertFromSqlServerToMySql(ByVal mSql As String) As String
        Dim er As String = ""
        Dim rSql As String = ""
        Try
            rSql = mSql.Replace("[", "`").Replace("]", "`")
        Catch exc As Exception
            er = exc.Message
            rSql = ""
        End Try
        Return rSql
    End Function

    Public Function ConvertFromMySqlToOracle(ByVal mSql As String) As String
        Dim er As String = ""
        Dim rSql As String = ""
        Try
            Dim sqlparts() As String = mSql.Split("`")
            If sqlparts.Length > 0 Then
                For i = 0 To sqlparts.Length - 1
                    rSql = rSql & " " & sqlparts(i)
                    If i < sqlparts.Length - 1 Then
                        rSql = rSql & "[" & sqlparts(i + 1) & "]"
                        i = i + 1
                    End If
                Next
            End If
            'remove TOP N
            'rSql = rSql.Replace(" TOP ", " FIRST ")
            Dim t As Integer = rSql.ToUpper.IndexOf(" TOP ")
            If t > 0 AndAlso rSql.Trim.ToUpper.StartsWith("SELECT ") Then
                Dim temp As String = rSql.Substring(t + 5).Trim
                t = temp.IndexOf(" ")
                If t > 0 Then
                    rSql = "SELECT " & temp.Substring(t + 1)
                End If
            End If
        Catch exc As Exception
            er = exc.Message
            rSql = mSql
        End Try
        Return rSql.Trim
    End Function
    Public Function ConvertFromCacheToOracle(ByVal mSql As String) As String
        Dim er As String = ""
        Dim rSql As String = ""
        Try
            rSql = mSql.Replace("""", "")
            'remove TOP N
            Dim t As Integer = rSql.ToUpper.IndexOf(" TOP ")
            If t > 0 AndAlso rSql.Trim.ToUpper.StartsWith("SELECT ") Then
                Dim temp As String = rSql.Substring(t + 5).Trim
                t = temp.IndexOf(" ")
                If t > 0 Then
                    rSql = "SELECT " & temp.Substring(t + 1)
                End If
            End If
        Catch exc As Exception
            er = exc.Message
            rSql = ""
        End Try
        Return rSql
    End Function
    Public Function ConvertFromSqlServerToOracle(ByVal mSql As String) As String
        Dim er As String = ""
        Dim rSql As String = ""
        Try
            rSql = mSql.Replace("[", "").Replace("]", "")
            'remove TOP N
            Dim t As Integer = rSql.ToUpper.IndexOf(" TOP ")
            If t > 0 AndAlso rSql.Trim.ToUpper.StartsWith("SELECT ") Then
                Dim temp As String = rSql.Substring(t + 5).Trim
                t = temp.IndexOf(" ")
                If t > 0 Then
                    rSql = "SELECT " & temp.Substring(t + 1)
                End If
            End If
        Catch exc As Exception
            er = exc.Message
            rSql = ""
        End Try
        Return rSql
    End Function

    Public Function DatabaseConnected(Optional ByVal myconstr As String = "", Optional ByVal myconprv As String = "", Optional ByRef er As String = "", Optional ByRef userODBCdriver As String = "", Optional ByRef userODBCdatabase As String = "", Optional ByRef userODBCdatasource As String = "") As Boolean
        Dim r As Boolean = False
        Dim myconstring, myprovider As String
        Dim dbcase As String = String.Empty
        Try
            If myconstr.Trim = "" Then
                myconstring = System.Configuration.ConfigurationManager.ConnectionStrings.Item("mySQLconnection").ToString
                myprovider = System.Configuration.ConfigurationManager.ConnectionStrings.Item("mySQLconnection").ProviderName.ToString
                If myprovider = "Npgsql" Then dbcase = ConfigurationManager.AppSettings("ourdbcase").ToString
            Else
                If myconprv = "Npgsql" Then
                    If System.Configuration.ConfigurationManager.ConnectionStrings.Item("CSVconnection") IsNot Nothing AndAlso myconstr.Trim.ToUpper = System.Configuration.ConfigurationManager.ConnectionStrings.Item("CSVconnection").ToString.Trim.ToUpper Then
                        dbcase = ConfigurationManager.AppSettings("csvdbcase").ToString
                    ElseIf myconstr.Trim.ToUpper = System.Configuration.ConfigurationManager.ConnectionStrings.Item("mySQLconnection").ToString.Trim.ToUpper Then
                        dbcase = ConfigurationManager.AppSettings("ourdbcase").ToString
                    Else 'userdbcase
                        dbcase = userdbcase
                    End If
                End If
                myconstring = myconstr
                myprovider = myconprv
            End If

            If myprovider = "InterSystems.Data.IRISClient" Then
                r = DatabaseConnected_IRIS(myconstring, er)
            ElseIf myprovider = "InterSystems.Data.CacheClient" Then
                r = DatabaseConnected_Cache(myconstring, er)
            ElseIf myprovider = "MySql.Data.MySqlClient" Then
                Dim myConnection As MySqlConnection
                myConnection = New MySqlConnection(myconstring)
                Try
                    If myConnection.State = ConnectionState.Closed Then myConnection.Open()
                    If myConnection.State = ConnectionState.Open Then r = True
                    myConnection.Close()
                    myConnection.Dispose()
                Catch ex As Exception
                    myConnection.Close()
                    myConnection.Dispose()
                    er = ex.Message
                    r = False
                End Try
            ElseIf myprovider = "Oracle.ManagedDataAccess.Client" Then
                r = DatabaseConnected_Oracle(myconstring, er)
            ElseIf myprovider = "System.Data.Odbc" Then
                myconstring = myconstring.Replace("Password", "Pwd").Replace("User ID", "UID")
                Dim myConnection As System.Data.Odbc.OdbcConnection
                myConnection = New System.Data.Odbc.OdbcConnection(myconstring)
                Try
                    If myConnection.State = ConnectionState.Closed Then myConnection.Open()
                    If myConnection.State = ConnectionState.Open Then r = True
                    userODBCdriver = myConnection.Driver.ToString
                    userODBCdatabase = myConnection.Database.ToString
                    userODBCdatasource = myConnection.DataSource.ToString
                    myConnection.Close()
                    myConnection.Dispose()
                Catch ex As Exception
                    myConnection.Close()
                    myConnection.Dispose()
                    er = ex.Message
                    r = False
                End Try
            ElseIf myprovider = "System.Data.OleDb" Then
                myconstring = ConfigurationManager.AppSettings("ACEOLEDBversion").ToString & "Persist Security Info=True;" & myconstring
                Dim myConnection As System.Data.OleDb.OleDbConnection
                myConnection = New System.Data.OleDb.OleDbConnection(myconstring)
                Try
                    If myConnection.State = ConnectionState.Closed Then myConnection.Open()
                    If myConnection.State = ConnectionState.Open Then r = True
                    userODBCdriver = myConnection.Provider.ToString
                    userODBCdatabase = myConnection.Database.ToString
                    userODBCdatasource = myConnection.DataSource.ToString
                    myConnection.Close()
                    myConnection.Dispose()
                Catch ex As Exception
                    myConnection.Close()
                    myConnection.Dispose()
                    er = ex.Message
                    r = False
                End Try
            ElseIf myprovider = "Npgsql" Then  'PostgreSQL  Npgsql
                myconstring = CorrectConnstringForPostgres(myconstring, dbcase)
                Dim myConnection As Npgsql.NpgsqlConnection
                myConnection = New Npgsql.NpgsqlConnection(myconstring)
                Try
                    If myConnection.State = ConnectionState.Closed Then myConnection.Open()
                    If myConnection.State = ConnectionState.Open Then
                        r = True
                    Else
                        r = False
                    End If
                    myConnection.Close()
                    myConnection.Dispose()
                Catch ex As Exception
                    myConnection.Close()
                    myConnection.Dispose()
                    r = False
                    er = ex.Message
                End Try
            Else
                Dim myConnection As SqlConnection
                myConnection = New SqlConnection(myconstring)
                Try
                    If myConnection.State = ConnectionState.Closed Then myConnection.Open()
                    If myConnection.State = ConnectionState.Open Then
                        r = True
                    Else
                        r = False
                    End If
                    myConnection.Close()
                    myConnection.Dispose()
                Catch ex As Exception
                    myConnection.Close()
                    myConnection.Dispose()
                    r = False
                    er = ex.Message
                End Try
            End If
        Catch exc As Exception
            r = False
            er = exc.Message
        End Try
        Return r
    End Function
    Public Function CorrectConnstringForPostgres(ByVal mconnstr As String, ByVal dbcase As String) As String
        If dbcase = "mix" OrElse dbcase = "doublequoted" Then
            Return mconnstr
        End If
        Dim db As String = GetDataBase(mconnstr, "Npgsql")
        Dim dbcorrected As String = db
        Dim connstrcorrected As String = mconnstr
        If userdbcase = "upper" Then
            dbcorrected = dbcorrected.ToUpper
        Else
            dbcorrected = dbcorrected.ToLower
        End If
        Return connstrcorrected.Replace(db, dbcorrected).Trim
    End Function

    Public Function ExequteSQLquery(ByVal SQLq As String, Optional ByVal myconstr As String = "", Optional ByVal myconprv As String = "") As String
        Dim r As String
        r = "Query executed fine."
        Dim myconstring, myprovider As String
        Dim dbcase As String = String.Empty
        Try
            If myconstr.Trim = "" Then
                myconstring = System.Configuration.ConfigurationManager.ConnectionStrings.Item("mySQLconnection").ToString
                myprovider = System.Configuration.ConfigurationManager.ConnectionStrings.Item("mySQLconnection").ProviderName.ToString
                If myprovider = "Npgsql" Then dbcase = ConfigurationManager.AppSettings("ourdbcase").ToString
            Else
                If myconprv = "Npgsql" Then
                    If System.Configuration.ConfigurationManager.ConnectionStrings.Item("CSVconnection") IsNot Nothing AndAlso myconstr.Trim.ToUpper = System.Configuration.ConfigurationManager.ConnectionStrings.Item("CSVconnection").ToString.Trim.ToUpper Then
                        dbcase = ConfigurationManager.AppSettings("csvdbcase").ToString
                    ElseIf myconstr.Trim.ToUpper = System.Configuration.ConfigurationManager.ConnectionStrings.Item("mySQLconnection").ToString.Trim.ToUpper Then
                        dbcase = ConfigurationManager.AppSettings("ourdbcase").ToString
                    ElseIf myconstr.Trim.ToUpper = System.Configuration.ConfigurationManager.ConnectionStrings.Item("UserSqlConnection").ToString.Trim.ToUpper Then
                        dbcase = ConfigurationManager.AppSettings("userdbcase").ToString
                    Else 'postgres, etc...
                        dbcase = ConfigurationManager.AppSettings("ourdbcase").ToString
                    End If
                End If
                myconstring = myconstr
                myprovider = myconprv
            End If
            If myprovider = "InterSystems.Data.IRISClient" Then
                SQLq = CorrectSQLforCache(SQLq)
                r = ExequteSQLquery_IRIS(SQLq, myconstring)
                If r <> "Query executed fine." Then Return r
            ElseIf myprovider = "InterSystems.Data.CacheClient" Then
                SQLq = CorrectSQLforCache(SQLq)
                r = ExequteSQLquery_Cache(SQLq, myconstring)
                If r <> "Query executed fine." Then Return r
            ElseIf myprovider = "MySql.Data.MySqlClient" Then
                SQLq = CorrectSQLforMySql(SQLq, myconstring)
                Dim myConnection As MySqlConnection
                Dim myCommand As New MySqlCommand
                myConnection = New MySqlConnection(myconstring)
                Try
                    myCommand.Connection = myConnection
                    myCommand.CommandType = CommandType.Text
                    myCommand.CommandTimeout = 300000
                    myCommand.CommandText = SQLq
                    If myCommand.Connection.State = ConnectionState.Closed Then myCommand.Connection.Open()
                    myCommand.ExecuteNonQuery()
                    myCommand.Connection.Close()
                    myCommand.Dispose()
                    myConnection.Dispose()
                Catch ex As Exception
                    myCommand.Connection.Close()
                    myCommand.Dispose()
                    myConnection.Dispose()
                    r = ex.Message
                    Return r
                End Try
            ElseIf myprovider = "Npgsql" Then  'PostgreSQL  Npgsql
                myconstring = CorrectConnstringForPostgres(myconstring, dbcase)
                SQLq = ConvertFromSqlServerToPostgres(SQLq, dbcase, myconstring, myprovider)
                Dim myConnection As Npgsql.NpgsqlConnection
                Dim myCommand As New Npgsql.NpgsqlCommand
                myConnection = New Npgsql.NpgsqlConnection(myconstring)
                Try
                    myCommand.Connection = myConnection
                    myCommand.CommandType = CommandType.Text
                    myCommand.CommandTimeout = 300000
                    myCommand.CommandText = SQLq
                    If myCommand.Connection.State = ConnectionState.Closed Then myCommand.Connection.Open()
                    myCommand.ExecuteNonQuery()
                    myCommand.Connection.Close()
                    myCommand.Dispose()
                    myConnection.Dispose()
                Catch ex As Exception
                    myCommand.Connection.Close()
                    myCommand.Dispose()
                    myConnection.Dispose()
                    r = ex.Message
                    Return r
                End Try
            ElseIf myprovider = "Oracle.ManagedDataAccess.Client" Then
                SQLq = CorrectSQLforOracle(SQLq, myconstring)
                r = ExequteSQLquery_Oracle(SQLq, myconstring)
                If r <> "Query executed fine." Then Return r

            Else
                SQLq = CorrectSQLforSQLServer(SQLq)
                Dim myConnection As SqlConnection
                Dim myCommand As New SqlClient.SqlCommand
                myConnection = New SqlConnection(myconstring)
                Try
                    myCommand.Connection = myConnection
                    myCommand.CommandType = CommandType.Text
                    myCommand.CommandTimeout = 300000
                    myCommand.CommandText = SQLq
                    If myCommand.Connection.State = ConnectionState.Closed Then myCommand.Connection.Open()
                    myCommand.ExecuteNonQuery()
                    myCommand.Connection.Close()
                    myCommand.Dispose()
                    myConnection.Dispose()
                Catch ex As Exception
                    myCommand.Connection.Close()
                    myCommand.Dispose()
                    myConnection.Dispose()
                    r = ex.Message
                    Return r
                End Try
            End If
        Catch exc As Exception
            r = exc.Message
            Return r
        End Try
        Return r
    End Function

    Public Function DataSetFromDataView(ByVal dv As DataView) As DataSet
        '// Clone the structure of the table behind the view
        Dim dtTemp As DataTable = dv.Table.Clone
        dtTemp.Clear()
        ''// Populate the table with rows in the view
        Dim drv As DataRowView
        For Each drv In dv
            dtTemp.ImportRow(drv.Row)
        Next
        Dim dsTemp As DataSet = New DataSet
        '// Add the new table to a DataSet
        dsTemp.Tables.Add(dtTemp)
        Return dsTemp
    End Function
    Public Function myTableFromRow(ByVal dv As DataRow) As DataTable
        '// Clone the structure of the table behind the view
        Dim dtTemp As DataTable = dv.Table.Clone
        dtTemp.Clear()
        dtTemp.ImportRow(dv)
        Return dtTemp
    End Function



    Public Function ExportDataTableToExcel(ByVal dt As DataTable, ByVal Expdir As String, ByVal Expfile As String, Optional ByVal hdr As String = "", Optional ByVal ftr As String = "", Optional ByVal dlm As String = "") As String
        Dim txtline As String = String.Empty
        Dim ret As String = String.Empty
        Dim m, i, j As Integer
        If dt Is Nothing Then
            Return ret
            Exit Function
        End If

        If dlm = "" Then dlm = Chr(9)
        Dim MyFile As StreamWriter = New StreamWriter(Expdir & Expfile)
        Try
            txtline = ""
            MyFile.WriteLine(hdr)
            txtline = ""
            MyFile.WriteLine(txtline)
            m = dt.Columns.Count
            For i = 0 To m - 1
                txtline = txtline & dt.Columns(i).ColumnName
                If i < m - 1 Then txtline = txtline & dlm
            Next
            MyFile.WriteLine(txtline)
            txtline = ""
            For j = 0 To dt.Rows.Count - 1
                txtline = ""
                For i = 0 To m - 1
                    txtline = txtline & """" & cleanText(dt.Rows(j).Item(i).ToString) & """"
                    If i < m - 1 Then txtline = txtline & dlm
                Next
                If Trim(txtline) <> "" Then
                    MyFile.WriteLine(txtline)
                End If
            Next
            txtline = ""
            MyFile.WriteLine(txtline)
            MyFile.WriteLine(ftr)
            MyFile.Flush()
            MyFile.Close()
            MyFile = Nothing
            ret = Expdir & Expfile
        Catch ex As Exception
            ret = "Error creating Excel file" & ex.Message
            MyFile.Close()
            MyFile = Nothing
        End Try
        Return ret
    End Function
    Public Function ExportToExcel(ByVal dt As DataTable, ByVal Expdir As String, ByVal Expfile As String, Optional ByVal hdr As String = "", Optional ByVal ftr As String = "", Optional ByVal dlm As String = "") As String
        Dim txtline As String = String.Empty
        Dim ret As String = String.Empty
        Dim m, i, j As Integer
        If dt Is Nothing Then
            Return ret
            Exit Function
        End If

        If dlm = "" Then dlm = Chr(9)
        Dim MyFile As StreamWriter = New StreamWriter(Expdir & Expfile)
        Try
            txtline = ""
            MyFile.WriteLine(hdr)
            txtline = ""
            MyFile.WriteLine(txtline)
            m = dt.Columns.Count
            For i = 0 To m - 1
                txtline = txtline & dt.Columns(i).ColumnName
                If i < m - 1 Then txtline = txtline & dlm
            Next
            MyFile.WriteLine(txtline)
            txtline = ""
            For j = 0 To dt.Rows.Count - 1
                txtline = ""
                For i = 0 To m - 1
                    txtline = txtline & dt.Rows(j).Item(i).ToString
                    If i < m - 1 Then txtline = txtline & dlm
                Next
                If Trim(txtline) <> "" Then
                    MyFile.WriteLine(txtline)
                End If
            Next
            txtline = ""
            MyFile.WriteLine(txtline)
            MyFile.WriteLine(ftr)
            MyFile.Flush()
            MyFile.Close()
            MyFile = Nothing
            ret = Expdir & Expfile
        Catch ex As Exception
            ret = "Error creating Excel file" & ex.Message
            MyFile.Close()
            MyFile = Nothing
        End Try
        Return ret
    End Function

    Public Function ExportToCSV(ByVal dt As DataTable, ByVal Expdir As String, ByVal csvfile As String, ByVal delimtr As String, Optional ByVal hdr As String = "", Optional ByVal ftr As String = "") As String
        Dim txtline As String
        Dim m, i, j As Integer
        Dim fso = CreateObject("Scripting.FileSystemObject")
        'Dim applpath As String = System.AppDomain.CurrentDomain.BaseDirectory()
        Dim MyFile As Object
        On Error GoTo ErrMsg
        MyFile = fso.CreateTextFile(Expdir & csvfile, True)
        On Error GoTo ErrMsg
        txtline = ""
        MyFile.WriteLine(hdr)
        txtline = ""
        MyFile.WriteLine(txtline)
        m = dt.Columns.Count
        For i = 0 To m - 1
            txtline = txtline & dt.Columns(i).ColumnName
            If i < m - 1 Then txtline = txtline & delimtr
        Next
        MyFile.WriteLine(txtline)
        txtline = ""
        For j = 0 To dt.Rows.Count - 1
            txtline = ""
            For i = 0 To m - 1
                txtline = txtline & dt.Rows(j).Item(i).ToString
                If i < m - 1 Then txtline = txtline & delimtr
            Next
            If Trim(txtline) <> "" Then
                MyFile.WriteLine(txtline)
            End If
        Next
        txtline = ""
        MyFile.WriteLine(txtline)
        MyFile.WriteLine(ftr)
        MyFile.Close()
        MyFile = Nothing
        fso = Nothing
        ExportToCSV = csvfile
        Exit Function
ErrMsg:
        MyFile.Close()
        MyFile = Nothing
        fso = Nothing
        ExportToCSV = "Error creating Excel file"
    End Function

    Public Function AddRowIntoSQLtable(ByVal mRow As DataRow, ByVal SQLq As String, ByVal connstr As String) As String
        'get updatable DataTable
        Dim cmdBuilder As SqlCommandBuilder
        Dim cmdTmp As New SqlCommand
        cmdTmp.Connection.ConnectionString = connstr
        If cmdTmp.Connection.State = ConnectionState.Closed Then cmdTmp.Connection.Open()
        cmdTmp.CommandType = CommandType.Text
        cmdTmp.CommandText = SQLq
        Dim rs = New SqlClient.SqlDataAdapter(cmdTmp)
        cmdBuilder = New SqlCommandBuilder(rs)
        Dim myTable = New DataTable
        rs.Fill(myTable)
        'add row
        On Error GoTo ErrMsg
        myTable.ImportRow(mRow)
        rs.Update(myTable)
        AddRowIntoSQLtable = "The row has been inserted into the table."
        Exit Function
ErrMsg:
        AddRowIntoSQLtable = "Error inserting the row into the table."

    End Function
    Public Function CreateRowInHTMLtable(ByVal mTable As HtmlTable) As String
        'create new row in mTable
        Dim j As Integer
        Dim AddRow As HtmlTableRow = New HtmlTableRow()
        For j = 0 To mTable.Rows(0).Cells.Count - 1
            Dim cell As HtmlTableCell
            cell = New HtmlTableCell()
            cell.Controls.Add(New LiteralControl("row0, " & "column " & j.ToString()))
            AddRow.Cells.Add(cell)
        Next j
        mTable.Rows.Add(AddRow)
        Return "done"
    End Function
    Public Function CreateNewRowInHtmlTable(ByVal mTable As HtmlTable) As String
        Try
            'create new row in mTable
            Dim AddRow As HtmlTableRow = New HtmlTableRow()
            For j = 0 To mTable.Rows(0).Cells.Count - 1
                Dim cell As HtmlTableCell
                cell = New HtmlTableCell()
                cell.Controls.Add(New LiteralControl("row0, " & "column " & j.ToString()))
                AddRow.Cells.Add(cell)
            Next j
            mTable.Rows.Add(AddRow)
            Return "Record created"
        Catch ex As Exception
            Return "ERROR!! " & ex.Message
        End Try
    End Function
    Public Function AddRowIntoHTMLtable(ByVal mRow As DataRow, ByVal mTable As HtmlTable) As String
        Dim j As Integer
        Try


            If mRow.RowState = DataRowState.Deleted Then Return "Error inserting the row into the table."
            'create new row in mTable
            Dim AddRow As HtmlTableRow = New HtmlTableRow()
            For j = 0 To mTable.Rows(0).Cells.Count - 1
                Dim cell As HtmlTableCell
                cell = New HtmlTableCell()
                cell.Controls.Add(New LiteralControl("row0, " & "column " & j.ToString()))
                AddRow.Cells.Add(cell)
            Next j

            'assign values for row's fields
            For j = 0 To Min(mTable.Rows(0).Cells.Count - 1, mRow.ItemArray.Length - 1)
                AddRow.Cells(j).InnerText = mRow(j).ToString
            Next
            mTable.Rows.Add(AddRow)
            AddRowIntoHTMLtable = "Row has been inserted into the table."
            Exit Function
        Catch ex As Exception
            AddRowIntoHTMLtable = "Error inserting the row into the table."
        End Try
    End Function
    Public Function AddArrayIntoHTMLtable(ByVal mRow As Array, ByVal mTable As HtmlTable) As String
        Dim i, j As Integer

        'create new row in mTable
        Dim AddRow As HtmlTableRow = New HtmlTableRow()
        For j = 0 To mTable.Rows(0).Cells.Count - 1
            Dim cell As HtmlTableCell
            cell = New HtmlTableCell()
            cell.Controls.Add(New LiteralControl("row " & i.ToString() & ", " & "column " & j.ToString()))
            AddRow.Cells.Add(cell)
        Next j

        'assign values for row's fields
        For j = 1 To mTable.Rows(0).Cells.Count - 2
            AddRow.Cells(j).InnerText = mRow(j - 1)
        Next
        mTable.Rows.Add(AddRow)

        AddArrayIntoHTMLtable = "Row has been inserted into the table."
        Exit Function
ErrMsg:
        AddArrayIntoHTMLtable = "Error inserting the row into the table."

    End Function


    Public Function TblFieldIsNumeric(ByVal tbl As String, ByVal fld As String, Optional ByVal userconstr As String = "", Optional ByVal userconprv As String = "") As Boolean
        Dim num As Boolean = Nothing
        Dim i As Integer
        Dim er As String = String.Empty
        Dim typ As String = String.Empty
        If tbl.Trim = "" OrElse fld.Trim = "" Then
            Return num
            Exit Function
        End If
        Dim dvt As DataTable = Nothing
        Dim myprovider As String = userconprv
        Dim myconstring As String = userconstr
        If userconstr = "" Then
            myconstring = System.Configuration.ConfigurationManager.ConnectionStrings.Item("mySQLconnection").ToString
            myprovider = System.Configuration.ConfigurationManager.ConnectionStrings.Item("mySQLconnection").ProviderName.ToString
        End If
        Dim sqls As String
        sqls = String.Empty
        If myprovider.StartsWith("InterSystems.Data.") Then
            sqls = "SELECT Name AS COLUMN_NAME, Type AS DATA_TYPE FROM %Dictionary.PropertyDefinition WHERE UCASE(parent) = UCASE('" & tbl & "')"
        ElseIf myprovider = "MySql.Data.MySqlClient" Then
            Dim db As String = GetDataBase(myconstring, myprovider)
            sqls = "SELECT COLUMN_NAME, DATA_TYPE FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = '" & tbl & "' AND TABLE_SCHEMA='" & db.ToLower & "'"
        ElseIf myprovider = "Oracle.ManagedDataAccess.Client" Then
            sqls = "SELECT * FROM all_tab_cols WHERE UPPER(TABLE_NAME) = UPPER('" & tbl & "')"

        ElseIf myprovider = "Npgsql" Then  'PostgreSQL  Npgsql
            Dim db As String = GetDataBase(myconstring, myprovider)
            sqls = "SELECT COLUMN_NAME, DATA_TYPE FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = '" & tbl & "' AND TABLE_SCHEMA ='public' AND LOWER(TABLE_CATALOG) ='" & db.ToLower & "'"

        Else
            sqls = "SELECT COLUMN_NAME, DATA_TYPE FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = '" & tbl & "'"
        End If
        If myprovider <> "System.Data.Odbc" AndAlso myprovider <> "System.Data.OleDb" Then
            dvt = mRecords(sqls, er, myconstring, myprovider).Table
        End If

        For i = 0 To dvt.Rows.Count - 1
            If (dvt.Rows(i)("COLUMN_NAME") = fld) Then
                typ = dvt.Rows(i)("DATA_TYPE").ToString.Replace("%", "").ToUpper
                If typ.ToUpper.Contains("INTEGER") OrElse typ.Contains("SmallInt".ToUpper) OrElse typ = "int".ToUpper OrElse typ.Contains("smallint".ToUpper) OrElse typ.Contains("float".ToUpper) OrElse typ.Contains("bigint".ToUpper) OrElse typ.ToUpper.Contains("DECIMAL") OrElse typ.ToUpper.Contains("NUMERIC") OrElse typ.ToUpper.Contains("DOUBLE") OrElse typ.ToUpper.Contains("BINARY") OrElse typ.ToUpper.Contains("NUMBER") Then
                    num = True
                Else
                    num = False
                End If
                Return num
                Exit Function
            End If
        Next
        Return num
    End Function
    Public Function GetListOfTableFields(ByVal tbl As String, ByVal userconstr As String, ByVal userconprv As String, Optional ByRef er As String = "") As DataView
        Dim ert As String = String.Empty
        Dim dv As DataView = Nothing
        Dim i As Integer
        Dim dtb As New DataTable
        Dim col As DataColumn
        col = New DataColumn
        col.DataType = System.Type.GetType("System.String")
        col.ColumnName = "COLUMN_NAME"
        dtb.Columns.Add(col)
        col = New DataColumn
        col.DataType = System.Type.GetType("System.String")
        col.ColumnName = "DATA_TYPE"
        dtb.Columns.Add(col)
        Dim dt As DataTable
        Try
            Dim sqls As String = "SELECT TOP 1 * FROM [" & tbl & "]"

            Dim dvt As DataView = mRecords(sqls, er, userconstr, userconprv)
            If er = "" AndAlso dvt IsNot Nothing Then
                dt = dvt.Table
                If dt.Rows.Count = 0 Then
                    er = "Not data in the table " & tbl
                End If
                For i = 0 To dt.Columns.Count - 1
                    Dim Row As DataRow = dtb.NewRow()
                    Row("COLUMN_NAME") = dt.Columns(i).Caption
                    Row("DATA_TYPE") = dt.Columns(i).DataType
                    dtb.Rows.Add(Row)
                Next
                dv = dtb.DefaultView
            Else

                dv = GetListOfTableColumns(tbl, userconstr, userconprv, ert)

            End If

        Catch ex As Exception
            er = ex.Message
            dv = GetListOfTableColumns(tbl, userconstr, userconprv, ert)
        End Try
        Return dv
    End Function
    Public Function GetListOfTableColumns(ByVal tbl As String, Optional ByVal userconstr As String = "", Optional ByVal userconprv As String = "", Optional ByRef err As String = "") As DataView
        Dim dv As DataView = Nothing
        Dim myprovider, sqls, strConnect As String
        sqls = String.Empty
        If userconstr = "" Then
            strConnect = System.Configuration.ConfigurationManager.ConnectionStrings.Item("mySQLconnection").ToString
            myprovider = System.Configuration.ConfigurationManager.ConnectionStrings.Item("mySQLconnection").ProviderName.ToString
        Else
            strConnect = userconstr
            myprovider = userconprv
        End If
        If myprovider.StartsWith("InterSystems.Data.") Then
            'TODO fix name of class with dots
            tbl = tbl.Replace("_", ".")
            If tbl.ToUpper.StartsWith("INFORMATION.SCHEMA") Then
                sqls = "Select SqlFieldName As COLUMN_NAME, Type AS DATA_TYPE FROM %Dictionary.PropertyDefinition WHERE UCASE(parent) = UCASE('" & tbl & "') AND Cardinality Is NULL  AND Inverse Is NULL AND Private=0 AND Transient=0 AND NOT SqlFieldName IS NULL"
            Else
                If Not tbl.Contains(".") Then
                    tbl = "UserData" & "." & tbl
                End If
                sqls = "Select Name As COLUMN_NAME, Type AS DATA_TYPE FROM %Dictionary.PropertyDefinition WHERE UCASE(parent) = UCASE('" & tbl & "') AND Cardinality Is NULL  AND Inverse Is NULL AND Private=0 AND Transient=0"
            End If
        ElseIf myprovider = "MySql.Data.MySqlClient" Then
            Dim db As String = GetDataBase(strConnect, myprovider)
            sqls = "SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_SCHEMA ='" & db.ToLower & "' AND TABLE_NAME = '" & tbl & "'"

        ElseIf myprovider = "Oracle.ManagedDataAccess.Client" Then
            'Dim db As String = GetDataBase(strConnect,myprovider)
            sqls = "SELECT * FROM all_tab_cols WHERE UPPER(TABLE_NAME) = UPPER('" & tbl & "')"

        ElseIf myprovider = "System.Data.Odbc" OrElse myprovider = "System.Data.OleDb" Then

            dv = GetListOfTableFields(tbl, userconstr, userconprv)
            Return dv

        ElseIf myprovider = "Npgsql" Then  'PostgreSQL  Npgsql
            Dim db As String = GetDataBase(strConnect, myprovider)
            sqls = "SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_SCHEMA ='public' AND LOWER(TABLE_CATALOG) ='" & db.ToLower & "' AND TABLE_NAME = '" & tbl & "'"

        Else
            sqls = "SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = '" & tbl & "'"
        End If
        dv = mRecords(sqls, err, userconstr, userconprv)
        If myprovider.StartsWith("InterSystems.Data") AndAlso err = "" Then
            Dim dt As DataTable = dv.Table
            Dim NewRow As Object() = New Object(0) {"ID"}
            dt.BeginLoadData()
            dt.LoadDataRow(NewRow, True)
            dt.EndLoadData()
        End If
        Return dv
    End Function
    Public Function TblFieldIsDateTime(ByVal tbl As String, ByVal fld As String, Optional ByVal userconstr As String = "", Optional ByVal userconprv As String = "") As Boolean
        Dim num As Boolean = Nothing
        Dim i As Integer
        Dim er As String = String.Empty
        Dim typ As String = String.Empty
        If tbl.Trim = "" OrElse fld.Trim = "" Then
            Return num
            Exit Function
        End If
        Dim dvt As DataTable = Nothing
        Dim myprovider As String = userconprv
        Dim myconstring As String = userconstr
        If userconstr = "" Then
            myconstring = System.Configuration.ConfigurationManager.ConnectionStrings.Item("mySQLconnection").ToString
            myprovider = System.Configuration.ConfigurationManager.ConnectionStrings.Item("mySQLconnection").ProviderName.ToString
        End If
        Dim sqls As String
        sqls = String.Empty
        If myprovider.StartsWith("InterSystems.Data.") Then
            sqls = "SELECT Name AS COLUMN_NAME, Type AS DATA_TYPE FROM %Dictionary.PropertyDefinition WHERE parent = '" & tbl & "'"
        ElseIf myprovider = "MySql.Data.MySqlClient" Then
            Dim db As String = GetDataBase(myconstring, myprovider)
            sqls = "SELECT COLUMN_NAME, DATA_TYPE FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = '" & tbl & "' AND TABLE_SCHEMA='" & db.ToLower & "'"
        ElseIf myprovider = "Oracle.ManagedDataAccess.Client" Then
            'for Oracle.ManagedDataAccess.Client
            sqls = "SELECT * FROM all_tab_cols WHERE UPPER(TABLE_NAME) = UPPER('" & tbl & "')"

        ElseIf myprovider = "Npgsql" Then  'PostgreSQL  Npgsql
            Dim db As String = GetDataBase(myconstring, myprovider)
            sqls = "SELECT COLUMN_NAME, DATA_TYPE FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = '" & tbl & "' AND TABLE_SCHEMA ='public' AND LOWER(TABLE_CATALOG) ='" & db.ToLower & "'"

        Else
            sqls = "SELECT COLUMN_NAME, DATA_TYPE FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = '" & tbl & "'"
        End If
        If myprovider <> "System.Data.Odbc" AndAlso myprovider <> "System.Data.OleDb" Then
            dvt = mRecords(sqls, er, myconstring, myprovider).Table
        End If

        For i = 0 To dvt.Rows.Count - 1
            If (dvt.Rows(i)("COLUMN_NAME") = fld) Then
                typ = dvt.Rows(i)("DATA_TYPE").ToString.Replace("%", "").ToUpper
                If typ.Contains("DATE") OrElse typ.Contains("Time".ToUpper) Then
                    num = True
                Else
                    num = False
                End If
                Return num
                Exit Function
            End If
        Next
        Return num
    End Function
    Public Function TblSQLqueryFieldIsNumeric(ByVal rep As String, ByVal fld As String, Optional ByVal userconstr As String = "", Optional ByVal userconprv As String = "") As Boolean
        Dim num As Boolean = Nothing
        'Dim i As Integer
        Dim typ As String = String.Empty
        If rep.Trim = "" OrElse fld.Trim = "" Then
            Return num
            Exit Function
        End If
        Dim dvf As DataView = mRecords("SELECT * FROM OURReportSQLquery WHERE (ReportID='" & rep & "' AND Doing='SELECT' AND Tbl1Fld1='" & fld & "')")
        If dvf Is Nothing OrElse dvf.Count = 0 OrElse dvf.Table.Rows.Count = 0 Then
            Return num
            Exit Function
        End If
        If TblFieldIsNumeric(dvf.Table.Rows(0)("Tbl1").ToString, dvf.Table.Rows(0)("Tbl1Fld1").ToString, userconstr, userconprv) Then
            num = True
        Else
            num = False
        End If
        Return num
    End Function


    Public Function CreateInitialClass(Optional ByVal myconstr As String = "", Optional ByVal myconprv As String = "") As String
        Dim rt As String = String.Empty
        Try
            Dim sqlq As String = String.Empty
            sqlq = "DROP TABLE OUR.INIT"
            rt = ExequteSQLquery(sqlq, myconstr, myconprv)
            sqlq = "CREATE TABLE OUR.INIT (STAT VARCHAR(30), DJ VARCHAR(30), NODE VARCHAR(1000), VALUE VARCHAR(32500))"
            rt = ExequteSQLquery(sqlq, myconstr, myconprv)
            sqlq = "CREATE PROCEDURE IMPORTCLASSFROMXMLFILE(IN ClassPath VARCHAR(100)) FOR OUR.INIT LANGUAGE OBJECTSCRIPT { D $system.OBJ.Load(ClassPath,""c-lfr-d"")  q 1}"
            rt = ExequteSQLquery(sqlq, myconstr, myconprv)
            sqlq = "CREATE PROCEDURE BUILDCLASSFROMSTRING(IN ClassText VARCHAR(32000)) FOR OUR.INIT LANGUAGE OBJECTSCRIPT { set stream=##class(%GlobalBinaryStream).%New()  d stream.Write(ClassText)	D $system.OBJ.LoadStream(stream,""c-lfr-d"") q 1 }"
            rt = ExequteSQLquery(sqlq, myconstr, myconprv)

            sqlq = "CREATE PROCEDURE BUILDROUTINE(IN RoutineName VARCHAR(100), RoutineText VARCHAR(32000)) FOR OUR.INIT LANGUAGE OBJECTSCRIPT { S routine=##class(%Routine).%New(RoutineName) "
            'sqlq = sqlq & " ;create routine RoutineName from RoutineText string "
            sqlq = sqlq & " S rtext=RoutineText	"
            sqlq = sqlq & " F I=1:1:$L(rtext,$c(13,10)) "
            sqlq = sqlq & " { s textline=$P(rtext,$c(13,10),I) "
            sqlq = sqlq & " D routine.WriteLine(textline) } "
            sqlq = sqlq & " D routine.WriteLine($c(13,10)) "
            sqlq = sqlq & " D routine.Save() "
            'sqlq = sqlq & " ; Silent compile ; -dk : -Don't display and Keep source code"
            sqlq = sqlq & " S sc=routine.Compile("" - dk"") "
            sqlq = sqlq & " IF 'sc K err D $system.Status.DecomposeStatus(sc,.err) "
            sqlq = sqlq & " D routine.%Close() q 1 }"
            rt = ExequteSQLquery(sqlq, myconstr, myconprv)

            sqlq = "CREATE PROCEDURE GETGLOBALNODEDATA(IN glb VARCHAR(1000)) FOR OUR.INIT RESULT SETS LANGUAGE OBJECTSCRIPT { "
            sqlq = sqlq & " s value=$G(@glb)  s tag=+$J "
            sqlq = sqlq & " s rst = ##class(%ResultSet).%New() "
            sqlq = sqlq & " SET sc=rst.Prepare(""DELETE FROM OUR.INIT WHERE DJ='""_tag_""'"") "
            sqlq = sqlq & " SET sc=rst.Execute() "
            sqlq = sqlq & " s clm = ##class(OUR.INIT).%New()	"
            sqlq = sqlq & " SET clm.STAT=tag "
            sqlq = sqlq & " SET clm.NODE=glb "
            sqlq = sqlq & " SET clm.VALUE=value "
            sqlq = sqlq & " SET clm.DJ=tag "
            sqlq = sqlq & " s sc=clm.%Save() "
            sqlq = sqlq & " s sqls=""Select * FROM OUR.INIT WHERE DJ='""_tag_""'"" "
            sqlq = sqlq & " s rset = ##class(%ResultSet).%New() "
            sqlq = sqlq & " SET sc=rset.Prepare(sqls) "
            sqlq = sqlq & " SET sc=rset.Execute() "
            sqlq = sqlq & " If '$isobject($Get(%sqlcontext)) {set %sqlcontext = ##class(%Library.ProcedureContext).%New()} "
            sqlq = sqlq & " do %sqlcontext.AddResultSet(rset) }"
            'sqlq = sqlq & " q 1 }"
            ExequteSQLquery(sqlq, myconstr, myconprv)
            If rt = "Query executed fine." Then
                rt = "Class OUR.INIT created"
            Else
                Return rt
            End If

            'create Fred's class
            'inc
            Dim applpath As String = System.AppDomain.CurrentDomain.BaseDirectory()
            Dim routinetext As String = GetTextFromFile(applpath & "include\", "GenASys_SQL_inc", ".xml")
            rt = CreateRoutine("OUR.GenASys.SQL.inc", routinetext, myconstr, myconprv)
            Dim classtext As String = GetTextFromFile(applpath & "include\", "OUR.Utils.StoredProcs", ".CLS")
            rt = CreateClass(classtext, myconstr, myconprv)

        Catch ex As Exception
            Return ex.Message
        End Try
        Return rt
    End Function
    Private Function CreateClass(ByVal ClassText As String, Optional ByVal myconstr As String = "", Optional ByVal myconprv As String = "") As String
        'Only for Cache !!
        Dim ret As String = String.Empty
        Try
            Dim myprovider As String
            If myconstr = "" Then
                myprovider = System.Configuration.ConfigurationManager.ConnectionStrings.Item("mySQLconnection").ProviderName.ToString
            Else
                myprovider = myconprv
            End If
            If myprovider = "InterSystems.Data.CacheClient" AndAlso ClassText <> String.Empty Then
                Dim StoredProcName As String = String.Empty
                Dim ParamName(0) As String
                Dim ParamType(0) As String
                Dim ParamValue(0) As String
                'make params
                ReDim ParamName(0)
                ReDim ParamType(0)
                ReDim ParamValue(0)
                ParamName(0) = "ClassText"
                ParamValue(0) = ClassText
                StoredProcName = "OUR.BUILDCLASSFROMSTRING"  'StorProc in INIT class has [ SqlName = BUILDCLASSFROMSTRING, SqlProc ] 
                ret = RunSP(StoredProcName, 1, ParamName, ParamType, ParamValue)
                Return ret
            End If
        Catch ex As Exception
            Return ex.Message
        End Try
        Return ret
    End Function

    Public Function ColumnExists(tbl As String, column As String, Optional myconstr As String = "", Optional ByVal myconprv As String = "", Optional ByRef er As String = "") As Boolean
        Dim ret As Boolean = False
        Dim sqlq As String = String.Empty
        Dim err As String = String.Empty

        If TableExists(tbl, myconstr, myconprv, err) Then
            Try
                If myconstr = "" Then
                    myconprv = System.Configuration.ConfigurationManager.ConnectionStrings.Item("mySQLconnection").ProviderName.ToString
                    myconstr = System.Configuration.ConfigurationManager.ConnectionStrings.Item("mySQLconnection").ToString
                End If
                If myconprv = "MySql.Data.MySqlClient" Then
                    Dim db As String = GetDataBase(myconstr, myconprv)
                    sqlq = "SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_SCHEMA='" & db.ToLower & "' And TABLE_NAME='" & tbl.ToLower & "' And COLUMN_NAME='" & column & "'"
                ElseIf myconprv = "Oracle.ManagedDataAccess.Client" Then
                    If Not column.Contains("""") Then
                        column = column.ToUpper
                    Else
                        column = column.Replace("""", "")
                    End If
                    sqlq = "SELECT * FROM all_tab_cols WHERE UPPER(TABLE_NAME) = UPPER('" & tbl & "') AND UPPER(COLUMN_NAME)='" & column.ToUpper & "'"

                ElseIf myconprv.StartsWith("InterSystems.Data.") Then
                    sqlq = "Select Name As COLUMN_NAME FROM %Dictionary.PropertyDefinition WHERE parent = '" & tbl.ToUpper & "' AND Cardinality Is NULL AND Name = '" & column & "'"

                ElseIf myconprv = "Npgsql" Then  'PostgreSQL  Npgsq
                    Dim db As String = GetDataBase(myconstr, myconprv)
                    sqlq = "SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_SCHEMA ='public' AND LOWER(TABLE_CATALOG) ='" & db.ToLower & "' And TABLE_NAME='" & tbl.ToLower & "' And COLUMN_NAME='" & column & "'"

                Else 'sql server
                    sqlq = "SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='" & tbl & "' AND COLUMN_NAME='" & column & "'"
                End If
                ret = HasRecords(sqlq, myconstr, myconprv)
            Catch ex As Exception
                er = ex.Message
            End Try
        End If
        If err <> String.Empty Then _
            er = err
        Return ret
    End Function
    Public Function TableExists(ByVal tbl As String, Optional ByVal myconstr As String = "", Optional ByVal myconprv As String = "", Optional ByRef er As String = "") As Boolean
        Dim rt As Boolean = False

        Dim sqlq As String = String.Empty
        Try
            If myconstr = "" Then
                myconprv = System.Configuration.ConfigurationManager.ConnectionStrings.Item("mySQLconnection").ProviderName.ToString
                myconstr = System.Configuration.ConfigurationManager.ConnectionStrings.Item("mySQLconnection").ToString
            End If

            If myconprv = "MySql.Data.MySqlClient" Then
                Dim db As String = GetDataBase(myconstr, myconprv)
                sqlq = "SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_SCHEMA='" & db.ToLower & "' AND TABLE_NAME='" & tbl.ToLower & "'"

            ElseIf myconprv = "Oracle.ManagedDataAccess.Client" Then
                Dim db As String = GetDataBase(myconstr, myconprv)
                sqlq = "SELECT * FROM all_tables WHERE UPPER(TABLE_NAME) = UPPER('" & tbl & "')"

                'ElseIf myconprv = "System.Data.Odbc" Then
                '    'ODBC
                '    Dim dtsh As New DataTable
                '    Dim dv As DataView = GetListOfUserTables(True, myconstr, myconprv)
                '    dv.RowFilter = "TABLE_NAME = '" & tbl & "' OR (TABLE_NAME LIKE '%." & tbl & "')"
                '    ' loop if dv.totable.rows.count>1   TABLE_NAME = TABLE_SCHEM & '." & tbl & "'
                '    If dv.ToTable.DefaultView.Table.Rows.Count = 1 Then
                '        Return True
                '    ElseIf dv.ToTable.DefaultView.Table.Rows.Count > 1 Then
                '        For i = 0 To dv.ToTable.DefaultView.Table.Rows.Count = 1
                '            If dv.ToTable.DefaultView.Table.Rows("TABLE_NAME").ToString.Trim = dv.ToTable.DefaultView.Table.Rows("TABLE_SCHEM").ToString.Trim & "." & tbl Then
                '                Return True
                '            End If
                '        Next
                '        Return False
                '    Else
                '        Return False
                '    End If
                'ElseIf myconprv = "System.Data.OleDb" Then
                '    'OleDb
                '    Dim dtsh As New DataTable
                '    Dim dv As DataView = GetListOfUserTables(True, myconstr, myconprv)
                '    dv.RowFilter = "TABLE_NAME = '" & tbl & "' OR (TABLE_NAME LIKE '%." & tbl & "')"
                '    ' loop if dv.totable.rows.count>1   TABLE_NAME = TABLE_SCHEM & '." & tbl & "'
                '    If dv.ToTable.DefaultView.Table.Rows.Count = 1 Then
                '        Return True
                '    ElseIf dv.ToTable.DefaultView.Table.Rows.Count > 1 Then
                '        For i = 0 To dv.ToTable.DefaultView.Table.Rows.Count = 1
                '            If dv.ToTable.DefaultView.Table.Rows("TABLE_NAME").ToString.Trim = dv.ToTable.DefaultView.Table.Rows("TABLE_SCHEM").ToString.Trim & "." & tbl Then
                '                Return True
                '            End If
                '        Next
                '        Return False
                '    Else
                '        Return False
                '    End If
            ElseIf myconprv.StartsWith("InterSystems.Data.") Then
                If Not tbl.Contains(".") AndAlso Not tbl.ToUpper.StartsWith("OUR") Then tbl = "userdata." & tbl
                sqlq = "Select ID As TABLE_NAME FROM %Dictionary.ClassDefinition WHERE Not (ID %STARTSWITH '%') AND UCASE(ID)='" & tbl.ToUpper & "'"

            ElseIf myconprv = "Npgsql" Then  'PostgreSQL  Npgsql
                Dim db As String = GetDataBase(myconstr, myconprv)
                sqlq = "SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_SCHEMA ='public' AND LOWER(TABLE_CATALOG) ='" & db.ToLower & "' AND TABLE_NAME='" & tbl & "'"

            Else
                sqlq = "SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME='" & tbl & "'"
            End If
            Dim myView As DataView = mRecords(sqlq, er, myconstr, myconprv)
            rt = HasRecords(sqlq, myconstr, myconprv)
        Catch ex As Exception
            er = ex.Message
        End Try
        Return rt
    End Function

    Public Function DatabaseExist(ByVal dbnm As String, Optional ByVal myconstr As String = "", Optional ByVal myconprv As String = "", Optional ByRef er As String = "") As Boolean
        Dim ret As Boolean = True
        Dim dv As DataView = Nothing
        Dim sq As String = String.Empty
        Dim myconstring, myprovider As String
        Try
            If myconstr.Trim = "" Then
                myconstring = System.Configuration.ConfigurationManager.ConnectionStrings.Item("mySQLconnection").ToString
                myprovider = System.Configuration.ConfigurationManager.ConnectionStrings.Item("mySQLconnection").ProviderName.ToString
            Else
                myconstring = myconstr
                myprovider = myconprv
            End If

            If myprovider.StartsWith("InterSystems.Data.") Then  ' = "InterSystems.Data.IRISClient" Then 'ElseIf myprovider = "InterSystems.Data.CacheClient" Then
                ret = False
                sq = "CALL %SYS.Namespace_List()"
                dv = mRecords(sq, er, myconstring, myprovider)
                If (dv Is Nothing OrElse dv.Table Is Nothing OrElse dv.Table.Rows.Count = 0) OrElse er <> "" Then
                    ret = False
                Else
                    For i = 0 To dv.Table.Rows.Count - 1
                        If dv.Table.Rows(i)("Nsp").ToString.ToUpper = dbnm.ToUpper Then
                            ret = True
                            Exit For
                        End If
                    Next
                End If

            ElseIf myprovider = "Oracle.ManagedDataAccess.Client" Then
                sq = "SELECT username AS schema_name FROM all_users WHERE UCASE(username)='" & dbnm.ToUpper & "'"
                dv = mRecords(sq, er, myconstring, myprovider)
                If (dv Is Nothing OrElse dv.Table Is Nothing OrElse dv.Table.Rows.Count = 0) OrElse er <> "" Then
                    ret = False
                End If

            ElseIf myprovider = "MySql.Data.MySqlClient" Then
                sq = "SELECT schema_name FROM information_schema.schemata WHERE UCASE(schema_name)='" & dbnm.ToUpper & "';"
                dv = mRecords(sq, er, myconstring, myprovider)
                If (dv Is Nothing OrElse dv.Table Is Nothing OrElse dv.Table.Rows.Count = 0) OrElse er <> "" Then
                    ret = False
                End If

            ElseIf myprovider = "System.Data.SqlClient" Then
                sq = "Select NAME FROM sys.databases WHERE NAME='" & dbnm & "'"
                If Not HasRecords(sq, myconstr, myconprv) Then
                    ret = False
                End If

            ElseIf myprovider = "Npgsql" Then  'PostgreSQL  Npgsql
                sq = "SELECT schema_name FROM information_schema.schemata WHERE SCHEMA_NAME ='public' AND LOWER(CATALOG_NAME) ='" & dbnm.ToLower & "';"
                dv = mRecords(sq, er, myconstring, myprovider)
                If (dv Is Nothing OrElse dv.Table Is Nothing OrElse dv.Table.Rows.Count = 0) OrElse er <> "" Then
                    ret = False
                End If
            End If
        Catch ex As Exception
            er = ex.Message
            ret = True
        End Try
        Return ret
    End Function
    Public Function CreateNewOURdbOnNewServer(Optional ByVal myconstr As String = "", Optional ByVal myconprv As String = "", Optional ByVal dbcase As String = "lower") As String
        Dim ret As String = String.Empty
        Dim err As String = String.Empty
        Dim sqlq As String = String.Empty
        Dim pss As String = String.Empty
        Try
            If myconstr = "" Then
                myconprv = System.Configuration.ConfigurationManager.ConnectionStrings.Item("mySQLconnection").ProviderName.ToString
                myconstr = System.Configuration.ConfigurationManager.ConnectionStrings.Item("mySQLconnection").ToString
            End If
            Dim db As String = GetDataBase(myconstr, myconprv)
            Dim maindb As String = String.Empty
            If myconprv.StartsWith("InterSystems.Data.Cache") Then
                maindb = "%SYS"
            ElseIf myconprv.StartsWith("InterSystems.Data.IRIS") Then

            ElseIf myconprv = "MySql.Data.MySqlClient" Then
                maindb = "sys"
            ElseIf myconprv = "Oracle.ManagedDataAccess.Client" Then
                'maindb = "XE" 'XEPDB1
                db = GetUserIDFromConnectionString(myconstr)  'db.Substring(db.LastIndexOf("/") + 1)
                pss = GetPasswordFromConnectionString(myconstr)
                pss = Regex.Replace(pss, "[^a-zA-Z0-9]", "")  'CLS-compliant
            ElseIf myconprv = "System.Data.Odbc" Then
                maindb = db
            ElseIf myconprv = "Npgsql" Then
                maindb = "postgres" '"public"
            Else  'SQL Server
                'maindb = "master"
            End If
            myconstr = myconstr.Replace(db, maindb)
            ret = "Creating " & db & ": <br/> "
            If myconprv.StartsWith("InterSystems.Data.Cache") Then
                'create namespace
                If Not DatabaseExist(db, myconstr, myconprv) Then
                    sqlq = "CREATE DATABASE " & db
                    ret = ExequteSQLquery(sqlq, myconstr, myconprv)
                    If ret = "Query executed fine." Then
                        ret = ret & "<br/> " & " ourdb created"
                        ret = ret & "<br/> " & CreateInitialClass(myconstr, myconprv)
                    Else
                        ret = ret & "<br/> ERROR!!" & " ourdb not created: " & ret
                        'Return ret
                    End If
                Else
                    ret = ret & "<br/> " & " ourdb already exists"
                    ret = ret & "<br/> " & CreateInitialClass(myconstr, myconprv)
                End If
            ElseIf myconprv.StartsWith("InterSystems.Data.IRIS") Then
                myconstr = System.Configuration.ConfigurationManager.ConnectionStrings.Item("SystemSQLconnection").ToString
                myconprv = System.Configuration.ConfigurationManager.ConnectionStrings.Item("SystemSQLconnection").ProviderName.ToString

                If DatabaseExist(db, myconstr, myconprv) Then
                    ret = ret & "<br/> " & " ourdb already exists"
                    Return ret
                End If
                sqlq = "CREATE DATABASE " & db
                ret = ExequteSQLquery(sqlq, myconstr, myconprv)
                If ret = "Query executed fine." Then
                    ret = ret & "<br/> " & " ourdb created"
                    ret = ret & "<br/> " & CreateInitialClass(myconstr, myconprv)
                Else
                    ret = ret & "<br/> ERROR!!" & " ourdb not created: " & ret
                    'Return ret
                End If

            ElseIf myconprv = "System.Data.SqlClient" Then
                'myconstr = System.Configuration.ConfigurationManager.ConnectionStrings.Item("SystemSQLconnection").ToString
                'myconprv = System.Configuration.ConfigurationManager.ConnectionStrings.Item("SystemSQLconnection").ProviderName.ToString

                If Not DatabaseExist(db, myconstr, myconprv) Then
                    sqlq = "CREATE DATABASE " & db
                    ret = ExequteSQLquery(sqlq, myconstr, myconprv)
                    If ret = "Query executed fine." Then
                        ret = ret & "<br/> " & " ourdb created"
                    Else
                        ret = ret & "<br/> ERROR!!" & " ourdb not created: " & ret
                        'Return ret
                    End If
                Else
                    ret = ret & "<br/> " & " ourdb already exists"
                End If

            ElseIf myconprv = "MySql.Data.MySqlClient" Then
                If Not DatabaseExist(db, myconstr, myconprv) Then
                    sqlq = "CREATE DATABASE `" & db.ToLower & "`"
                    ret = ExequteSQLquery(sqlq, myconstr, myconprv)
                    If ret = "Query executed fine." Then
                        ret = ret & "<br/> " & " ourdb created"
                    Else
                        ret = ret & "<br/> ERROR!!" & " ourdb not created: " & ret
                        'Return ret
                    End If
                Else
                    ret = ret & "<br/> " & " ourdb already exists"
                End If

            ElseIf myconprv = "Oracle.ManagedDataAccess.Client" Then
                Try
                    myconstr = System.Configuration.ConfigurationManager.ConnectionStrings.Item("SystemSQLconnection").ToString
                    myconprv = System.Configuration.ConfigurationManager.ConnectionStrings.Item("SystemSQLconnection").ProviderName.ToString

                    Dim retr As String = String.Empty
                    If Not DatabaseExist(db, myconstr, myconprv) Then
                        sqlq = "alter session set ""_ORACLE_SCRIPT""=true"
                        retr = ExequteSQLquery(sqlq, myconstr, myconprv)
                        If retr = "Query executed fine." Then
                            ret = ret & "<br/> " & " alter session "
                        Else
                            ret = ret & "<br/> ERROR!!" & " alter session crashed: " & retr
                            'Exit Try
                        End If
                        sqlq = "CREATE USER " & db & " "
                        sqlq = sqlq & "IDENTIFIED BY " & pss & " DEFAULT TABLESPACE USERS TEMPORARY TABLESPACE TEMP"
                        retr = ExequteSQLquery(sqlq, myconstr, myconprv)
                        If retr = "Query executed fine." Then
                            ret = ret & "<br/> " & " ourdb created"
                        Else
                            ret = ret & "<br/> ERROR!!" & " ourdb not created: " & retr
                            'Exit Try
                        End If
                    End If
                    sqlq = "GRANT CONNECT,RESOURCE,CREATE SESSION,CREATE VIEW,CREATE MATERIALIZED VIEW, ALTER SESSION,CREATE DATABASE LINK,CREATE PROCEDURE,CREATE PUBLIC SYNONYM,CREATE ROLE,CREATE SEQUENCE,CREATE SYNONYM,CREATE TABLE,CREATE TRIGGER,CREATE TYPE,UNLIMITED TABLESPACE TO """ & db & """"
                    retr = ExequteSQLquery(sqlq, myconstr, myconprv)
                    If retr = "Query executed fine." Then
                        ret = ret & "<br/> " & " ourdb permissions granted"
                    Else
                        ret = ret & "<br/> ERROR!!" & " ourdb permissions not granted: " & retr
                        Exit Try
                    End If

                Catch ex As Exception
                    ret = ret & "<br/> ERROR!!" & " ourdb not created: " & ret
                End Try

            ElseIf myconprv = "System.Data.Odbc" Then
                Try
                    Dim er As String = String.Empty
                    Dim userODBCdriver As String = String.Empty
                    Dim userODBCdatabase As String = String.Empty
                    Dim userODBCdatasource As String = String.Empty
                    myconstr = myconstr.Replace("Password", "Pwd").Replace("User ID", "UID")
                    Dim bConnect As Boolean = DatabaseConnected(myconstr, myconprv, er, userODBCdriver, userODBCdatabase, userODBCdatasource)
                    If Not bConnect Then
                        er = "ERROR!! Database not connected. " & er
                        Return er
                    End If
                    Dim sqlf As String = String.Empty

                    'TODO !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!! Check why? For another button?

                    If userODBCdriver.ToUpper.StartsWith("PSQL") Then
                        sqlf = "CREATE SCHEMA """ & userODBCdatabase & """;"
                    End If
                    ret = ExequteSQLquery(sqlf, myconstr, myconprv)
                    If ret = "Query executed fine." Then
                        ret = ret & "<br/> " & " ourdb created"
                    Else
                        ret = ret & "<br/> ERROR!!" & " ourdb not created: " & ret
                        'Return ret
                    End If
                Catch ex As Exception
                    ret = ret & "<br/> ERROR!!" & " ourdb not created: " & ret
                End Try

            ElseIf myconprv = "Npgsql" Then
                If Not DatabaseExist(db, myconstr, myconprv) Then
                    Try
                        Dim sqlf As String = String.Empty
                        If dbcase = "doublequoted" Then
                            sqlf = "CREATE DATABASE """ & db & """;"
                        ElseIf dbcase = "lower" Then
                            sqlf = "CREATE DATABASE [" & db.ToLower & "];"
                        ElseIf dbcase = "upper" Then
                            sqlf = "CREATE DATABASE [" & db.ToUpper & "];"
                        Else
                            sqlf = "CREATE DATABASE [" & db.ToLower & "];"
                        End If
                        ret = ExequteSQLquery(sqlf, myconstr, myconprv)
                        'tables will be created in public schema in db database
                        'sqlf = "CREATE SCHEMA """ & db & """;"
                        'ret = ExequteSQLquery(sqlf, myconstr, myconprv)
                        If ret = "Query executed fine." Then
                            ret = ret & "<br/> " & " ourdb created"
                        Else
                            ret = ret & "<br/> ERROR!!" & " ourdb not created: " & ret
                            'Return ret
                        End If
                    Catch ex As Exception
                        ret = ret & "<br/> ERROR!!" & " ourdb not created: " & ret
                    End Try
                Else
                    ret = ret & "<br/> " & " ourdb already exists"
                End If

            Else  'SQL Server
                If Not DatabaseExist(db, myconstr, myconprv) Then
                    Try
                        sqlq = "CREATE DATABASE " & db
                        'sqlq = sqlq & " ON PRIMARY (NAME = " & db & "_Data, "
                        'sqlq = sqlq & "FILENAME = ""C:\Program Files\Microsoft SQL Server\MSSQL14.MSSQLSERVER\MSSQL\DATA\" & db & ".mdf"", "
                        'sqlq = sqlq & "SIZE = 20MB, MAXSIZE = 100MB, FILEGROWTH = 10%) "
                        'sqlq = sqlq & "LOG ON (NAME = " & db & "_Log, "
                        'sqlq = sqlq & "FILENAME = ""C:\Program Files\Microsoft SQL Server\MSSQL14.MSSQLSERVER\MSSQL\DATA\" & db & ".ldf"", "
                        'sqlq = sqlq & "SIZE = 1MB, "
                        'sqlq = sqlq & "MAXSIZE = 5MB, "
                        'sqlq = sqlq & "FILEGROWTH = 10%)"
                        ret = ExequteSQLquery(sqlq, myconstr, myconprv)
                        If ret = "Query executed fine." Then
                            ret = ret & "<br/> " & " ourdb " & db & " created"
                        Else
                            ret = ret & "<br/> ERROR!!" & " ourdb " & db & " not created: " & ret
                            'Return ret
                        End If
                    Catch ex As Exception
                        ret = ret & "<br/> ERROR!!" & " ourdb " & db & " not created: " & ret
                    End Try

                End If
            End If
        Catch ex As Exception
            ret = "<br/> ERROR!!" & ex.Message
        End Try
        Return ret
    End Function
    Public Function UpdateOURdbToCurrentVersion(Optional ByVal myconstr As String = "", Optional ByVal myconprv As String = "") As String
        Dim ret As String = String.Empty
        Dim err As String = String.Empty
        Dim sqlq As String = String.Empty
        Try
            If myconstr = "" Then
                myconprv = System.Configuration.ConfigurationManager.ConnectionStrings.Item("mySQLconnection").ProviderName.ToString
                myconstr = System.Configuration.ConfigurationManager.ConnectionStrings.Item("mySQLconnection").ToString
            End If
            Dim db As String = GetDataBase(myconstr)
            ret = "Updating " & db & ": <br/> "
            'call Create
            If Not DatabaseExist(db, myconstr, myconprv) Then
                Dim ourdbcase As String = ConfigurationManager.AppSettings("ourdbcase").ToString.Trim
                CreateNewOURdbOnNewServer(myconstr, myconprv, ourdbcase)
            End If


            ret = ret & "<br/>OURUnits: " & InstallOURUnits(myconstr, myconprv)
            ret = ret & "<br/>OURHelpDesk: " & InstallOURHelpDesk(myconstr, myconprv)
            ret = ret & "<br/>OURAccessLog: " & InstallOURAccessLog(myconstr, myconprv)
            ret = ret & "<br/>OURPermits: " & InstallOURPermits(myconstr, myconprv)
            ret = ret & "<br/>ourtasklistsetting: " & InstallOURTaskListSetting(myconstr, myconprv)
        Catch ex As Exception
            ret = "<br/> ERROR!!" & ex.Message
        End Try
        Return ret
    End Function


    Public Function InstallOURPermits(Optional ByVal myconstr As String = "", Optional ByVal myconprv As String = "") As String
        Dim ret As String = String.Empty
        Dim err As String = String.Empty
        Dim sqlq As String = String.Empty
        Try
            If myconstr = "" Then
                myconprv = System.Configuration.ConfigurationManager.ConnectionStrings.Item("mySQLconnection").ProviderName.ToString
                myconstr = System.Configuration.ConfigurationManager.ConnectionStrings.Item("mySQLconnection").ToString
            End If
            Dim db As String = GetDataBase(myconstr, myconprv)
            If TableExists("OURPermits", myconstr, myconprv, err) Then
                ret = "Table exists"
                'Update table structure
                'ret = ret & "<br/>Update OURPermits: " & UpdateOURPermits(myconstr, myconprv)

            Else
                'create table OURPermits
                If myconprv = "MySql.Data.MySqlClient" Then
                    db = db.ToLower
                    sqlq = "CREATE TABLE `" & db & "`.`ourpermits` ("
                    sqlq = sqlq & " `Access` char(10) DEFAULT NULL,"
                    sqlq = sqlq & " `PERMIT` char(10) DEFAULT NULL,"
                    sqlq = sqlq & " `NetId` char(120)  DEFAULT NULL,"
                    sqlq = sqlq & " `localpass` varchar(50) DEFAULT NULL,"
                    sqlq = sqlq & " `Name` varchar(50) DEFAULT NULL,"
                    sqlq = sqlq & " `Unit` varchar(50) DEFAULT NULL,"
                    sqlq = sqlq & " `Application` varchar(50) DEFAULT NULL,"
                    sqlq = sqlq & " `RoleApp` varchar(50) DEFAULT NULL,"
                    sqlq = sqlq & " `Group1` varchar(120) DEFAULT NULL,"
                    sqlq = sqlq & " `Group2` varchar(120) DEFAULT NULL,"
                    sqlq = sqlq & " `Group3` varchar(120) DEFAULT NULL,"
                    sqlq = sqlq & " `Comments` varchar(240) DEFAULT NULL,"
                    sqlq = sqlq & " `ConnStr` varchar(250) DEFAULT NULL,"
                    sqlq = sqlq & " `ConnPrv` varchar(250) DEFAULT NULL,"
                    sqlq = sqlq & " `StartDate` datetime DEFAULT NULL,"
                    sqlq = sqlq & " `EndDate` datetime DEFAULT NULL,"
                    sqlq = sqlq & " `paid` int(11) DEFAULT NULL,"
                    sqlq = sqlq & " `Email` varchar(250) DEFAULT NULL,"
                    sqlq = sqlq & " `Indx` int(11) Not NULL AUTO_INCREMENT,"
                    sqlq = sqlq & " PRIMARY KEY (`Indx`)"
                    sqlq = sqlq & ");"

                ElseIf myconprv = "Oracle.ManagedDataAccess.Client" Then
                    sqlq = "CREATE TABLE OURPERMITS ("
                    sqlq = sqlq & """Access"" VARCHAR2(10 CHAR) DEFAULT NULL,"
                    sqlq = sqlq & "PERMIT VARCHAR2(10 CHAR) DEFAULT NULL,"
                    sqlq = sqlq & "NetId VARCHAR2(120 CHAR ) DEFAULT NULL,"
                    sqlq = sqlq & "localpass VARCHAR2(50 CHAR) DEFAULT NULL,"
                    sqlq = sqlq & "Name VARCHAR2(50 CHAR) DEFAULT NULL,"
                    sqlq = sqlq & "Unit VARCHAR2(50 CHAR) DEFAULT NULL,"
                    sqlq = sqlq & "Application VARCHAR2(50 CHAR) DEFAULT NULL,"
                    sqlq = sqlq & "RoleApp VARCHAR2(50 CHAR) DEFAULT NULL,"
                    sqlq = sqlq & "Group1 VARCHAR2(120 CHAR) DEFAULT NULL,"
                    sqlq = sqlq & "Group2 VARCHAR2(120 CHAR) DEFAULT NULL,"
                    sqlq = sqlq & "Group3 VARCHAR2(120 CHAR) DEFAULT NULL,"
                    sqlq = sqlq & "Comments VARCHAR2(240 CHAR) DEFAULT NULL,"
                    sqlq = sqlq & "ConnStr VARCHAR2(250 CHAR) DEFAULT NULL,"
                    sqlq = sqlq & "ConnPrv VARCHAR2(250 CHAR) DEFAULT NULL,"
                    sqlq = sqlq & "StartDate DATE DEFAULT NULL,"
                    sqlq = sqlq & "EndDate DATE DEFAULT NULL,"
                    sqlq = sqlq & "Email VARCHAR2(250 CHAR) DEFAULT NULL,"
                    sqlq = sqlq & "Indx INTEGER GENERATED ALWAYS AS IDENTITY,"
                    sqlq = sqlq & "CONSTRAINT OURPERMITS_PK PRIMARY KEY (Indx)"
                    sqlq = sqlq & ")"

                ElseIf myconprv = "Npgsql" Then  'PostgreSQL  Npgsql
                    sqlq = "CREATE Table [OURPermits]("
                    sqlq = sqlq & "[Access] character varying(10) NULL,"
                    sqlq = sqlq & "[PERMIT] character varying(10) NULL,"
                    sqlq = sqlq & "[NetId] character varying(120) NULL,"
                    sqlq = sqlq & "[localpass] character varying(50) NULL,"
                    sqlq = sqlq & "[Name] character varying(50) NULL,"
                    sqlq = sqlq & "[Unit] character varying(50) NULL,"
                    sqlq = sqlq & "[Application] character varying(50) NULL,"
                    sqlq = sqlq & "[RoleApp] character varying(50) NULL,"
                    sqlq = sqlq & "[Group1] character varying(250) NULL,"
                    sqlq = sqlq & "[Group2] character varying(250) NULL,"
                    sqlq = sqlq & "[Group3] character varying(250) NULL,"
                    sqlq = sqlq & "[Comments] character varying(240) NULL,"
                    sqlq = sqlq & "[ConnStr] character varying(250) NULL,"
                    sqlq = sqlq & "[ConnPrv] character varying(250) NULL,"
                    sqlq = sqlq & "[StartDate] date NULL,"
                    sqlq = sqlq & "[EndDate] date NULL,"
                    sqlq = sqlq & "[paid] integer NULL,"
                    sqlq = sqlq & "[Email] character varying(250) NULL,"
                    sqlq = sqlq & "[Indx] integer NOT NULL GENERATED ALWAYS AS IDENTITY ( INCREMENT 1 START 1 MINVALUE 1 MAXVALUE 2147483647 CACHE 1 ),"
                    sqlq = sqlq & "CONSTRAINT [OURPermits_pkey] PRIMARY KEY ([Indx]))"

                Else 'Sql Server or Cache
                    sqlq = "CREATE TABLE OURPermits("
                    sqlq = sqlq & "[Access] [nchar](10) NULL,"
                    sqlq = sqlq & "[PERMIT] [nchar](10) NULL,"
                    sqlq = sqlq & "[NetId] [nchar](120) NULL,"
                    sqlq = sqlq & "[localpass] [nvarchar](50) NULL,"
                    sqlq = sqlq & "[Name] [nvarchar](50) NULL,"
                    sqlq = sqlq & "[Unit] [nvarchar](50) NULL,"
                    sqlq = sqlq & "[Application] [nvarchar](50) NULL,"
                    sqlq = sqlq & "[RoleApp] [nvarchar](50) NULL,"
                    sqlq = sqlq & "[Group1] [nvarchar](250) NULL,"
                    sqlq = sqlq & "[Group2] [nvarchar](250) NULL,"
                    sqlq = sqlq & "[Group3] [nvarchar](250) NULL,"
                    sqlq = sqlq & "[Comments] [nvarchar](240) NULL,"
                    sqlq = sqlq & "[ConnStr] [nvarchar](250) NULL,"
                    sqlq = sqlq & "[ConnPrv] [nvarchar](250) NULL,"
                    sqlq = sqlq & "[StartDate] [datetime] NULL,"
                    sqlq = sqlq & "[EndDate] [datetime] NULL,"
                    sqlq = sqlq & "[paid] [int] NULL,"
                    sqlq = sqlq & "[Email] [nvarchar](250) NULL,"
                    sqlq = sqlq & "[Indx] [int] IDENTITY(1,1) NOT NULL)"
                End If
                ret = ExequteSQLquery(sqlq, myconstr, myconprv)
            End If

            If ret = "Query executed fine." OrElse (TableExists("OURPermits", myconstr, myconprv, err)) Then
                ret = "OURPermits created"
                'insert first record for super user
                If myconprv = "MySql.Data.MySqlClient" Then
                    sqlq = "SELECT * FROM `" & db & "`.`ourpermits`"
                ElseIf myconprv = "Npgsql" Then  'PostgreSQL  Npgsql
                    'sqlq = "SELECT * FROM `" & db & "`.`OURPermits`"
                    sqlq = "SELECT * FROM `OURPermits`"
                Else
                    sqlq = "SELECT * FROM OURPermits"
                End If

            End If
        Catch ex As Exception
            ret = "<br/> " & ex.Message
        End Try
        Return ret
    End Function
    Public Function InstallOURAccessLog(Optional ByVal myconstr As String = "", Optional ByVal myconprv As String = "") As String
        Dim ret As String = String.Empty
        Dim err As String = String.Empty
        Dim sqlq As String = String.Empty
        Try
            If myconstr = "" Then
                myconprv = System.Configuration.ConfigurationManager.ConnectionStrings.Item("mySQLconnection").ProviderName.ToString
                myconstr = System.Configuration.ConfigurationManager.ConnectionStrings.Item("mySQLconnection").ToString
            End If
            Dim db As String = GetDataBase(myconstr, myconprv)
            If TableExists("OURAccessLog", myconstr, myconprv, err) Then
                ret = "Table exists"
                'Update table structure
                'ret = ret & "<br/>Update OURAccessLog: " & UpdateOURAccessLog(myconstr, myconprv)

            Else
                'create table OURAccessLog
                If myconprv.StartsWith("InterSystems.Data.") Then
                    sqlq = "CREATE Table [OURAccessLog]("
                    sqlq = sqlq & "[ID] [Int] IDENTITY(1, 1) Not NULL,"
                    sqlq = sqlq & "[EventDate] [smalldatetime] NULL,"
                    sqlq = sqlq & "[Logon] [nvarchar](50) NULL,"
                    sqlq = sqlq & "[""Count""] [Int] Not NULL,"
                    sqlq = sqlq & "[Action] [nvarchar](2500) NULL,"
                    sqlq = sqlq & "[Comments] [nvarchar](2500) NULL"
                    sqlq = sqlq & ""
                    sqlq = sqlq & ")"
                ElseIf myconprv = "MySql.Data.MySqlClient" Then
                    db = db.ToLower
                    sqlq = "CREATE TABLE `" & db & "`.`ouraccesslog` ("
                    sqlq = sqlq & " `ID` int(11) NOT NULL AUTO_INCREMENT,"
                    sqlq = sqlq & " `EventDate` datetime DEFAULT NULL,"
                    sqlq = sqlq & " `Logon` varchar(50) DEFAULT NULL,"
                    sqlq = sqlq & " `Count` int(11) NOT NULL DEFAULT '0',"
                    sqlq = sqlq & " `Action` varchar(2500) DEFAULT NULL,"
                    sqlq = sqlq & " `Comments` varchar(2500) DEFAULT NULL,"
                    sqlq = sqlq & " PRIMARY KEY(`ID`)"
                    sqlq = sqlq & ");"

                ElseIf myconprv = "Oracle.ManagedDataAccess.Client" Then
                    sqlq = "CREATE TABLE OURACCESSLOG ("
                    sqlq = sqlq & "ID INTEGER GENERATED ALWAYS AS IDENTITY,"
                    sqlq = sqlq & "EventDate TIMESTAMP(0),"
                    sqlq = sqlq & "Logon VARCHAR2(50 CHAR) DEFAULT NULL,"
                    sqlq = sqlq & """Count"" NUMBER(11,0) NOT NULL,"
                    sqlq = sqlq & "Action VARCHAR2(2500 CHAR) DEFAULT NULL,"
                    sqlq = sqlq & "Comments VARCHAR2(2500 CHAR) DEFAULT NULL,"
                    sqlq = sqlq & "CONSTRAINT OURACCESSLOG_PK PRIMARY KEY(ID)"
                    sqlq = sqlq & ")"

                ElseIf myconprv = "Npgsql" Then  'PostgreSQL  Npgsql
                    sqlq = "CREATE Table [OURAccessLog]("
                    sqlq = sqlq & "[EventDate] date NULL,"
                    sqlq = sqlq & "[Logon] character varying(50) NULL,"
                    sqlq = sqlq & "[Count] integer NULL,"
                    sqlq = sqlq & "[Action] character varying(2500) NULL,"
                    sqlq = sqlq & "[Comments] character varying(2500) NULL,"
                    sqlq = sqlq & "[ID] integer NOT NULL GENERATED ALWAYS AS IDENTITY ( INCREMENT 1 START 1 MINVALUE 1 MAXVALUE 2147483647 CACHE 1 ),"
                    sqlq = sqlq & "CONSTRAINT [OURAccessLog_pkey] PRIMARY KEY ([ID]))"

                Else 'Sql Server
                    sqlq = "CREATE Table [OURAccessLog]("
                    sqlq = sqlq & "[ID] [Int] IDENTITY(1, 1) Not NULL,"
                    sqlq = sqlq & "[EventDate] [smalldatetime] NULL,"
                    sqlq = sqlq & "[Logon] [nvarchar](50) NULL,"
                    sqlq = sqlq & "[Count] [Int] Not NULL,"
                    sqlq = sqlq & "[Action] [nvarchar](2500) NULL,"
                    sqlq = sqlq & "[Comments] [nvarchar](2500) NULL"
                    sqlq = sqlq & ""
                    sqlq = sqlq & ")"
                End If
                ret = ExequteSQLquery(sqlq, myconstr, myconprv)
                If ret = "Query executed fine." Then
                    ret = "OURAccessLog created"
                Else
                    ret = "OURAccessLog creation crashed: " & ret
                End If

            End If

        Catch ex As Exception
            ret = "<br/> " & ex.Message
        End Try
        Return ret
    End Function

    Public Function UninstallOURTable(ByVal tbl As String, Optional ByVal myconstr As String = "", Optional ByVal myconprv As String = "") As String
        Dim ret As String = String.Empty
        Dim err As String = String.Empty
        Dim sqlq As String = String.Empty
        If myconstr = "" Then
            myconprv = System.Configuration.ConfigurationManager.ConnectionStrings.Item("mySQLconnection").ProviderName.ToString
            myconstr = System.Configuration.ConfigurationManager.ConnectionStrings.Item("mySQLconnection").ToString
        End If
        Try
            If TableExists(tbl, myconstr, myconprv, err) Then
                If err = "" Then
                    'delete 
                    Dim db As String = GetDataBase(myconstr, myconprv)
                    If myconprv = "MySql.Data.MySqlClient" Then
                        db = db.ToLower
                        sqlq = "DROP TABLE `" & db & "`.`" & tbl.ToLower & "`;"
                    ElseIf myconprv = "Oracle.ManagedDataAccess.Client" Then
                        sqlq = "DROP Table " & tbl.ToUpper & " CASCADE CONSTAINTS"
                    ElseIf myconprv = "Npgsql" Then  'PostgreSQL  Npgsql
                        sqlq = "DROP TABLE [" & tbl & "];" 'CASCADE???
                    Else
                        sqlq = "DROP Table [" & tbl.ToUpper & "]"
                    End If
                    ret = ret & "<br/> " & " " & sqlq & "  " & ExequteSQLquery(sqlq, myconstr, myconprv)
                Else
                    ret = err
                End If
            End If
        Catch ex As Exception
            ret = ret & "<br/> " & sqlq & "  " & ex.Message
        End Try
        Return ret
    End Function
    Public Function InstallOURUnits(Optional ByVal myconstr As String = "", Optional ByVal myconprv As String = "") As String
        Dim ret As String = String.Empty
        Dim err As String = String.Empty
        Dim sqlq As String = String.Empty
        Try
            If myconstr = "" Then
                myconprv = System.Configuration.ConfigurationManager.ConnectionStrings.Item("mySQLconnection").ProviderName.ToString
                myconstr = System.Configuration.ConfigurationManager.ConnectionStrings.Item("mySQLconnection").ToString
            End If
            Dim db As String = GetDataBase(myconstr, myconprv)
            If TableExists("OURUnits", myconstr, myconprv, err) Then
                ret = "Table exists"
                'Update table structure
                'ret = ret & "<br/>Update OURUnits: " & UpdateOURUnits(myconstr, myconprv)

            ElseIf err = String.Empty Then
                If myconprv = "MySql.Data.MySqlClient" Then
                    db = db.ToLower
                    sqlq = "CREATE TABLE `" & db & "`.`ourunits` ("
                    sqlq = sqlq & " `Unit` varchar(120) DEFAULT 'OUR',"
                    sqlq = sqlq & " `DistrMode` varchar(50) DEFAULT NULL,"
                    sqlq = sqlq & " `UnitWeb` varchar(240) DEFAULT NULL,"
                    sqlq = sqlq & " `OURConnStr` varchar(240) DEFAULT NULL,"
                    sqlq = sqlq & " `OURConnPrv` varchar(120) DEFAULT NULL,"
                    sqlq = sqlq & " `UserConnStr` varchar(240) DEFAULT NULL,"
                    sqlq = sqlq & " `UserConnPrv` varchar(120) DEFAULT NULL,"
                    sqlq = sqlq & " `StartDate` varchar(120) DEFAULT NULL,"
                    sqlq = sqlq & " `EndDate` varchar(120) DEFAULT NULL,"
                    sqlq = sqlq & " `Comments` varchar(1200) DEFAULT NULL,"
                    sqlq = sqlq & " `Indx` int(11) NOT NULL AUTO_INCREMENT,"
                    sqlq = sqlq & " PRIMARY KEY (`Indx`)"
                    sqlq = sqlq & ");"
                ElseIf myconprv = "Oracle.ManagedDataAccess.Client" Then
                    sqlq = "CREATE TABLE OURUNITS ("
                    sqlq = sqlq & "Unit VARCHAR2(120 CHAR) DEFAULT NULL,"
                    sqlq = sqlq & "DistrMode VARCHAR2(50 CHAR) DEFAULT NULL,"
                    sqlq = sqlq & "UnitWeb VARCHAR2(240 CHAR) DEFAULT NULL,"
                    sqlq = sqlq & "OURConnStr VARCHAR2(240 CHAR) DEFAULT NULL,"
                    sqlq = sqlq & "OURConnPrv VARCHAR2(120 CHAR) DEFAULT NULL,"
                    sqlq = sqlq & "UserConnStr VARCHAR2(240 CHAR) DEFAULT NULL,"
                    sqlq = sqlq & "UserConnPrv VARCHAR2(120 CHAR) DEFAULT NULL,"
                    sqlq = sqlq & "StartDate VARCHAR2(120 CHAR) DEFAULT NULL,"
                    sqlq = sqlq & "EndDate VARCHAR2(120 CHAR) DEFAULT NULL,"
                    sqlq = sqlq & "Comments VARCHAR2(1200 CHAR) DEFAULT NULL,"
                    sqlq = sqlq & "Indx INTEGER GENERATED ALWAYS AS IDENTITY,"
                    sqlq = sqlq & "CONSTRAINT OURUNITS_PK PRIMARY KEY (Indx)"
                    sqlq = sqlq & ")"
                ElseIf myconprv = "Npgsql" Then  'PostgreSQL  Npgsql
                    sqlq = "CREATE Table [OURUnits]("
                    sqlq = sqlq & "[Unit] character varying(120) Not NULL DEFAULT 'OUR',"
                    sqlq = sqlq & "[DistrMode] character varying(50) NULL,"
                    sqlq = sqlq & "[UnitWeb] character varying(240) NULL,"
                    sqlq = sqlq & "[OURConnStr] character varying(240) NULL,"
                    sqlq = sqlq & "[OURConnPrv] character varying(120) NULL,"
                    sqlq = sqlq & "[UserConnStr] character varying(240) NULL,"
                    sqlq = sqlq & "[UserConnPrv] character varying(120) NULL,"
                    sqlq = sqlq & "[StartDate] character varying(120) NULL,"
                    sqlq = sqlq & "[EndDate] character varying(120) NULL,"
                    sqlq = sqlq & "[Comments] character varying(1200) NULL,"
                    sqlq = sqlq & "[Indx] integer NOT NULL GENERATED ALWAYS AS IDENTITY ( INCREMENT 1 START 1 MINVALUE 1 MAXVALUE 2147483647 CACHE 1 ),"
                    sqlq = sqlq & "CONSTRAINT [OURUnits_pkey] PRIMARY KEY ([Indx]))"
                Else
                    sqlq = "CREATE TABLE [OURUnits]("
                    sqlq = sqlq & "[Unit] [nvarchar](120) Not NULL DEFAULT 'OUR',"
                    sqlq = sqlq & "[DistrMode] [nvarchar](50) NULL,"
                    sqlq = sqlq & "[UnitWeb] [nvarchar](240) NULL,"
                    sqlq = sqlq & "[OURConnStr] [nvarchar](240) NULL,"
                    sqlq = sqlq & "[OURConnPrv] [nvarchar](120) NULL,"
                    sqlq = sqlq & "[UserConnStr] [nvarchar](240) NULL,"
                    sqlq = sqlq & "[UserConnPrv] [nvarchar](120) NULL,"
                    sqlq = sqlq & "[StartDate] [nvarchar](120) NULL,"
                    sqlq = sqlq & "[EndDate] [nvarchar](120) NULL,"
                    sqlq = sqlq & "[Comments] [nvarchar](1200) NULL,"
                    sqlq = sqlq & "[Indx] [Int] IDENTITY(1, 1) Not NULL"
                    sqlq = sqlq & ")"
                End If
                ret = ExequteSQLquery(sqlq, myconstr, myconprv)
                If ret = "Query executed fine." Then
                    ret = "OURUnits created"
                Else
                    ret = "OURUnits creation crashed: " & ret
                End If
            Else
                ret = err
            End If
            'Update column changes or additions
            If ret = "Table exists" OrElse ret = "OURUnits created" Then
                If Not ColumnExists("OurUnits", "Official", myconstr, myconprv) Then
                    If myconprv = "MySql.Data.MySqlClient" Then
                        sqlq = "ALTER TABLE `" & db & "`.`ourunits` ADD COLUMN `Official` VARCHAR(200) NULL DEFAULT NULL AFTER `Comments`;"
                    ElseIf myconprv = "Oracle.ManagedDataAccess.Client" Then
                        sqlq = "ALTER TABLE OURUNITS ADD Official VARCHAR2(200 CHAR) DEFAULT NULL"
                    ElseIf myconprv = "Npgsql" Then  'PostgreSQL  Npgsql
                        sqlq = "ALTER TABLE `OURUnits` ADD COLUMN `Official` character varying(200) NULL DEFAULT NULL;"
                    Else
                        sqlq = "ALTER TABLE [OURUnits] ADD [Official] NVARCHAR(200) NULL DEFAULT NULL"
                    End If
                    ret = ret & ExequteSQLquery(sqlq, myconstr, myconprv).Replace("Query executed fine.", "")
                End If
                If Not ColumnExists("OurUnits", "Address", myconstr, myconprv) Then
                    If myconprv = "MySql.Data.MySqlClient" Then
                        sqlq = "ALTER TABLE `" & db & "`.`ourunits` ADD COLUMN `Address` VARCHAR(1000) NULL DEFAULT NULL AFTER `Official`;"
                    ElseIf myconprv = "Oracle.ManagedDataAccess.Client" Then
                        sqlq = "ALTER TABLE OURUNITS ADD Address VARCHAR2(1000 CHAR) DEFAULT NULL"
                    ElseIf myconprv = "Npgsql" Then  'PostgreSQL  Npgsql
                        sqlq = "ALTER TABLE `OURUnits` ADD COLUMN `Address` character varying(1000) NULL DEFAULT NULL;"
                    Else
                        sqlq = "ALTER TABLE [OURUnits] ADD [Address] NVARCHAR(1000) NULL DEFAULT NULL"
                    End If
                    ret = ret & ExequteSQLquery(sqlq, myconstr, myconprv).Replace("Query executed fine.", "")
                End If
                If Not ColumnExists("OurUnits", "Phone", myconstr, myconprv) Then
                    If myconprv = "MySql.Data.MySqlClient" Then
                        sqlq = "ALTER TABLE `" & db & "`.`ourunits` ADD COLUMN `Phone` VARCHAR(100) NULL DEFAULT NULL AFTER `Address`;"
                    ElseIf myconprv = "Oracle.ManagedDataAccess.Client" Then
                        sqlq = "ALTER TABLE OURUNITS ADD Phone VARCHAR2(100 CHAR) DEFAULT NULL"
                    ElseIf myconprv = "Npgsql" Then  'PostgreSQL  Npgsql
                        sqlq = "ALTER TABLE `OURUnits` ADD COLUMN `Phone` character varying(100) NULL DEFAULT NULL;"
                    Else
                        sqlq = "ALTER TABLE [OURUnits] ADD [Phone] NVARCHAR(100) NULL DEFAULT NULL"
                    End If
                    ret = ret & ExequteSQLquery(sqlq, myconstr, myconprv).Replace("Query executed fine.", "")
                End If
                If Not ColumnExists("OurUnits", "Email", myconstr, myconprv) Then
                    If myconprv = "MySql.Data.MySqlClient" Then
                        sqlq = "ALTER TABLE `" & db & "`.`ourunits` ADD COLUMN `Email` VARCHAR(200) NULL DEFAULT NULL AFTER `Phone`;"
                    ElseIf myconprv = "Oracle.ManagedDataAccess.Client" Then
                        sqlq = "ALTER TABLE OURUNITS ADD Email VARCHAR2(200 CHAR) DEFAULT NULL"
                    ElseIf myconprv = "Npgsql" Then  'PostgreSQL  Npgsql
                        sqlq = "ALTER TABLE `OURUnits` ADD COLUMN `Email` character varying(200) NULL DEFAULT NULL;"
                    Else
                        sqlq = "ALTER TABLE [OURUnits] ADD [Email] NVARCHAR(200) NULL DEFAULT NULL"
                    End If
                    ret = ret & ExequteSQLquery(sqlq, myconstr, myconprv).Replace("Query executed fine.", "")
                End If
                If Not ColumnExists("OurUnits", "Agent", myconstr, myconprv) Then
                    If myconprv = "MySql.Data.MySqlClient" Then
                        sqlq = "ALTER TABLE `" & db & "`.`ourunits` ADD COLUMN `Agent` VARCHAR(200) NULL DEFAULT NULL AFTER `Email`;"
                    ElseIf myconprv = "Oracle.ManagedDataAccess.Client" Then
                        sqlq = "ALTER TABLE OURUNITS ADD Agent VARCHAR2(200 CHAR) DEFAULT NULL"
                    ElseIf myconprv = "Npgsql" Then  'PostgreSQL  Npgsql
                        sqlq = "ALTER TABLE `OURUnits` ADD COLUMN `Agent` character varying(200) NULL DEFAULT NULL;"
                    Else
                        sqlq = "ALTER TABLE [OURUnits] ADD [Agent] NVARCHAR(200) NULL DEFAULT NULL"
                    End If
                    ret = ret & ExequteSQLquery(sqlq, myconstr, myconprv).Replace("Query executed fine.", "")
                End If
                If Not ColumnExists("OurUnits", "Prop1", myconstr, myconprv) Then
                    If myconprv = "MySql.Data.MySqlClient" Then
                        sqlq = "ALTER TABLE `" & db & "`.`ourunits` ADD COLUMN `Prop1` VARCHAR(200) NULL DEFAULT NULL AFTER `Agent`;"
                    ElseIf myconprv = "Oracle.ManagedDataAccess.Client" Then
                        sqlq = "ALTER TABLE OURUNITS ADD Prop1 VARCHAR2(200 CHAR) DEFAULT NULL"
                    ElseIf myconprv = "Npgsql" Then  'PostgreSQL  Npgsql
                        sqlq = "ALTER TABLE `OURUnits` ADD COLUMN `Prop1` character varying(200) NULL DEFAULT NULL;"
                    Else
                        sqlq = "ALTER TABLE [OURUnits] ADD [Prop1] NVARCHAR(200) NULL DEFAULT NULL"
                    End If
                    ret = ret & ExequteSQLquery(sqlq, myconstr, myconprv).Replace("Query executed fine.", "")
                End If
                If Not ColumnExists("OurUnits", "Prop2", myconstr, myconprv) Then
                    If myconprv = "MySql.Data.MySqlClient" Then
                        sqlq = "ALTER TABLE `" & db & "`.`ourunits` ADD COLUMN `Prop2` VARCHAR(200) NULL DEFAULT NULL AFTER `Prop1`;"
                    ElseIf myconprv = "Oracle.ManagedDataAccess.Client" Then
                        sqlq = "ALTER TABLE OURUNITS ADD Prop2 VARCHAR2(200 CHAR) DEFAULT NULL"
                    ElseIf myconprv = "Npgsql" Then  'PostgreSQL  Npgsql
                        sqlq = "ALTER TABLE `OURUnits` ADD COLUMN `Prop2` character varying(200) NULL DEFAULT NULL;"
                    Else
                        sqlq = "ALTER TABLE [OURUnits] ADD [Prop2] NVARCHAR(200) NULL DEFAULT NULL"
                    End If
                    ret = ret & ExequteSQLquery(sqlq, myconstr, myconprv).Replace("Query executed fine.", "")
                End If
                If Not ColumnExists("OurUnits", "Prop3", myconstr, myconprv) Then
                    If myconprv = "MySql.Data.MySqlClient" Then
                        sqlq = "ALTER TABLE `" & db & "`.`ourunits` ADD COLUMN `Prop3` VARCHAR(200) NULL DEFAULT NULL AFTER `Prop2`;"
                    ElseIf myconprv = "Oracle.ManagedDataAccess.Client" Then
                        sqlq = "ALTER TABLE OURUNITS ADD Prop3 VARCHAR2(200 CHAR) DEFAULT NULL"
                    ElseIf myconprv = "Npgsql" Then  'PostgreSQL  Npgsql
                        sqlq = "ALTER TABLE `OURUnits` ADD COLUMN `Prop3` character varying(200) NULL DEFAULT NULL;"
                    Else
                        sqlq = "ALTER TABLE [OURUnits] ADD [Prop3] NVARCHAR(200) NULL DEFAULT NULL"
                    End If
                    ret = ret & ExequteSQLquery(sqlq, myconstr, myconprv).Replace("Query executed fine.", "")
                End If
            End If

            'for clons add initial unit OUR
            'insert first record for super user
            If myconprv = "MySql.Data.MySqlClient" Then
                sqlq = "SELECT * FROM `" & db & "`.`OURUnits`"
            ElseIf myconprv = "Npgsql" Then  'PostgreSQL  Npgsql
                'sqlq = "SELECT * FROM `" & db & "`.`OURUnits`"
                sqlq = "SELECT * FROM `OURUnits`"
            Else
                sqlq = "SELECT * FROM OURUnits"
            End If

        Catch ex As Exception
            ret = "<br/> " & ex.Message
        End Try
        Return ret
    End Function

    Public Function InstallOURHelpDesk(Optional ByVal myconstr As String = "", Optional ByVal myconprv As String = "") As String
        Dim ret As String = String.Empty
        Dim err As String = String.Empty
        Dim sqlq As String = String.Empty
        Try
            If myconstr = "" Then
                myconprv = System.Configuration.ConfigurationManager.ConnectionStrings.Item("mySQLconnection").ProviderName.ToString
                myconstr = System.Configuration.ConfigurationManager.ConnectionStrings.Item("mySQLconnection").ToString
            End If
            Dim db As String = GetDataBase(myconstr, myconprv)
            If TableExists("OURHelpDesk", myconstr, myconprv, err) Then
                ret = "Table exists"
                'Update table structure
                'ret = ret & "<br/>Update OURHelpDesk: " & UpdateOURHelpDesk(myconstr, myconprv)

            Else    'create table
                'HelpDesk
                If myconprv = "MySql.Data.MySqlClient" Then
                    db = db.ToLower
                    sqlq = "CREATE TABLE `" & db & "`.`ourhelpdesk` ("
                    sqlq = sqlq & "`Start` varchar(50) Default NULL,"
                    sqlq = sqlq & "`Name` varchar(50) Default NULL,"
                    sqlq = sqlq & "`Ticket` varchar(2500) Default NULL,"
                    sqlq = sqlq & "`Deadline` varchar(50) Default NULL,"
                    sqlq = sqlq & "`Status` varchar(50) Default NULL,"
                    sqlq = sqlq & "`comments` varchar(2500) Default NULL,"
                    sqlq = sqlq & "`ToWhom` varchar(250) Default NULL,"
                    sqlq = sqlq & "`Version` varchar(45) Default NULL,"
                    sqlq = sqlq & " `Prop1` varchar(200) DEFAULT NULL,"
                    sqlq = sqlq & " `Prop2` varchar(200) DEFAULT NULL,"
                    sqlq = sqlq & " `Prop3` varchar(200) DEFAULT NULL,"
                    sqlq = sqlq & " `ID` int(11) Not NULL AUTO_INCREMENT,"
                    sqlq = sqlq & " PRIMARY KEY (`ID`)"
                    sqlq = sqlq & ");"
                ElseIf myconprv = "Oracle.ManagedDataAccess.Client" Then
                    sqlq = "CREATE TABLE OURHELPDESK ("
                    sqlq = sqlq & """Start"" VARCHAR2(50 CHAR) DEFAULT NULL,"
                    sqlq = sqlq & "Name VARCHAR2(50 CHAR) DEFAULT NULL,"
                    sqlq = sqlq & "Ticket VARCHAR2(2500 CHAR) DEFAULT NULL,"
                    sqlq = sqlq & "Deadline VARCHAR2(50 CHAR) DEFAULT NULL,"
                    sqlq = sqlq & "Status VARCHAR2(50 CHAR) DEFAULT NULL,"
                    sqlq = sqlq & "comments VARCHAR2(2500 CHAR) DEFAULT NULL,"
                    sqlq = sqlq & "ToWhom VARCHAR2(250 CHAR) DEFAULT NULL,"
                    sqlq = sqlq & "Version VARCHAR2(45 CHAR) DEFAULT NULL,"
                    sqlq = sqlq & "Prop1 VARCHAR2(200 CHAR) DEFAULT NULL,"
                    sqlq = sqlq & "Prop2 VARCHAR2(200 CHAR) DEFAULT NULL,"
                    sqlq = sqlq & "Prop3 VARCHAR2(200 CHAR) DEFAULT NULL,"
                    sqlq = sqlq & "ID INTEGER GENERATED ALWAYS AS IDENTITY,"
                    sqlq = sqlq & "CONSTRAINT OURHELPDESK_PK PRIMARY KEY (ID)"
                    sqlq = sqlq & ")"

                ElseIf myconprv = "Npgsql" Then  'PostgreSQL  Npgsql
                    sqlq = "CREATE Table [OURHelpDesk]("
                    sqlq = sqlq & "[Start] character varying(50) NULL,"
                    sqlq = sqlq & "[Name] character varying(50) NULL,"
                    sqlq = sqlq & "[Ticket] character varying(10000) NULL,"
                    sqlq = sqlq & "[Deadline] character varying(50) NULL,"
                    sqlq = sqlq & "[Status] character varying(50) NULL,"
                    sqlq = sqlq & "[comments] character varying(10000) NULL,"
                    sqlq = sqlq & "[ToWhom] character varying(250) NULL,"
                    sqlq = sqlq & "[Version] character varying(45) NULL,"
                    sqlq = sqlq & "[Prop1] character varying(200) NULL,"
                    sqlq = sqlq & "[Prop2] character varying(200) NULL,"
                    sqlq = sqlq & "[Prop3] character varying(200) NULL,"
                    sqlq = sqlq & "[ID] integer NOT NULL GENERATED ALWAYS AS IDENTITY ( INCREMENT 1 START 1 MINVALUE 1 MAXVALUE 2147483647 CACHE 1 ),"
                    sqlq = sqlq & "CONSTRAINT [OURHelpDesk_pkey] PRIMARY KEY ([ID]))"

                Else 'SQL Server or Cache
                    sqlq = "CREATE TABLE [OURHelpDesk]("
                    sqlq = sqlq & "[Start] [nvarchar](50) NULL,"
                    sqlq = sqlq & "[Name] [nvarchar](50) NULL,"
                    sqlq = sqlq & "[Ticket] [nvarchar](MAX) NULL,"
                    sqlq = sqlq & "[Deadline] [nvarchar](50) NULL,"
                    sqlq = sqlq & "[Status] [nvarchar](50) NULL,"
                    sqlq = sqlq & "[comments] [nvarchar](MAX) NULL,"
                    sqlq = sqlq & "[ToWhom] [nvarchar](250) NULL,"
                    sqlq = sqlq & "[Version] [nvarchar](45) NULL,"
                    sqlq = sqlq & "[Prop1] [nvarchar](200) NULL,"
                    sqlq = sqlq & "[Prop2] [nvarchar](200) NULL,"
                    sqlq = sqlq & "[Prop3] [nvarchar](200) NULL,"
                    sqlq = sqlq & "[ID] [Int] IDENTITY(1, 1) Not NULL"
                    sqlq = sqlq & ")"
                End If
                ret = ExequteSQLquery(sqlq, myconstr, myconprv)
                If ret = "Query executed fine." Then
                    ret = "OURHelpDesk created"
                Else
                    ret = "OURHelpDesk creation crashed: " & ret
                End If
            End If

        Catch ex As Exception
            ret = "<br/> " & ex.Message
        End Try
        Return ret
    End Function
    Public Function InstallOURTaskListSetting(Optional ByVal myconstr As String = "", Optional ByVal myconprv As String = "") As String
        Dim ret As String = String.Empty
        Dim err As String = String.Empty
        Dim sqlq As String = String.Empty
        Try
            If myconstr = "" Then
                myconprv = System.Configuration.ConfigurationManager.ConnectionStrings.Item("mySQLconnection").ProviderName.ToString
                myconstr = System.Configuration.ConfigurationManager.ConnectionStrings.Item("mySQLconnection").ToString
            End If
            Dim db As String = GetDataBase(myconstr, myconprv)
            If TableExists("ourtasklistsetting", myconstr, myconprv, err) Then
                ret = "Table exists"
                'Update table structure
                'ret = ret & "<br/>Update OURTaskListSetting: " & UpdateOURTaskListSetting(myconstr, myconprv)

            Else
                If myconprv = "MySql.Data.MySqlClient" Then
                    sqlq = "CREATE TABLE `" & db.ToLower & "`.`ourtasklistsetting` ("
                    sqlq = sqlq & " `Unit` int(11) DEFAULT NULL,"
                    sqlq = sqlq & " `UnitName` VARCHAR( 250 ) NULL DEFAULT NULL,"
                    sqlq = sqlq & " `FldText` VARCHAR( 50 ) NOT NULL,"
                    sqlq = sqlq & " `FldOrder` INT( 3 ) NOT NULL DEFAULT 0,"
                    sqlq = sqlq & " `FldColor` VARCHAR( 20 ) NULL DEFAULT NULL ,"
                    sqlq = sqlq & " `User` VARCHAR( 250 ) NULL DEFAULT NULL ,"
                    sqlq = sqlq & " `Prop1` varchar(120) DEFAULT NULL,"
                    sqlq = sqlq & " `Prop2` varchar(120) DEFAULT NULL,"
                    sqlq = sqlq & " `Prop3` varchar(120) DEFAULT NULL,"
                    sqlq = sqlq & " `Indx` int(11) NOT NULL AUTO_INCREMENT,"
                    sqlq = sqlq & " PRIMARY KEY (`Indx`)"
                    sqlq = sqlq & ");"
                ElseIf myconprv = "Oracle.ManagedDataAccess.Client" Then
                    sqlq = "CREATE TABLE OURTASKLISTSETTING ("
                    sqlq = sqlq & "Unit NUMBER(11,0) DEFAULT NULL,"
                    sqlq = sqlq & "UnitName VARCHAR2(250 CHAR) DEFAULT NULL,"
                    sqlq = sqlq & "FldText VARCHAR2(50 CHAR) DEFAULT NULL,"
                    sqlq = sqlq & "FldOrder NUMBER(3,0) DEFAULT 0 NOT NULL,"
                    sqlq = sqlq & "FldColor VARCHAR2(20 CHAR) DEFAULT NULL,"
                    sqlq = sqlq & """User"" VARCHAR2(250 CHAR) DEFAULT NULL,"
                    sqlq = sqlq & "Prop1 VARCHAR2(120 CHAR) DEFAULT NULL,"
                    sqlq = sqlq & "Prop2 VARCHAR2(120 CHAR) DEFAULT NULL,"
                    sqlq = sqlq & "Prop3 VARCHAR2(120 CHAR) DEFAULT NULL,"
                    sqlq = sqlq & "Indx INTEGER GENERATED ALWAYS AS IDENTITY,"
                    sqlq = sqlq & "CONSTRAINT OURTASKLISTSETTING_PK PRIMARY KEY (Indx)"
                    sqlq = sqlq & ")"
                ElseIf myconprv = "Npgsql" Then  'PostgreSQL  Npgsql
                    sqlq = "CREATE Table [ourtasklistsetting]("
                    sqlq = sqlq & "[Unit] smallint NOT NULL DEFAULT 0,"
                    sqlq = sqlq & "[UnitName] character varying(250) NULL,"
                    sqlq = sqlq & "[FldText] character varying(50) NOT NULL,"
                    sqlq = sqlq & "[FldOrder] smallint NOT NULL DEFAULT 0,"
                    sqlq = sqlq & "[FldColor] character varying(20) NULL,"
                    sqlq = sqlq & """User"" character varying(250) NULL,"
                    sqlq = sqlq & "[Prop1] character varying(120) NULL,"
                    sqlq = sqlq & "[Prop2] character varying(120) NULL,"
                    sqlq = sqlq & "[Prop3] character varying(120) NULL,"
                    sqlq = sqlq & "[Indx] integer NOT NULL GENERATED ALWAYS AS IDENTITY ( INCREMENT 1 START 1 MINVALUE 1 MAXVALUE 2147483647 CACHE 1 ),"
                    sqlq = sqlq & "CONSTRAINT [ourtasklistsetting_pkey] PRIMARY KEY ([Indx]))"
                Else 'sql server or cache
                    sqlq = "CREATE TABLE [ourtasklistsetting]("
                    sqlq = sqlq & "[Unit] [Int] NULL,"
                    sqlq = sqlq & "[UnitName] [nvarchar](250) NULL,"
                    sqlq = sqlq & "[FldText] [nvarchar](50) NOT NULL,"
                    sqlq = sqlq & "[FldOrder] [Int] NOT NULL DEFAULT 0,"
                    sqlq = sqlq & "[FldColor] [nvarchar](20) NULL,"
                    If myconprv.StartsWith("InterSystems.Data.") Then
                        sqlq = sqlq & """User"" [nvarchar](250) NULL,"
                    Else
                        sqlq = sqlq & "[User] [nvarchar](250) NULL,"
                    End If
                    sqlq = sqlq & "[Prop1] [nvarchar](120) NULL,"
                    sqlq = sqlq & "[Prop2] [nvarchar](120) NULL,"
                    sqlq = sqlq & "[Prop3] [nvarchar](120) NULL,"
                    sqlq = sqlq & "[Indx] [Int] IDENTITY(1, 1) not NULL"
                    sqlq = sqlq & ")"
                End If
                ret = ExequteSQLquery(sqlq, myconstr, myconprv)
                If ret = "Query executed fine." Then
                    ret = "ourtasklistsetting created"
                Else
                    ret = "ourtasklistsetting creation crashed: " & ret
                End If

            End If

        Catch ex As Exception
            ret = "<br/> " & ex.Message
        End Try
        Return ret
    End Function


    Public Function UninstallOURTablesClasses(Optional ByVal myconstr As String = "", Optional ByVal myconprv As String = "") As String
        Dim ret As String = String.Empty
        Dim err As String = String.Empty
        Dim sqlq As String = String.Empty
        If myconstr = "" Then
            myconprv = System.Configuration.ConfigurationManager.ConnectionStrings.Item("mySQLconnection").ProviderName.ToString
            myconstr = System.Configuration.ConfigurationManager.ConnectionStrings.Item("mySQLconnection").ToString
        End If
        Dim db As String = GetDataBase(myconstr)
        If db.StartsWith("OURdata") Then
            ret = "Uninstalling OUR tables from " & db & " is not allowed!"
            Return ret
            Exit Function
        End If
        ret = "Uninstalling OUR tables from " & db & ": <br/> "

        ret = ret & "<br/>" & UninstallOURTable("OURUnits", myconstr, myconprv)
        ret = ret & "<br/>" & UninstallOURTable("OURHelpDesk", myconstr, myconprv)
        ret = ret & "<br/>" & UninstallOURTable("OURAccessLog", myconstr, myconprv)
        ret = ret & "<br/>" & UninstallOURTable("OURPermits", myconstr, myconprv)
        ret = ret & "<br/>" & UninstallOURTable("ourtasklistsetting", myconstr, myconprv)

        If myconprv.StartsWith("InterSystems.Data.") Then
            'uninstall OUR.Init class
            sqlq = "DROP TABLE OUR.INIT"
            ret = ret & "<br/> " & sqlq & "  " & ExequteSQLquery(sqlq)
        End If
        Return ret
    End Function


    Public Function IsCacheDatabase() As Boolean
        Dim ret As Boolean = False
        Try
            Dim myprovider As String
            myprovider = System.Configuration.ConfigurationManager.ConnectionStrings.Item("mySQLconnection").ProviderName.ToString
            If myprovider.StartsWith("InterSystems.Data.") Then
                ret = True
            End If
        Catch ex As Exception
        End Try
        Return ret
    End Function
    Public Function IsOracleDatabase() As Boolean
        Dim ret As Boolean = False
        Try
            Dim myprovider As String
            myprovider = System.Configuration.ConfigurationManager.ConnectionStrings.Item("mySQLconnection").ProviderName.ToString
            If myprovider.StartsWith("Oracle.") Then
                ret = True
            End If
        Catch ex As Exception
        End Try
        Return ret
    End Function
    Public Function FixReservedWords(ByVal fldname As String, Optional dbtype As String = "", Optional connstr As String = "") As String
        'This function is used for table names and field names as well.
        Dim ret As String = String.Empty
        Dim dbcase = String.Empty
        If dbtype = "" OrElse connstr.ToUpper = System.Configuration.ConfigurationManager.ConnectionStrings.Item("mySQLconnection").ToString.ToUpper Then
            dbtype = System.Configuration.ConfigurationManager.ConnectionStrings.Item("mySQLconnection").ProviderName.ToString
            dbcase = ConfigurationManager.AppSettings("ourdbcase").ToString
        ElseIf System.Configuration.ConfigurationManager.ConnectionStrings.Item("CSVconnection") IsNot Nothing AndAlso connstr.ToUpper = System.Configuration.ConfigurationManager.ConnectionStrings.Item("CSVconnection").ToString.ToUpper Then
            dbtype = System.Configuration.ConfigurationManager.ConnectionStrings.Item("CSVconnection").ProviderName.ToString
            dbcase = ConfigurationManager.AppSettings("csvdbcase").ToString
        Else
            dbcase = userdbcase
        End If
        If dbtype.StartsWith("InterSystems.Data.") AndAlso fldname.StartsWith("User.") Then
            fldname = fldname.Replace("User.", "")
        End If
        'If dbtype.StartsWith("InterSystems.Data.") AndAlso fldname.StartsWith("UserData.") Then
        '    fldname = fldname.Replace("UserData.", "")
        'End If
        If dbtype = "Npgsql" Then
            If dbcase = "doublequoted" Then
                ret = fldname
                If Not fldname.StartsWith("""") Then
                    ret = """" & ret
                End If
                If Not fldname.EndsWith("""") Then
                    ret = ret & """"
                End If
            ElseIf dbcase = "lower" Then
                ret = fldname.ToLower
            ElseIf dbcase = "upper" Then
                ret = fldname.ToUpper
            Else
                ret = fldname
            End If
            Return ret
        End If
        Try
            Dim sCheck As String = String.Empty
            'The SQL query will be created in OUR db provider syntax, and will be corrected for User if needed in Correct* and Convert* functions
            If dbtype = "" AndAlso IsCacheDatabase() Then    'checks OUR db provider
                'TODO check where we need this
                dbtype = "InterSystems.Data."
            ElseIf dbtype = "" Then
                dbtype = System.Configuration.ConfigurationManager.ConnectionStrings.Item("mySQLconnection").ProviderName.ToString
            End If

            'reserved words in Cache 
            Dim CacheReservedWords As String = "| ABSOLUTE | ADD | ALL | ALLOCATE | ALTER | AND | ANY | ARE | AS | "
            CacheReservedWords = CacheReservedWords & "ASC | ASSERTION | AT | AUTHORIZATION | AVG | BEGIN | BETWEEN | "
            CacheReservedWords = CacheReservedWords & "BIT | BIT_LENGTH | BOTH | BY | CASCADE | CASE | CAST | "
            CacheReservedWords = CacheReservedWords & "CHAR | CHARACTER | CHARACTER_LENGTH | CHAR_LENGTH | "
            CacheReservedWords = CacheReservedWords & "CHECK | CLOSE | COALESCE | COLLATE | COMMIT | CONNECT | "
            CacheReservedWords = CacheReservedWords & "CONNECTION | CONSTRAINT | CONSTRAINTS | CONTINUE | CONVERT | "
            CacheReservedWords = CacheReservedWords & "CORRESPONDING | COUNT | CREATE | CROSS | CURRENT | "
            CacheReservedWords = CacheReservedWords & "CURRENT_DATE | CURRENT_TIME | CURRENT_TIMESTAMP | "
            CacheReservedWords = CacheReservedWords & "CURRENT_USER | CURSOR | DATE | DEALLOCATE | DEC | DECIMAL | "
            CacheReservedWords = CacheReservedWords & "Declare | DEFAULT | DEFERRABLE | DEFERRED | DELETE | DESC | "
            CacheReservedWords = CacheReservedWords & "DESCRIBE | DESCRIPTOR | DIAGNOSTICS | DISCONNECT | DISTINCT | "
            CacheReservedWords = CacheReservedWords & "DOMAIN | Double | DROP | Else | End | ENDEXEC | ESCAPE | EXCEPT | "
            CacheReservedWords = CacheReservedWords & "EXCEPTION | EXEC | EXECUTE | EXISTS | EXTERNAL | EXTRACT | "
            CacheReservedWords = CacheReservedWords & "FALSE | FETCH | FIRST | FLOAT | FOR | FOREIGN | FOUND | FROM | FULL | "
            CacheReservedWords = CacheReservedWords & "Get | Global | GO | GoTo | GRANT | GROUP | HAVING | HOUR | "
            CacheReservedWords = CacheReservedWords & "IDENTITY | IMMEDIATE | IN | INDICATOR | INITIALLY | "
            CacheReservedWords = CacheReservedWords & "INNER | INPUT | INSENSITIVE | INSERT | INT | Integer | INTERSECT | "
            CacheReservedWords = CacheReservedWords & "INTERVAL | INTO | Is | ISOLATION | JOIN | LANGUAGE | LAST | "
            CacheReservedWords = CacheReservedWords & "LEADING | LEFT | LEVEL | Like | LOCAL | LOWER | MATCH | MAX | MIN | "
            CacheReservedWords = CacheReservedWords & "MINUTE | MODULE | NAMES | NATIONAL | NATURAL | NCHAR | "
            CacheReservedWords = CacheReservedWords & "Next | NO | Not | NULL | NULLIF | NUMERIC | OCTET_LENGTH | OF | ON | "
            CacheReservedWords = CacheReservedWords & "ONLY | OPEN | OPTION | Or | OUTER | OUTPUT | OVERLAPS | "
            CacheReservedWords = CacheReservedWords & "PAD | PARTIAL | PREPARE | PRESERVE | PRIMARY | PRIOR | PRIVILEGES | "
            CacheReservedWords = CacheReservedWords & "PROCEDURE | PUBLIC | READ | REAL | REFERENCES | RELATIVE | "
            CacheReservedWords = CacheReservedWords & "RESTRICT | REVOKE | RIGHT | ROLE | ROLLBACK | ROWS | "
            CacheReservedWords = CacheReservedWords & "SCHEMA | SCROLL | SECOND | SECTION | SELECT | SESSION_USER | "
            CacheReservedWords = CacheReservedWords & "Set | SMALLINT | SOME | SPACE | SQLERROR | SQLSTATE | STATISTICS | "
            CacheReservedWords = CacheReservedWords & "SUBSTRING | SUM | SYSDATE | SYSTEM_USER | TABLE | TEMPORARY | "
            CacheReservedWords = CacheReservedWords & "THEN | TIME | TIMEZONE_HOUR | TIMEZONE_MINUTE | TO | TOP | "
            CacheReservedWords = CacheReservedWords & "TRAILING | TRANSACTION | TRIM | TRUE | UNION | UNIQUE | "
            CacheReservedWords = CacheReservedWords & "UPDATE | UPPER | USER | USING | VALUES | VARCHAR | VARYING | WHEN | "
            CacheReservedWords = CacheReservedWords & "WHENEVER | WHERE | WITH | WORK | WRITE |"
            CacheReservedWords = CacheReservedWords.ToUpper

            'reserved words in SQL Server and MySql
            Dim SqlReservedWords As String = ",ADD,ALL,ALTER,AD,ANY,AS,ASC,AUTHORIZATION"
            SqlReservedWords = SqlReservedWords & ",BACKUP,BEGIN,BETWEEN,BREAK,BROWSE,BULK,BY"
            SqlReservedWords = SqlReservedWords & ",CASCADE,CASE,CHECK,CHECKPOINT,CLOSE,CLUSTERED,COALESCE"
            SqlReservedWords = SqlReservedWords & ",COLLATE,COLUMN,COMMIT,COMPUTE,CONSTRAINT,CONTAINS"
            SqlReservedWords = SqlReservedWords & ",CONTAINSTABLE,CONTINUE,CONVERT,CREATE,CROSS,CURRENT"
            SqlReservedWords = SqlReservedWords & ",CURRENT_DATE,CURRENT_TIME,CURRENT_TIMESTAMP,CURRENT_USER"
            SqlReservedWords = SqlReservedWords & ",CURSOR,DATABASE,DBCC,DEALLOCATE,DECLARE,DEFAULT,DELETE"
            SqlReservedWords = SqlReservedWords & ",DENY,DESC,DISK,DISTINCT,DISTRIBUTED,DOUBLE,DROP,DUMMY"
            SqlReservedWords = SqlReservedWords & ",DUMMY,DUMP,ELSE,END,ERRLVL,ESCAPE,EXCEPT,EXEC,EXECUTE"
            SqlReservedWords = SqlReservedWords & ",EXISTS,EXIT,FETCH,FILE,FILLFACTOR,FOR,FOREIGN,FREETEXT"
            SqlReservedWords = SqlReservedWords & ",FREETEXTTABLE,FROM,FULL,FUNCTION,GOTO,GRANT,GROUP,HAVING"
            SqlReservedWords = SqlReservedWords & ",HOLDLOCK,IDENTITY,IDENTITY_INSERT,IDENTITYCOL,IF,IN,INDEX"
            SqlReservedWords = SqlReservedWords & ",INNER,INSERT,INTERSECT,INTO,IS,JOIN,KEY,KILL,LEFT,LIKE"
            SqlReservedWords = SqlReservedWords & ",LINENO,LOAD,NATIONAL,NOCHECK,NONCLUSTERED,NOT,NULL,NULLIF"
            SqlReservedWords = SqlReservedWords & ",OF,OFF,OFFSETS,ON,OPEN,OPENDATASOURCE,OPENQUERY,OPENROWSET"
            SqlReservedWords = SqlReservedWords & ",OPENXML,OPTION,OR,ORDER,OUTER,OVER,PERCENT,PLAN,PRECISION"
            SqlReservedWords = SqlReservedWords & ",PRIMARY,PRINT,PROC,PROCEDURE,PUBLIC,RANGE,RAISEERROR,READ"
            SqlReservedWords = SqlReservedWords & ",READTEXT,RECONFIGURE,REFERENCES,REPLICATION,RESTORE"
            SqlReservedWords = SqlReservedWords & ",RESTRICT,RETURN,REVOKE,RIGHT,ROLLBACK,ROWCOUNT,ROWGUIDCOL"
            SqlReservedWords = SqlReservedWords & ",RULE,SAVE,SCHEMA,SELECT,SESSION_USER,SET,SETUSER,SHUTDOWN"
            SqlReservedWords = SqlReservedWords & ",SOME,STATISTICS,SYSTEM_USER,TABLE,TEXTSIZE,THEN,TO,TOP"
            SqlReservedWords = SqlReservedWords & ",TRAN,TRANSACTION,TRIGGER,TRUNCATE,TSEQUAL,UNION,UNIQUE"
            SqlReservedWords = SqlReservedWords & ",UPDATE,UPDATETEXT,USE,USER,VALUES,VARYING,VIEW,WAITFOR"
            SqlReservedWords = SqlReservedWords & ",WHEN,WHERE,WHILE,WITH,WRITETEXT,ABSOLUTE,ACTION,ADMIN"
            SqlReservedWords = SqlReservedWords & ",AFTER,AGGREGATE,ALIAS,ALLOCATE,ARE,ARRAY,ASSERTION,AT"
            SqlReservedWords = SqlReservedWords & ",BEFORE,BINARY,BIT,BLOB,BOOLEAN,BOTH,BREADTH,CALL,CASCADED"
            SqlReservedWords = SqlReservedWords & ",CAST,CATALOG,CHAR,CHARACTER,CLASS,CLOB,COLLATION,COMPLETION"
            SqlReservedWords = SqlReservedWords & ",CONNECT,CONNECTION,CONTRAINTS,CONSTRUCTOR,CORRESPONDING"
            SqlReservedWords = SqlReservedWords & ",CUBE,CURRENT_PATH,CURRENT_ROLE,CYCLE,DATA,DATE,DAY,DEC"
            SqlReservedWords = SqlReservedWords & ",DECIMAL,DEFERRABLE,DEFFERRED,DEPTH,DEREF,DESCRIBE,DESCRIPTOR"
            SqlReservedWords = SqlReservedWords & ",DESTROY,DESTRUCTOR,DETERMINISTIC,DICTIONARY,DIAGNOSTICS"
            SqlReservedWords = SqlReservedWords & ",DISCONNECT,DOMAIN,DYNAMIC,EACH,END-EXEC,EQUALS,EVERY"
            SqlReservedWords = SqlReservedWords & ",EXCEPTION,EXTERNAL,FALSE,FIRST,FLOAT,FOUND,FREE,GENERAL"
            SqlReservedWords = SqlReservedWords & ",GET,GLOBAL,GO,GROUPING,HOST,HOUR,IGNORE,IMMEDIATE,INDICATOR"
            SqlReservedWords = SqlReservedWords & ",INITIALIZE,INITIALLY,INOUT,INPUT,INT,INTEGER,INTERVAL"
            SqlReservedWords = SqlReservedWords & ",ISOLATION,ITERATE,LANGUAGE,LARGE,LAST,LATERAL,LEADING,LESS"
            SqlReservedWords = SqlReservedWords & ",LEVEL,LIMIT,LOCAL,LOCALTIME,LOCALTIMESTAMP,LOCATIOR,MAP"
            SqlReservedWords = SqlReservedWords & ",MATCH,MINUTE,MODIFIES,MODIFY,MODULE,MONTH,NAMES,NATURAL"
            SqlReservedWords = SqlReservedWords & ",NCLOB,NEW,NEXT,NO,NONE,NUMERIC,OBJECT,OLD,ONLY,OPERATION"
            SqlReservedWords = SqlReservedWords & ",ORDINALITY,OUT,OUTPUT,PAD,PARAMETER,PARAMETERS,PARTIAL"
            SqlReservedWords = SqlReservedWords & ",PATH,POSTFIX,PREFIX,PREORDER,PREPARE,PRESERVE,PRIOR"
            SqlReservedWords = SqlReservedWords & ",PRIVILEGES,READS,REAL,RECURSIVE,REF,REFERENCING,RELATIVE"
            SqlReservedWords = SqlReservedWords & ",RESULT,RETURNS,ROLE,ROLLUP,ROUTINE,ROW,ROWS,SAVEPOINT,SCROLL"
            SqlReservedWords = SqlReservedWords & ",SCROLL,SCOPE,SEARCH,SECOND,SECTION,SEQUENCE,SESSION,SETS"
            SqlReservedWords = SqlReservedWords & ",SIZE,SHOW,SMALLINT,SPACE,SPECIFIC,SPECIFICTYPE,SQL,SQLEXCEPTION"
            SqlReservedWords = SqlReservedWords & ",SQLSTATE,SQLWARNING,START,STATE,STATEMENT,STATIC,STRUCTURE"
            SqlReservedWords = SqlReservedWords & ",TEMPORARY,TERMINATE,THAN,TIME,TIMESTAMP,TIMEZONE_HOUR"
            SqlReservedWords = SqlReservedWords & ",TIMEZONE_MINUTE,TRAILING,TRANSLATION,TREAT,TRUE,UNDER"
            SqlReservedWords = SqlReservedWords & ",UNKNOWN,UNNEST,USAGE,USING,VALUE,VARCHAR,VARIABLE,WHENEVER"
            SqlReservedWords = SqlReservedWords & ",WITHOUT,WORK,WRITE,YEAR,ZONE,"

            Dim OracleReservedWords As String = ",AGGREGATE,AGGREGATES,ALL,ALLOW,ANALYZE,ANCESTOR,AND"
            OracleReservedWords &= ",ANY,AS,AS,AT,AVG,BETWEEN,BINARY_DOUBLE,BINARY_FLOAT,BLOB"
            OracleReservedWords &= ",BRANCH,BUILD,BY,BYTE,CASE,CAST,CHAR,CHILD,CLEAR,CLOB,COMMIT"
            OracleReservedWords &= ",COMPILE,CONSIDER,COUNT,DATATYPE,DATE,DATE_MEASURE,DAY"
            OracleReservedWords &= ",DECIMAL,DELETE,DESC,DESCENDANT,DIMENSION,DISALLOW,DIVISION"
            OracleReservedWords &= ",DML,ELSE,END,ESCAPE,EXECUTE,FIRST,FLOAT,FOR,FROM,HIERARCHIES"
            OracleReservedWords &= ",HIERARCHY,HOUR,IGNORE,IN,INFINITE,INSERT,INTEGER,INTERVAL"
            OracleReservedWords &= ",INTO,IS,LAST,LEAF_DESCENDANT,LEAVES,LEVEL,LEVELS,LIKE"
            OracleReservedWords &= ",LIKEC,LIKE2,LIKE4,LOAD,LOCAL,LOG_SPEC,LONG,MAINTAIN,MAX"
            OracleReservedWords &= ",MEASURE,MEASURES,MEMBER,MEMBERS,MERGE,MLSLABEL,MIN,MINUTE"
            OracleReservedWords &= ",MODEL,MONTH,NAN,NCHAR,NCLOB,NO,NONE,NOT,NULL,NULLS,NUMBER"
            OracleReservedWords &= ",NVARCHAR2,OF,OLAP,OLAP_DML_EXPRESSION,ON,ONLY,OPERATOR,OR"
            OracleReservedWords &= ",ORDER,OVER,OVERFLOW,PARALLEL,PARENT,PLSQL,PRUNE,RAW"
            OracleReservedWords &= ",RELATIVE,ROOT_ANCESTOR,ROWID,SCN,SECOND,SELF,SERIAL,SET"
            OracleReservedWords &= ",SOLVE,SOME,SORT,SPEC,SUM,SYNCH,TEXT_MEASURE,THEN,TIME"
            OracleReservedWords &= ",TIMESTAMP,TO,UNBRANCH,UPDATE,USING,VALIDATE,VALUES,VARCHAR2"
            OracleReservedWords &= ",WHEN,WHERE,WITHIN,WITH,YEAR,ZERO,ZONE,ACCESS,ADD,ALTER"
            OracleReservedWords &= ",ASC,AUDIT,CHECK,CLUSTER,COLUMN,COLUMN_VALUE,COMMENT,COMPRESS"
            OracleReservedWords &= ",CONNECT,CREATE,CURRENT,DEFAULT,DISTINCT,DROP,EXCLUSIVE"
            OracleReservedWords &= ",EXISTS,FILE,GRANT,GROUP,HAVING,IDENTIFIED,IMMEDIATE,INCREMENT"
            OracleReservedWords &= ",INDEX,INITIAL,INTERSECT,LOCK,MAXEXTENTS,MINUS,MODE,MODIFY"
            OracleReservedWords &= ",NESTED_TABLE_ID,NOAUDIT,NOCOMPRESS,NOWAIT,OFFLINE,ONLINE"
            OracleReservedWords &= ",OPTION,PCTFREE,PRIOR,PUBLIC,RANGE,RENAME,RESOURCE,REVOKE,ROW"
            OracleReservedWords &= ",ROWNUM,ROWS,SELECT,SESSION,SHARE,SIZE,SMALLINT,START"
            OracleReservedWords &= ",SUCCESSFUL,SYNONYM,SYSDATE,TABLE,TRIGGER,UID,UNION,UNIQUE"
            OracleReservedWords &= ",USER,VARCHAR,VIEW,WHENEVER,"

            'check if ret is reserved word in Cache, SQL Server, Oracle, or MySql 
            'And fix it putting [] or "" around depending of dbtype

            'Dim i As Integer
            'Dim fldparts() As String = fldname.Split(".")
            'For i = 0 To fldparts.Length - 1
            '    sCheck = "| " & fldparts(i).ToUpper & " |"
            '    If CacheReservedWords.Contains(sCheck) Then
            '        If dbtype.StartsWith("InterSystems.Data.") Then
            '            fldparts(i) = """" & fldparts(i) & """"
            '        Else
            '            fldparts(i) = "[" & fldparts(i) & "]"
            '        End If
            '    Else
            '        sCheck = "," & fldparts(i).ToUpper & ","
            '        If SqlReservedWords.Contains(sCheck) Then
            '            If dbtype.StartsWith("InterSystems.Data.") Then
            '                fldparts(i) = """" & fldparts(i) & """"
            '            Else
            '                fldparts(i) = "[" & fldparts(i) & "]"
            '            End If
            '        ElseIf OracleReservedWords.Contains(sCheck) Then
            '            If dbtype.StartsWith("InterSystems.Data.") Then
            '                fldparts(i) = """" & fldparts(i) & """"
            '            Else
            '                fldparts(i) = "[" & fldparts(i) & "]"
            '            End If
            '        End If

            '    End If
            'Next

            Dim i As Integer
            Dim fldparts() As String = fldname.Split(".")
            For i = 0 To fldparts.Length - 1
                sCheck = "," & fldparts(i).ToUpper & ","
                If dbtype.StartsWith("InterSystems.Data.") Then
                    sCheck = "| " & fldparts(i).ToUpper & " |"
                    If CacheReservedWords.Contains(sCheck) Then
                        fldparts(i) = """" & fldparts(i) & """"
                    End If
                ElseIf dbtype = "Oracle.ManagedDataAccess.Client" Then
                    If OracleReservedWords.Contains(sCheck) Then
                        fldparts(i) = """" & fldparts(i) & """"
                    End If
                ElseIf dbtype = "System.Data.Odbc" Then
                    'for ODBC we use Sql Server syntax
                    If SqlReservedWords.Contains(sCheck) OrElse CacheReservedWords.Contains(sCheck) OrElse OracleReservedWords.Contains(sCheck) Then
                        fldparts(i) = "[" & fldparts(i) & "]"
                    End If
                ElseIf dbtype = "System.Data.OleDb" Then
                    'for OleDb we use Sql Server syntax
                    If SqlReservedWords.Contains(sCheck) OrElse CacheReservedWords.Contains(sCheck) OrElse OracleReservedWords.Contains(sCheck) Then
                        fldparts(i) = "[" & fldparts(i) & "]"
                    End If
                Else
                    If SqlReservedWords.Contains(sCheck) Then
                        'TODO other providers ?
                        If dbtype.StartsWith("InterSystems.Data.") Then
                            fldparts(i) = """" & fldparts(i) & """"
                        ElseIf dbtype.Contains("MySql") Then
                            fldparts(i) = "`" & fldparts(i) & "`"
                        ElseIf dbtype.Contains("Npgsql") Then
                            fldparts(i) = """" & fldparts(i) & """"
                        Else 'not Cache
                            fldparts(i) = "[" & fldparts(i) & "]"
                        End If
                    End If
                End If
            Next
            ret = ""
            For i = 0 To fldparts.Length - 1
                If i > 0 Then ret = ret & "."
                ret = ret & fldparts(i)
            Next
        Catch ex As Exception
            ret = ex.Message
        End Try
        Return ret
    End Function
    Public Function RegisterTeamAdmin(ByVal accss As String, ByVal permit As String, ByVal logon As String, ByVal pass As String, ByVal name As String, ByVal org As String, ByVal appl As String, ByVal rol As String, ByVal grp1 As String, ByVal grp2 As String, ByVal grp3 As String, ByVal email As String, ByVal commts As String, ByVal userconnstr As String, ByVal userconnprv As String, ByVal StartDate As String, ByVal EndDate As String, Optional ByVal phone As String = "", Optional ByVal ourconnstr As String = "", Optional ByVal ourconnprv As String = "") As String
        Dim sqlq As String
        Dim ret As String = String.Empty
        Dim pagettl As String = ConfigurationManager.AppSettings("pagettl").ToString
        Dim adminemail = ConfigurationManager.AppSettings("supportemail").ToString
        Dim webhelpdesk As String = ConfigurationManager.AppSettings("webhelpdesk").ToString
        Try
            'sqlq = "INSERT INTO [OURPERMITS] ([Access],[PERMIT],[NetId],[localpass],[Name],[Unit],[Application],[RoleApp],[Group1],[Group2],[Group3],[Email],[Comments],[ConnStr],[ConnPrv],[StartDate],[EndDate]) "
            'sqlq = sqlq & " VALUES ('" & accss & "','" & permit & "','" & logon & "','" & pass & "','" & name & "','" & org & "','" & appl & "','" & rol & "','" & grp1 & "','" & grp2 & "','" & grp3 & "','" & email & "','" & commts & "','" & userconnstr & "','" & userconnprv & "','" & StartDate & "','" & EndDate & "')"

            sqlq = "INSERT INTO [OURPermits] (" & FixReservedWords("Access") & ",[PERMIT],[NetId],[localpass]," & FixReservedWords("Name") & ",[Unit]," & FixReservedWords("Application") & ",[RoleApp],[Group1],[Group2],[Group3],[Email],[Comments],[ConnStr],[ConnPrv],[StartDate],[EndDate]) "
            sqlq = sqlq & " VALUES ('" & accss & "','" & permit & "','" & logon & "','" & pass & "','" & name & "','" & org & "','" & appl & "','" & rol & "','" & grp1 & "','" & grp2 & "','" & grp3 & "','" & email & "','" & commts & "','" & userconnstr & "','" & userconnprv & "','"
            If userconnprv = "Oracle.ManagedDataAccess.Client" Then
                sqlq = sqlq & DateToStringFormat(CDate(StartDate), userconnprv, "dd-MMM-yy") & "','" & DateToStringFormat(CDate(EndDate), userconnprv, "dd-MMM-yy") & "')"
            Else
                sqlq = sqlq & StartDate & "','" & EndDate & "')"
            End If


            ret = ExequteSQLquery(sqlq, ourconnstr, ourconnprv)
            If ret = "Query executed fine." Then
                ret = "User has been registered."
                ret = WriteToAccessLog(logon, "Registered", 1)
                ret = SendHTMLEmail("", "Team Admin " & logon & " has been registered at " & pagettl & " for team " & org, "Team Admin " & logon & " has been registered in " & pagettl & " for team " & org & " at " & webhelpdesk, email, adminemail)
                WriteToAccessLog(logon, "Team Admin, phone " & phone & ", registered  in " & pagettl & " for team " & org & " at " & webhelpdesk & "  with result:  " & ret, 1)
                ret = "User has been registered."
            Else
                WriteToAccessLog(logon, "Register user crashed: " & ret, 1)
            End If
        Catch ex As Exception
            ret = ex.Message
            WriteToAccessLog(logon, "Registered email crashed: " & ret, 1)
        End Try
        Return ret
    End Function

    Public Function GetDatabaseFromConnectionString(ByVal cs As String) As String
        'extract namespace from the connection string cs
        cs = UCase(cs).Replace(" ", "")
        Try
            If cs = "" OrElse cs.IndexOf("DATABASE") < 0 Then
                Return ""
                Exit Function
            End If
            Dim nmspace As String = String.Empty
            'extract namespace from the connection string
            nmspace = cs.Substring(cs.IndexOf("DATABASE="))
            If nmspace = "" OrElse nmspace.IndexOf(";") < 0 Then
                Return ""
                Exit Function
            End If
            nmspace = nmspace.Substring(9, nmspace.IndexOf(";") - 9)
            Return nmspace
        Catch ex As Exception
            MsgBox(ex.Message)
            Return String.Empty
        End Try
    End Function
    Public Function GetPasswordFromConnectionString(ByVal cs As String) As String
        'extract Password from the connection string cs
        Try
            If cs = "" OrElse cs.IndexOf("Password") < 0 Then
                Return ""
                Exit Function
            End If
            cs = cs.Replace(" ", "") & ";"
            Dim nmspace As String = String.Empty
            'extract PASSWORD from the connection string
            'nmspace = cs.Substring(cs.IndexOf("Password="))
            'If nmspace = "" OrElse nmspace.IndexOf(";") < 0 Then
            '    Return ""
            '    Exit Function
            'End If
            'nmspace = nmspace.Substring(8, nmspace.IndexOf(";") - 11)
            nmspace = cs.Substring(cs.ToUpper.IndexOf("Password".ToUpper))
            If nmspace = "" Then
                Return ""
                Exit Function
            End If
            If nmspace.IndexOf(";") < 0 Then
                nmspace = nmspace.Substring(8).Replace("=", "").Replace(";", "")
            Else
                nmspace = nmspace.Substring(8, nmspace.IndexOf(";") - 8).Replace("=", "").Replace(";", "")
            End If
            Return nmspace
        Catch ex As Exception
            MsgBox(ex.Message)
            Return String.Empty
        End Try
    End Function
    Public Function GetUserIDFromConnectionString(ByVal cs As String) As String
        'extract user id from the connection string cs
        cs = UCase(cs).Replace(" ", "") & ";"
        Try
            If cs = "" OrElse cs.IndexOf("USERID=") < 0 Then
                Return ""
                Exit Function
            End If
            Dim nmspace As String = String.Empty
            'extract user from the connection string
            nmspace = cs.Substring(cs.IndexOf("USERID="))
            If nmspace = "" Then
                Return ""
                Exit Function
            End If
            nmspace = nmspace.Substring(7, nmspace.IndexOf(";") - 7).Replace("=", "")

            Return nmspace
        Catch ex As Exception
            MsgBox(ex.Message)
            Return String.Empty
        End Try
        ''extract namespace from the connection string cs
        'cs = UCase(cs)
        'Try
        '    If cs = "" OrElse cs.IndexOf("USER ID=") < 0 Then
        '        Return ""
        '        Exit Function
        '    End If
        '    Dim nmspace As String = String.Empty
        '    'extract user from the connection string
        '    nmspace = cs.Substring(cs.IndexOf("USER ID="))
        '    If nmspace = "" OrElse nmspace.IndexOf(";") < 0 Then
        '        Return ""
        '        Exit Function
        '    End If
        '    nmspace = nmspace.Substring(8, nmspace.IndexOf(";") - 8).Replace("=", "")
        '    Return nmspace
        'Catch ex As Exception
        '    MsgBox(ex.Message)
        '    Return String.Empty
        'End Try
    End Function
    Public Function GetServerNameFromConnectionString(ByVal cs As String) As String
        'extract namespace from the connection string cs
        cs = UCase(cs)
        Try
            If cs = "" OrElse cs.IndexOf("SERVER=") < 0 Then
                Return ""
                Exit Function
            End If
            Dim nmspace As String = String.Empty
            'extract server name from the connection string
            nmspace = cs.Substring(cs.IndexOf("SERVER="))
            If nmspace = "" OrElse nmspace.IndexOf(";") < 0 Then
                Return ""
                Exit Function
            End If
            nmspace = nmspace.Substring(6, nmspace.IndexOf(";") - 9)
            Return nmspace
        Catch ex As Exception
            MsgBox(ex.Message)
            Return String.Empty
        End Try
    End Function
    Public Function GetNamespaceFromConnectionString(ByVal cs As String) As String
        'extract namespace from the connection string cs
        cs = UCase(cs)
        Try
            If cs = "" OrElse cs.IndexOf("NAMESPACE =") < 0 Then
                Return ""
                Exit Function
            End If
            Dim nmspace As String = String.Empty
            'extract namespace from the connection string
            nmspace = cs.Substring(cs.IndexOf("NAMESPACE ="))
            If nmspace = "" OrElse nmspace.IndexOf(";") < 0 Then
                Return ""
                Exit Function
            End If
            nmspace = nmspace.Substring(12, nmspace.IndexOf(";") - 12)
            Return nmspace
        Catch ex As Exception
            MsgBox(ex.Message)
            Return String.Empty
        End Try
    End Function

    Public Function WriteToAccessLog(ByVal logon As String, ByVal actn As String, ByVal cnt As Integer) As String
        Dim ret As String = String.Empty
        actn = cleanText(actn)
        Try
            Dim EventDate As String = DateToString(DateTime.Now)
            If IsCacheDatabase() Then
                ret = ExequteSQLquery("INSERT INTO OURAccessLog (EventDate,LOGON,Action,[""Count""]) VALUES('" & EventDate & "','" & logon & "','" & actn & "'," & cnt.ToString & ")")
            ElseIf IsOracleDatabase() Then
                ret = ExequteSQLquery("INSERT INTO OURAccessLog (EventDate,LOGON,Action,[""Count""]) VALUES('" & DateToStringFormat(EventDate, "", "dd-MMM-yy") & "','" & logon & "','" & actn & "'," & cnt.ToString & ")")
            Else
                ret = ExequteSQLquery("INSERT INTO OURAccessLog (EventDate,LOGON,Action,[Count]) VALUES('" & EventDate & "','" & logon & "','" & actn & "'," & cnt.ToString & ")")
            End If
        Catch ex As Exception
            ret = ex.Message
        End Try
        Return ret
    End Function
    Public Function ServiceExpired(ByVal enddate As String, Optional ByRef err As String = "") As Boolean
        Dim ret As String = String.Empty
        Dim b As Boolean = False
        Try
            If enddate = "" Then
                b = False
                Return b
            End If
            'if enddate<curdate then b=true - expired
            If enddate < DateToString(Now()) Then
                b = True
            End If
        Catch ex As Exception
            ret = ex.Message
        End Try
        err = ret
        Return b
    End Function

    Public Function FixDatetimeInOURUnits() As String
        Dim ret As String = String.Empty
        'fix DateTime fields
        If Not IsCacheDatabase() Then
            ret = ExequteSQLquery("UPDATE OURUnits SET StartDate=NULL WHERE StartDate=0")
            ret = ret & ExequteSQLquery("UPDATE OURUnits SET EndDate=NULL WHERE EndDate=0")
        End If
        Return ret.Replace("Query executed fine.", "").Trim
    End Function

    Public Function FixDatetimeInOURPermits() As String
        Dim ret As String = String.Empty
        'fix DateTime fields
        If Not IsCacheDatabase() Then
            ret = ExequteSQLquery("UPDATE OURPERMITS SET StartDate=NULL WHERE StartDate=0")
            ret = ret & ExequteSQLquery("UPDATE OURPERMITS SET EndDate=NULL WHERE EndDate=0")
        End If
        Return ret.Replace("Query executed fine.", "").Trim
    End Function
    Public Function InsertRowIntoTable(ByVal tbl As String, ByVal mRow As DataRow, ByVal dt As DataTable, ByVal connstr1 As String, ByVal connprv1 As String, ByVal connstr2 As String, ByVal connprv2 As String) As String
        Dim ret As String = String.Empty
        Dim sqlstmt As String = ""
        Try
            If dt Is Nothing Then
                Dim dv As DataView = mRecords("SELECT * FROM " & tbl.ToLower, ret, connstr1, connprv1)
                If dv Is Nothing Then
                    Return ret
                Else
                    dt = dv.Table
                End If
            End If

            Dim sFields As String = String.Empty
            Dim sValues As String = String.Empty
            Dim fld As String = String.Empty

            For j As Integer = 0 To dt.Columns.Count - 1
                If dt.Columns(j).Caption.ToUpper <> "INDX" AndAlso dt.Columns(j).Caption.ToUpper <> "ID" Then
                    fld = FixReservedWords(dt.Columns(j).Caption, connprv2)
                    If sFields <> String.Empty Then
                        sFields &= "," & fld
                    Else
                        sFields = fld
                    End If
                    If TblFieldIsNumeric(tbl, dt.Columns(j).Caption, connstr2, connprv2) Then
                        If mRow(j).ToString.Trim = "" OrElse mRow(j).ToString.Trim = "NULL" Then
                            mRow(j) = "0"
                        End If
                        If sValues = String.Empty Then
                            sValues = mRow(j).ToString
                        Else
                            sValues &= "," & mRow(j).ToString
                        End If
                    ElseIf TblFieldIsDateTime(tbl, dt.Columns(j).Caption, connstr2, connprv2) Then
                        If mRow(j).ToString.Trim = "" OrElse mRow(j).ToString.Trim = "0" OrElse mRow(j).ToString.Trim = "NULL" Then
                            mRow(j) = "NULL"
                            If sValues = String.Empty Then
                                sValues = mRow(j)
                            Else
                                sValues &= "," & mRow(j)
                            End If
                        Else
                            If sValues = String.Empty Then
                                sValues = "'" & DateToString(CDate(mRow(j)), connprv2) & "'"
                            Else
                                sValues &= ",'" & DateToString(CDate(mRow(j)), connprv2) & "'"
                            End If
                        End If
                    Else
                        If mRow(j).ToString.Trim = "NULL" Then
                            If sValues = String.Empty Then
                                sValues = mRow(j).ToString
                            Else
                                sValues &= "," & mRow(j).ToString
                            End If
                        Else
                            If sValues = String.Empty Then
                                sValues = "'" & mRow(j).ToString & "'"
                            Else
                                sValues &= ",'" & mRow(j).ToString & "'"
                            End If
                        End If
                    End If
                End If
            Next
            sqlstmt = "INSERT INTO " & tbl & " (" & sFields & ") VALUES (" & sValues & ")"

            'sqlstmt = "INSERT INTO " & tbl & " SET "
            'For j = 0 To dt.Columns.Count - 1
            '    If dt.Columns(j).Caption <> "Indx" AndAlso dt.Columns(j).Caption <> "ID" Then
            '        sqlstmt = sqlstmt & FixReservedWords(dt.Columns(j).Caption, connprv2) & "="
            '        If TblFieldIsNumeric(tbl, dt.Columns(j).Caption, connstr2, connprv2) Then
            '            If mRow(j).ToString.Trim = "" OrElse mRow(j).ToString.Trim = "NULL" Then
            '                mRow(j) = "0"
            '            End If
            '            sqlstmt = sqlstmt & mRow(j)
            '        ElseIf TblFieldIsDateTime(tbl, dt.Columns(j).Caption, connstr2, connprv2) Then
            '            If mRow(j).ToString.Trim = "" OrElse mRow(j).ToString.Trim = "0" OrElse mRow(j).ToString.Trim = "NULL" Then
            '                mRow(j) = "NULL"
            '                sqlstmt = sqlstmt & mRow(j)
            '            Else
            '                sqlstmt = sqlstmt & "'" & mRow(j) & "'"
            '            End If

            '        Else
            '            sqlstmt = sqlstmt & "'" & mRow(j) & "'"
            '        End If
            '        If j < dt.Columns.Count - 1 Then
            '            sqlstmt = sqlstmt & ","
            '        End If
            '    End If
            'Next
            'If sqlstmt.EndsWith(",") Then
            '    sqlstmt = sqlstmt.Substring(0, sqlstmt.Length - 1)
            'End If
            ret = ExequteSQLquery(sqlstmt, connstr2, connprv2)
        Catch ex As Exception
            ret = ex.Message
        End Try
        Return ret
    End Function
    'Public Function InsertRowIntoTable(ByVal tbl As String, ByVal mRow As DataRow, ByVal dt As DataTable, ByVal connstr1 As String, ByVal connprv1 As String, ByVal connstr2 As String, ByVal connprv2 As String) As String
    '    Dim ret As String = String.Empty
    '    Dim sqlstmt As String = "INSERT INTO " & tbl & " SET "
    '    Try
    '        If dt Is Nothing Then
    '            Dim dv As DataView = mRecords("SELECT * FROM " & tbl.ToLower, ret, connstr1, connprv1)
    '            If dv Is Nothing Then
    '                Return ret
    '            Else
    '                dt = dv.Table
    '            End If
    '        End If

    '        sqlstmt = "INSERT INTO " & tbl & " SET "
    '        For j = 0 To dt.Columns.Count - 1
    '            If dt.Columns(j).Caption <> "Indx" AndAlso dt.Columns(j).Caption <> "ID" Then
    '                sqlstmt = sqlstmt & FixReservedWords(dt.Columns(j).Caption, connprv2) & "="
    '                If TblFieldIsNumeric(tbl, dt.Columns(j).Caption, connstr2, connprv2) Then
    '                    If mRow(j).ToString.Trim = "" OrElse mRow(j).ToString.Trim = "NULL" Then
    '                        mRow(j) = "0"
    '                    End If
    '                    sqlstmt = sqlstmt & mRow(j)
    '                ElseIf TblFieldIsDateTime(tbl, dt.Columns(j).Caption, connstr2, connprv2) Then
    '                    If mRow(j).ToString.Trim = "" OrElse mRow(j).ToString.Trim = "0" OrElse mRow(j).ToString.Trim = "NULL" Then
    '                        mRow(j) = "NULL"
    '                        sqlstmt = sqlstmt & mRow(j)
    '                    Else
    '                        sqlstmt = sqlstmt & "'" & mRow(j) & "'"
    '                    End If

    '                Else
    '                    sqlstmt = sqlstmt & "'" & mRow(j) & "'"
    '                End If
    '                If j < dt.Columns.Count - 1 Then
    '                    sqlstmt = sqlstmt & ","
    '                End If
    '            End If
    '        Next
    '        If sqlstmt.EndsWith(",") Then
    '            sqlstmt = sqlstmt.Substring(0, sqlstmt.Length - 1)
    '        End If
    '        ret = ExequteSQLquery(sqlstmt, connstr2, connprv2)
    '    Catch ex As Exception
    '        ret = ex.Message
    '    End Try
    '    Return ret
    'End Function
    'Public Function GetDbName(ByVal userconnstr As String, ByVal userconnprv As String) As String
    '    Dim repdb As String = userconnstr.Substring(0, userconnstr.IndexOf("User ID")).Trim
    '    Dim dbname As String = repdb.Substring(repdb.LastIndexOf("=")).Replace("=", "").Replace(";", "").Trim
    '    If userconnprv.StartsWith("InterSystems.Data.") Then
    '        dbname = GetNamespaceFromConnectionString(userconnstr)
    '    Else
    '        dbname = GetDatabaseFromConnectionString(userconnstr)
    '        If userconnprv = "Oracle.ManagedDataAccess.Client" Then
    '            dbname = Piece(dbname, "/", 2).Replace(";", "")
    '        End If
    '    End If
    '    Return dbname
    'End Function
    Public Function ExportToCSVtext(ByVal dt As DataTable, Optional ByVal dlm As String = "", Optional ByVal hdr As String = "", Optional ByVal ftr As String = "") As String
        Dim ExpText As String
        Dim txtline As String = String.Empty
        Dim ret As String = String.Empty
        Dim m, i, j As Integer
        If dt Is Nothing AndAlso hdr.Trim = "" Then
            Return ret
            Exit Function
        End If
        If dlm = "" Then dlm = Chr(9)
        Dim MyFile As StringBuilder = New StringBuilder

        Try
            txtline = ""
            MyFile.AppendLine(hdr)
            'MyFile.AppendLine(txtline)
            If Not dt Is Nothing Then
                'dt columns names
                m = dt.Columns.Count
                For i = 0 To m - 1
                    If dt.Columns(i).ColumnName.ToString.IndexOf(dlm) >= 0 Then
                        txtline = txtline & """" & dt.Columns(i).ColumnName & """"
                    Else
                        txtline = txtline & dt.Columns(i).ColumnName
                    End If

                    If i < m - 1 Then txtline = txtline & dlm
                Next
                MyFile.AppendLine(txtline)
                'dt rows
                txtline = ""
                For j = 0 To dt.Rows.Count - 1
                    txtline = ""
                    For i = 0 To m - 1
                        If dt.Rows(j).Item(i).ToString.IndexOf(dlm) >= 0 Then
                            txtline = txtline & """" & dt.Rows(j).Item(i).ToString & """"
                        Else
                            txtline = txtline & dt.Rows(j).Item(i).ToString
                        End If

                        If i < m - 1 Then txtline = txtline & dlm
                    Next
                    If Trim(txtline) <> "" Then
                        MyFile.AppendLine(txtline)
                    End If
                Next
            End If
            txtline = ""
            MyFile.AppendLine(txtline)
            MyFile.AppendLine(ftr)
            ExpText = MyFile.ToString

            MyFile = Nothing
            ret = ExpText.ToString
        Catch ex As Exception
            ret = "Error creating csv text" & ex.Message
            MyFile = Nothing
        End Try
        Return ret
    End Function
    Public Function GetReportMapFields(ByVal rep As String, ByVal MapName As String) As DataTable
        Dim ret As String = String.Empty
        Dim dt As DataTable = Nothing
        Try
            Dim selectSQL As String = "SELECT Val AS MapField,Prop1 AS Friendly,Prop2 AS ForMap,Prop3 AS MapName,[Order] as ord,Prop4 as descrtext,Prop5,Prop6,Prop7,Indx FROM OURReportFormat WHERE (ReportID='" & rep & "' AND Prop='MAPS' AND Prop3='" & MapName & "') ORDER BY ord,Indx"
            dt = mRecords(selectSQL, ret).Table
        Catch ex As Exception
            ret = ex.Message
        End Try
        Return dt
    End Function
End Module





