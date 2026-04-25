Imports System.Data
'Imports Oracle.ManagedDataAccess.Client

Public Module mDataProprietary

#Region "HasRecords - Oracle and InterSystems"
    Public Function HasRecords_IRIS(ByVal mySQL As String, ByVal myconstring As String, ByRef myRecords As DataTable) As String
        Dim ret As String = String.Empty
        'Try
        '    Dim dataIRISConnection As New InterSystems.Data.IRISClient.IRISConnection
        '    Dim dataCommand As New InterSystems.Data.IRISClient.IRISCommand
        '    Dim ds As New System.Data.DataSet
        '    dataIRISConnection = New InterSystems.Data.IRISClient.IRISConnection(myconstring)

        '    If dataIRISConnection.State = ConnectionState.Closed Then dataIRISConnection.Open()
        '    dataCommand.Connection = dataIRISConnection
        '    dataCommand.CommandType = CommandType.Text
        '    dataCommand.CommandTimeout = 300000
        '    dataCommand.CommandText = mySQL
        '    Dim dataAdapter As New InterSystems.Data.IRISClient.IRISDataAdapter(dataCommand)
        '    ' catch "Incorrect list format" error and ignore it. It is an internal Cache
        '    ' error and does not affect getting the data into the dataset.
        '    ' All other errors are returned.
        '    Try
        '        dataAdapter.Fill(ds)
        '    Catch exc As Exception
        '        If Not exc.Message.ToUpper.StartsWith("INCORRECT LIST FORMAT:") Then
        '            ret = "RETURN_FALSE"
        '            Return ret
        '        End If
        '    End Try
        '    'dataAdapter.Fill(ds)
        '    If ds.Tables.Count > 0 Then myRecords = ds.Tables(0)
        '    dataAdapter.Dispose()
        '    dataCommand.Dispose()
        '    dataIRISConnection.Close()
        'Catch ex As Exception
        '    ret = ex.Message
        'End Try
        Return ret
    End Function

    Public Function HasRecords_Cache(ByVal mySQL As String, ByRef myconstring As String, ByRef myRecords As DataTable) As String
        Dim ret As String = String.Empty
        'Try
        '    Dim dataCacheConnectionString As String = String.Empty
        '    Dim dataCacheConnection As New InterSystems.Data.CacheClient.CacheConnection
        '    Dim dataCommand As New InterSystems.Data.CacheClient.CacheCommand
        '    Dim ds As New System.Data.DataSet
        '    If myconstring = String.Empty Then
        '        myconstring = InterSystems.Data.CacheClient.CacheConnection.ConnectDlg()
        '    End If
        '    dataCacheConnection = New InterSystems.Data.CacheClient.CacheConnection(myconstring)
        '    If dataCacheConnection.State = ConnectionState.Closed Then dataCacheConnection.Open()
        '    dataCommand.Connection = dataCacheConnection
        '    dataCommand.CommandType = CommandType.Text
        '    dataCommand.CommandTimeout = 300000
        '    dataCommand.CommandText = mySQL
        '    Dim dataAdapter As New InterSystems.Data.CacheClient.CacheDataAdapter(dataCommand)
        '    ' catch "Incorrect list format" error and ignore it. It is an internal Cache
        '    ' error and does not affect getting the data into the dataset.
        '    Try
        '        dataAdapter.Fill(ds)
        '    Catch exc As Exception
        '        If Not exc.Message.ToUpper.StartsWith("INCORRECT LIST FORMAT:") Then
        '            ret = "RETURN_FALSE"
        '            Return ret
        '        End If
        '    End Try
        '    'dataAdapter.Fill(ds)
        '    If ds.Tables.Count > 0 Then myRecords = ds.Tables(0)
        '    dataCacheConnection.Close()
        'Catch ex As Exception
        '    ret = ex.Message
        'End Try
        Return ret
    End Function

    Public Function HasRecords_Oracle(ByVal mySQL As String, ByVal myconstring As String, ByRef myRecords As DataTable) As String
        Dim ret As String = String.Empty
        'Try
        '    Dim myConnection As Oracle.ManagedDataAccess.Client.OracleConnection
        '    Dim myCommand As New Oracle.ManagedDataAccess.Client.OracleCommand
        '    Dim myAdapter As Oracle.ManagedDataAccess.Client.OracleDataAdapter
        '    myConnection = New Oracle.ManagedDataAccess.Client.OracleConnection(myconstring)
        '    myCommand.Connection = myConnection
        '    myCommand.CommandType = CommandType.Text
        '    myCommand.CommandTimeout = 300000
        '    myCommand.CommandText = mySQL
        '    If myCommand.Connection.State = ConnectionState.Closed Then myCommand.Connection.Open()
        '    myAdapter = New Oracle.ManagedDataAccess.Client.OracleDataAdapter(myCommand)
        '    myAdapter.Fill(myRecords)
        '    myAdapter.Dispose()
        '    myCommand.Connection.Close()
        '    myCommand.Dispose()
        'Catch ex As Exception
        '    ret = ex.Message
        'End Try
        Return ret
    End Function
#End Region

#Region "CountOfRecords - Oracle and InterSystems"
    Public Function CountOfRecords_IRIS(ByVal mySQL As String, ByVal myconstring As String, ByRef myRecords As DataTable) As String
        Dim ret As String = String.Empty
        'Try
        '    Dim dataIRISConnection As New InterSystems.Data.IRISClient.IRISConnection
        '    Dim dataCommand As New InterSystems.Data.IRISClient.IRISCommand
        '    Dim ds As New System.Data.DataSet
        '    dataIRISConnection = New InterSystems.Data.IRISClient.IRISConnection(myconstring)

        '    If dataIRISConnection.State = ConnectionState.Closed Then dataIRISConnection.Open()
        '    dataCommand.Connection = dataIRISConnection
        '    dataCommand.CommandType = CommandType.Text
        '    dataCommand.CommandTimeout = 300000
        '    dataCommand.CommandText = mySQL
        '    Dim dataAdapter As New InterSystems.Data.IRISClient.IRISDataAdapter(dataCommand)
        '    Try
        '        dataAdapter.Fill(ds)
        '    Catch exc As Exception
        '        If Not exc.Message.ToUpper.StartsWith("INCORRECT LIST FORMAT:") Then
        '            ret = "ERROR!! " & exc.Message
        '            Return ret
        '        End If
        '    End Try
        '    If ds.Tables.Count > 0 Then myRecords = ds.Tables(0)
        '    dataAdapter.Dispose()
        '    dataCommand.Dispose()
        '    dataIRISConnection.Close()
        'Catch ex As Exception
        '    ret = ex.Message
        'End Try
        Return ret
    End Function

    Public Function CountOfRecords_Cache(ByVal mySQL As String, ByRef myconstring As String, ByRef myRecords As DataTable) As String
        Dim ret As String = String.Empty
        'Try
        '    Dim dataCacheConnectionString As String = String.Empty
        '    Dim dataCacheConnection As New InterSystems.Data.CacheClient.CacheConnection
        '    Dim dataCommand As New InterSystems.Data.CacheClient.CacheCommand
        '    Dim ds As New System.Data.DataSet
        '    If myconstring = String.Empty Then
        '        myconstring = InterSystems.Data.CacheClient.CacheConnection.ConnectDlg()
        '    End If
        '    dataCacheConnection = New InterSystems.Data.CacheClient.CacheConnection(myconstring)
        '    If dataCacheConnection.State = ConnectionState.Closed Then dataCacheConnection.Open()
        '    dataCommand.Connection = dataCacheConnection
        '    dataCommand.CommandType = CommandType.Text
        '    dataCommand.CommandTimeout = 300000
        '    dataCommand.CommandText = mySQL
        '    Dim dataAdapter As New InterSystems.Data.CacheClient.CacheDataAdapter(dataCommand)
        '    Try
        '        dataAdapter.Fill(ds)
        '    Catch exc As Exception
        '        If Not exc.Message.ToUpper.StartsWith("INCORRECT LIST FORMAT:") Then
        '            ret = "ERROR!! " & exc.Message
        '            Return ret
        '        End If
        '    End Try
        '    If ds.Tables.Count > 0 Then myRecords = ds.Tables(0)
        '    dataCacheConnection.Close()
        'Catch ex As Exception
        '    ret = ex.Message
        'End Try
        Return ret
    End Function

    Public Function CountOfRecords_Oracle(ByVal mySQL As String, ByVal myconstring As String, ByRef myRecords As DataTable) As String
        Dim ret As String = String.Empty
        'Try
        '    Dim myConnection As Oracle.ManagedDataAccess.Client.OracleConnection
        '    Dim myCommand As New Oracle.ManagedDataAccess.Client.OracleCommand
        '    Dim myAdapter As Oracle.ManagedDataAccess.Client.OracleDataAdapter
        '    myConnection = New Oracle.ManagedDataAccess.Client.OracleConnection(myconstring)
        '    myCommand.Connection = myConnection
        '    myCommand.CommandType = CommandType.Text
        '    myCommand.CommandTimeout = 300000
        '    myCommand.CommandText = mySQL
        '    If myCommand.Connection.State = ConnectionState.Closed Then myCommand.Connection.Open()
        '    myAdapter = New Oracle.ManagedDataAccess.Client.OracleDataAdapter(myCommand)
        '    myAdapter.Fill(myRecords)
        '    myAdapter.Dispose()
        '    myCommand.Connection.Close()
        '    myCommand.Dispose()
        'Catch ex As Exception
        '    ret = ex.Message
        'End Try
        Return ret
    End Function
#End Region

#Region "mRecords - Oracle and InterSystems"
    Public Function mRecords_IRIS(ByVal mySQL As String, ByRef er As String, ByVal myconstring As String, ByRef myRecords As DataTable) As String
        Dim ret As String = String.Empty
        'Try
        '    If mySQL.ToUpper.IndexOf(" DISTINCT ") > 0 Then
        '        ret = "USE_DISTINCT"
        '        Return ret
        '    Else
        '        Dim dataIRISConnection As New InterSystems.Data.IRISClient.IRISConnection
        '        Dim dataCommand As New InterSystems.Data.IRISClient.IRISCommand
        '        Try
        '            Dim ds As New System.Data.DataSet
        '            dataIRISConnection = New InterSystems.Data.IRISClient.IRISConnection(myconstring)
        '            If dataIRISConnection.State = ConnectionState.Closed Then dataIRISConnection.Open()
        '            dataCommand.Connection = dataIRISConnection
        '            dataCommand.CommandType = CommandType.Text
        '            dataCommand.CommandTimeout = 300000
        '            dataCommand.CommandText = mySQL
        '            Dim dataAdapter As New InterSystems.Data.IRISClient.IRISDataAdapter(dataCommand)
        '            Try
        '                dataAdapter.Fill(ds)
        '            Catch exc As Exception
        '                If Not exc.Message.ToUpper.StartsWith("INCORRECT LIST FORMAT:") Then
        '                    er = "ERROR!! " & exc.Message
        '                    ret = "RETURN_VIEW"
        '                    Return ret
        '                End If
        '            End Try
        '            If ds.Tables.Count > 0 Then myRecords = ds.Tables(0)
        '            dataAdapter.Dispose()
        '        Catch ex As Exception
        '            er = "ERROR!! " & ex.Message
        '        End Try
        '        dataCommand.Dispose()
        '        dataIRISConnection.Close()
        '    End If
        'Catch ex As Exception
        '    ret = ex.Message
        'End Try
        Return ret
    End Function

    Public Function mRecords_Cache(ByVal mySQL As String, ByRef er As String, ByRef myconstring As String, ByRef myRecords As DataTable) As String
        Dim ret As String = String.Empty
        'Try
        '    If mySQL.ToUpper.IndexOf(" DISTINCT ") > 0 Then
        '        ret = "USE_DISTINCT"
        '        Return ret
        '    Else
        '        Dim dataCacheConnectionString As String = String.Empty
        '        Dim dataCacheConnection As New InterSystems.Data.CacheClient.CacheConnection
        '        Dim dataCommand As New InterSystems.Data.CacheClient.CacheCommand
        '        Try
        '            Dim ds As New System.Data.DataSet
        '            If myconstring = String.Empty Then
        '                myconstring = InterSystems.Data.CacheClient.CacheConnection.ConnectDlg()
        '            End If
        '            dataCacheConnection = New InterSystems.Data.CacheClient.CacheConnection(myconstring)
        '            If dataCacheConnection.State = ConnectionState.Closed Then dataCacheConnection.Open()
        '            dataCommand.Connection = dataCacheConnection
        '            dataCommand.CommandType = CommandType.Text
        '            dataCommand.CommandTimeout = 300000
        '            dataCommand.CommandText = mySQL
        '            Dim dataAdapter As New InterSystems.Data.CacheClient.CacheDataAdapter(dataCommand)
        '            Try
        '                dataAdapter.Fill(ds)
        '            Catch exc As Exception
        '                If Not exc.Message.ToUpper.StartsWith("INCORRECT LIST FORMAT:") Then
        '                    er = "ERROR!! " & exc.Message
        '                    ret = "RETURN_VIEW"
        '                    Return ret
        '                End If
        '            End Try
        '            If ds.Tables.Count > 0 Then myRecords = ds.Tables(0)
        '            dataAdapter.Dispose()
        '        Catch ex As Exception
        '            er = "ERROR!! " & ex.Message
        '        End Try
        '        dataCommand.Dispose()
        '        dataCacheConnection.Close()
        '    End If
        'Catch ex As Exception
        '    ret = ex.Message
        'End Try
        Return ret
    End Function

    Public Function mRecords_Oracle(ByVal mySQL As String, ByRef er As String, ByVal myconstring As String, ByRef myRecords As DataTable) As String
        Dim ret As String = String.Empty
        'Try
        '    Dim myConnection As Oracle.ManagedDataAccess.Client.OracleConnection
        '    Dim myCommand As New Oracle.ManagedDataAccess.Client.OracleCommand
        '    Dim myAdapter As Oracle.ManagedDataAccess.Client.OracleDataAdapter
        '    myConnection = New Oracle.ManagedDataAccess.Client.OracleConnection(myconstring)
        '    myCommand.Connection = myConnection
        '    myCommand.CommandType = CommandType.Text
        '    myCommand.CommandTimeout = 300000
        '    myCommand.CommandText = mySQL
        '    If myCommand.Connection.State = ConnectionState.Closed Then myCommand.Connection.Open()
        '    myAdapter = New Oracle.ManagedDataAccess.Client.OracleDataAdapter(myCommand)
        '    Try
        '        myAdapter.Fill(myRecords)
        '    Catch ex As Exception
        '        'er = "ERROR!! " & ex.Message
        '    End Try
        '    myAdapter.Dispose()
        '    myCommand.Connection.Close()
        '    myCommand.Dispose()
        'Catch ex As Exception
        '    ret = ex.Message
        'End Try
        Return ret
    End Function
#End Region

#Region "RunSP - Oracle and InterSystems"
    Public Function RunSP_IRIS(ByVal mySP As String, ByVal Nparameters As Integer, ByVal ParamName As Array, ByVal ParamType As Array, ByVal ParamValue As Array, ByVal myconstring As String) As String
        Dim ret As String = String.Empty
        'Try
        '    Dim dataCacheConnectionString As String = String.Empty
        '    Dim dataCacheConnection As New InterSystems.Data.IRISClient.IRISConnection
        '    Dim dataCommand As New InterSystems.Data.IRISClient.IRISCommand
        '    dataCacheConnection = New InterSystems.Data.IRISClient.IRISConnection(myconstring)
        '    If dataCacheConnection.State = ConnectionState.Closed Then dataCacheConnection.Open()
        '    dataCommand.Connection = dataCacheConnection
        '    dataCommand.CommandType = CommandType.StoredProcedure
        '    dataCommand.CommandText = mySP
        '    dataCommand.CommandTimeout = 300000
        '    Dim i As Integer
        '    If Nparameters > 0 Then
        '        For i = 0 To Nparameters - 1
        '            If Not ParamValue(i) Is Nothing Then
        '                dataCommand.Parameters.Add(ParamName(i).ToString, ParamValue(i))  'working !!!
        '            End If
        '        Next
        '    End If
        '    dataCommand.ExecuteNonQuery()
        '    dataCommand.Dispose()
        '    dataCacheConnection.Close()
        'Catch ex As Exception
        '    ret = ex.Message
        'End Try
        Return ret
    End Function

    Public Function RunSP_Cache(ByVal mySP As String, ByVal Nparameters As Integer, ByVal ParamName As Array, ByVal ParamType As Array, ByVal ParamValue As Array, ByRef myconstring As String) As String
        Dim ret As String = String.Empty
        'Try
        '    Dim dataCacheConnectionString As String = String.Empty
        '    Dim dataCacheConnection As New InterSystems.Data.CacheClient.CacheConnection
        '    Dim dataCommand As New InterSystems.Data.CacheClient.CacheCommand
        '    Dim ds As New System.Data.DataSet
        '    If myconstring = String.Empty Then
        '        myconstring = InterSystems.Data.CacheClient.CacheConnection.ConnectDlg()
        '    End If
        '    dataCacheConnection = New InterSystems.Data.CacheClient.CacheConnection(myconstring)
        '    If dataCacheConnection.State = ConnectionState.Closed Then dataCacheConnection.Open()
        '    dataCommand.Connection = dataCacheConnection
        '    dataCommand.CommandType = CommandType.StoredProcedure
        '    dataCommand.CommandText = mySP
        '    dataCommand.CommandTimeout = 300000
        '    If Nparameters > 0 Then
        '        For i = 0 To Nparameters - 1
        '            If Not ParamValue(i) Is Nothing Then
        '                dataCommand.Parameters.Add(ParamName(i).ToString, ParamValue(i))  'working !!!
        '            End If
        '        Next
        '    End If
        '    dataCommand.ExecuteNonQuery()
        '    dataCommand.Dispose()
        '    dataCacheConnection.Close()
        'Catch ex As Exception
        '    ret = ex.Message
        'End Try
        Return ret
    End Function

    Public Function RunSP_Oracle(ByVal mySP As String, ByVal Nparameters As Integer, ByVal ParamName As Array, ByVal ParamType As Array, ByVal ParamValue As Array, ByVal myconstring As String) As String
        Dim ret As String = String.Empty
        'Try
        '    Dim myConnection As Oracle.ManagedDataAccess.Client.OracleConnection
        '    Dim myCommand As New Oracle.ManagedDataAccess.Client.OracleCommand
        '    myConnection = New Oracle.ManagedDataAccess.Client.OracleConnection(myconstring)
        '    myCommand.Connection = myConnection
        '    myCommand.CommandType = CommandType.StoredProcedure
        '    myCommand.CommandTimeout = 300000
        '    myCommand.CommandText = mySP
        '    If myCommand.Connection.State = ConnectionState.Closed Then myCommand.Connection.Open()
        '    Dim param(Nparameters) As MySql.Data.MySqlClient.MySqlParameter
        '    For i = 0 To Nparameters - 1
        '        If ParamType(i) = "nvarchar" Then
        '            param(i) = New MySql.Data.MySqlClient.MySqlParameter("@" + ParamName(i), MySql.Data.MySqlClient.MySqlDbType.VarChar, 255, ParameterDirection.Input)
        '        ElseIf ParamType(i) = "datetime" Then
        '            param(i) = New MySql.Data.MySqlClient.MySqlParameter("@" + ParamName(i), MySql.Data.MySqlClient.MySqlDbType.DateTime, 255, ParameterDirection.Input)
        '        Else
        '            param(i) = New MySql.Data.MySqlClient.MySqlParameter("@" + ParamName(i), MySql.Data.MySqlClient.MySqlDbType.Int16, 255, ParameterDirection.Input)
        '        End If
        '        param(i).Value = ParamValue(i)
        '        myCommand.Parameters.Add(param(i))
        '    Next
        '    If myCommand.Connection.State = ConnectionState.Closed Then myCommand.Connection.Open()
        '    myCommand.ExecuteNonQuery()
        '    myCommand.Connection.Close()
        '    myCommand.Dispose()
        'Catch ex As Exception
        '    ret = ex.Message
        'End Try
        Return ret
    End Function
#End Region

#Region "DatabaseConnected - Oracle and InterSystems"
    Public Function DatabaseConnected_IRIS(ByVal myconstring As String, ByRef er As String) As Boolean
        Dim r As Boolean = False
        'Try
        '    Dim dataCacheConnection As New InterSystems.Data.IRISClient.IRISConnection
        '    Dim dataCommand As New InterSystems.Data.IRISClient.IRISCommand
        '    Dim ds As New System.Data.DataSet
        '    dataCacheConnection = New InterSystems.Data.IRISClient.IRISConnection(myconstring)
        '    Try
        '        If dataCacheConnection.State = ConnectionState.Closed Then dataCacheConnection.Open()
        '        If dataCacheConnection.State = ConnectionState.Open Then r = True
        '        dataCacheConnection.Close()
        '    Catch ex As Exception
        '        er = ex.Message
        '        'dataCommand.Connection.Close()
        '        dataCommand.Dispose()
        '        dataCacheConnection.Dispose()
        '        r = False
        '    End Try
        'Catch ex As Exception
        '    er = ex.Message
        'End Try
        Return r
    End Function

    Public Function DatabaseConnected_Cache(ByRef myconstring As String, ByRef er As String) As Boolean
        Dim r As Boolean = False
        'Try
        '    Dim dataCacheConnection As New InterSystems.Data.CacheClient.CacheConnection
        '    Dim dataCommand As New InterSystems.Data.CacheClient.CacheCommand
        '    Dim ds As New System.Data.DataSet
        '    If myconstring = String.Empty Then
        '        myconstring = InterSystems.Data.CacheClient.CacheConnection.ConnectDlg()
        '    End If
        '    dataCacheConnection = New InterSystems.Data.CacheClient.CacheConnection(myconstring)
        '    Try
        '        If dataCacheConnection.State = ConnectionState.Closed Then dataCacheConnection.Open()
        '        If dataCacheConnection.State = ConnectionState.Open Then r = True
        '        dataCacheConnection.Close()
        '    Catch ex As Exception
        '        er = ex.Message
        '        'dataCommand.Connection.Close()
        '        dataCommand.Dispose()
        '        dataCacheConnection.Dispose()
        '        r = False
        '    End Try
        'Catch ex As Exception
        '    er = ex.Message
        'End Try
        Return r
    End Function

    Public Function DatabaseConnected_Oracle(ByVal myconstring As String, ByRef er As String) As Boolean
        Dim r As Boolean = False
        'Try
        '    Dim myConnection As Oracle.ManagedDataAccess.Client.OracleConnection
        '    myConnection = New Oracle.ManagedDataAccess.Client.OracleConnection(myconstring)
        '    Try
        '        If myConnection.State = ConnectionState.Closed Then myConnection.Open()
        '        If myConnection.State = ConnectionState.Open Then r = True
        '        myConnection.Close()
        '        myConnection.Dispose()
        '    Catch ex As Exception
        '        myConnection.Close()
        '        myConnection.Dispose()
        '        er = ex.Message
        '        r = False
        '    End Try
        'Catch ex As Exception
        '    er = ex.Message
        'End Try
        Return r
    End Function
#End Region

#Region "ExequteSQLquery - Oracle and InterSystems"
    Public Function ExequteSQLquery_IRIS(ByVal SQLq As String, ByVal myconstring As String) As String
        Dim r As String = "Query executed fine."
        'Try
        '    Dim dataIRISConnection As New InterSystems.Data.IRISClient.IRISConnection
        '    Dim dataCommand As New InterSystems.Data.IRISClient.IRISCommand
        '    Dim ds As New System.Data.DataSet
        '    dataIRISConnection = New InterSystems.Data.IRISClient.IRISConnection(myconstring)
        '    Try
        '        If dataIRISConnection.State = ConnectionState.Closed Then dataIRISConnection.Open()
        '        dataCommand.Connection = dataIRISConnection
        '        dataCommand.CommandType = CommandType.Text
        '        dataCommand.CommandTimeout = 300000
        '        dataCommand.CommandText = SQLq
        '        dataCommand.ExecuteNonQuery()
        '        dataCommand.Dispose()
        '        dataIRISConnection.Close()
        '    Catch ex As Exception
        '        'dataCommand.Connection.Close()
        '        dataCommand.Dispose()
        '        dataIRISConnection.Dispose()
        '        r = ex.Message
        '        Return r
        '    End Try
        'Catch ex As Exception
        '    r = ex.Message
        'End Try
        Return r
    End Function

    Public Function ExequteSQLquery_Cache(ByVal SQLq As String, ByRef myconstring As String) As String
        Dim r As String = "Query executed fine."
        'Try
        '    Dim dataCacheConnection As New InterSystems.Data.CacheClient.CacheConnection
        '    Dim dataCommand As New InterSystems.Data.CacheClient.CacheCommand
        '    Dim ds As New System.Data.DataSet
        '    If myconstring = String.Empty Then
        '        myconstring = InterSystems.Data.CacheClient.CacheConnection.ConnectDlg()
        '    End If
        '    dataCacheConnection = New InterSystems.Data.CacheClient.CacheConnection(myconstring)
        '    Try
        '        If dataCacheConnection.State = ConnectionState.Closed Then dataCacheConnection.Open()
        '        dataCommand.Connection = dataCacheConnection
        '        dataCommand.CommandType = CommandType.Text
        '        dataCommand.CommandTimeout = 300000
        '        dataCommand.CommandText = SQLq
        '        dataCommand.ExecuteNonQuery()
        '        dataCommand.Dispose()
        '        dataCacheConnection.Close()
        '    Catch ex As Exception
        '        'dataCommand.Connection.Close()
        '        dataCommand.Dispose()
        '        dataCacheConnection.Dispose()
        '        r = ex.Message
        '        Return r
        '    End Try
        'Catch ex As Exception
        '    r = ex.Message
        'End Try
        Return r
    End Function

    Public Function ExequteSQLquery_Oracle(ByVal SQLq As String, ByVal myconstring As String) As String
        Dim r As String = "Query executed fine."
        'Try
        '    Dim myConnection As Oracle.ManagedDataAccess.Client.OracleConnection
        '    Dim myCommand As New Oracle.ManagedDataAccess.Client.OracleCommand
        '    myConnection = New Oracle.ManagedDataAccess.Client.OracleConnection(myconstring)
        '    Try
        '        myCommand.Connection = myConnection
        '        myCommand.CommandType = CommandType.Text
        '        myCommand.CommandTimeout = 300000
        '        myCommand.CommandText = SQLq
        '        If myCommand.Connection.State = ConnectionState.Closed Then myCommand.Connection.Open()
        '        myCommand.ExecuteNonQuery()
        '        myCommand.Connection.Close()
        '        myCommand.Dispose()
        '        myConnection.Dispose()
        '    Catch ex As Exception
        '        myCommand.Connection.Close()
        '        myCommand.Dispose()
        '        myConnection.Dispose()
        '        r = ex.Message
        '        Return r
        '    End Try
        'Catch ex As Exception
        '    r = ex.Message
        'End Try
        Return r
    End Function
#End Region

End Module
