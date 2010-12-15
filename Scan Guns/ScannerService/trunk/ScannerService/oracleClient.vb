Imports Oracle.DataAccess.Client

Public Class oracleClient
    Implements IDisposable

#Region "Properties and Members"
    Default Public ReadOnly Property Item(ByVal tableName As String) As DataTable
        Get
            Return dst.Tables(tableName)
        End Get
    End Property


    Dim _connectionString As String
    Public Property connectionString() As String
        Get
            Return _connectionString
        End Get
        Set(ByVal value As String)
            _connectionString = value
        End Set
    End Property

    Dim _errorText As String
    Public Property errorText() As String
        Get
            Return _errorText
        End Get
        Set(ByVal value As String)
            _errorText = value
        End Set
    End Property

    Dim _oracon As OracleConnection
    Public Property oraCon() As OracleConnection
        Get
            Return _oracon
        End Get
        Set(ByVal value As OracleConnection)
            _oracon = value
        End Set
    End Property

    Dim _dst As DataSet
    Public ReadOnly Property dst() As DataSet
        Get
            Return _dst
        End Get
    End Property

    Public ReadOnly Property dt(ByVal tableName As String) As DataTable
        Get
            Return _dst.Tables(tableName)
        End Get
    End Property

    Dim _T As OracleTransaction
    Dim _TDAs As New Dictionary(Of String, OracleDataAdapter)

#End Region


    'Create and open Oracle connection using optional provided connection string
    Sub New(Optional ByVal conString As String = "Data Source=ODG;User ID=ODG;Password=ODG;pooling=true")
        connectionString = conString
        errorText = ""
        _dst = New DataSet()

        Try
            oraCon = New OracleConnection(_connectionString)
            oraCon.Open()
        Catch ex As Exception
            _errorText &= "NEW: " & ex.Message & vbCr
        End Try
    End Sub

    Public Overloads Sub Dispose() Implements IDisposable.Dispose
        If _TDAs IsNot Nothing Then
            For Each tda As OracleDataAdapter In _TDAs.Values
                tda.Dispose()
            Next
        End If
        If dst IsNot Nothing Then
            dst.Dispose()
        End If
        If _T IsNot Nothing Then
            _T.Dispose()
        End If
        If oraCon IsNot Nothing Then
            oraCon.Close()
            oraCon.Dispose()
        End If
    End Sub


    Private Function GetDataAdapter(ByVal selectSQL As String, ByVal forUpdate As Boolean, ByVal ParamArray PARMs() As Object) As OracleDataAdapter
        Dim da As OracleDataAdapter = Nothing
        Try
            da = New OracleDataAdapter(selectSQL, oraCon)
            If PARMs IsNot Nothing Then
                CreateParameters(da, PARMs)
            End If
        Catch ex As Exception
            errorText &= "GDA: " & ex.Message & vbCr
        End Try
        Return da
    End Function

    Sub GetDataTable(ByVal tableName As String, ByVal selectSQL As String, ByVal forupdate As Boolean, ByVal ParamArray PARMs() As Object)
        If selectSQL = "" Then
            selectSQL = "SELECT * FROM " & tableName
        End If
        Try
            _TDAs.Add(tableName, GetDataAdapter(selectSQL, forupdate, PARMs))
            Dim dt As New DataTable(tableName)
            _TDAs(tableName).Fill(dt)
            If forupdate Then
                Dim cmdBuilder As New Oracle.DataAccess.Client.OracleCommandBuilder(_TDAs(tableName))
            End If
            dst.Tables.Add(dt)
        Catch ex As Exception
            errorText = "GDT: " & ex.Message & vbCr
        End Try
    End Sub

    Function GetDataValue(ByVal selectSQL As String, ByVal ParamArray PARMs() As Object) As Object
        Using cmd As New OracleCommand(selectSQL, oraCon)
            If PARMs IsNot Nothing Then
                CreateParameters(cmd, PARMs)
            End If
            GetDataValue = cmd.ExecuteScalar
        End Using
    End Function

    Sub ExecuteSP(ByVal spName As String, ByVal parmNAMES() As String, ByVal ParamArray PARMs() As Object)
        Using cmd As New OracleCommand(spName, oraCon)
            cmd.CommandType = CommandType.StoredProcedure
            If PARMs IsNot Nothing Then
                CreateParameters(cmd, parmNAMES, PARMs)
            End If
            cmd.ExecuteNonQuery()
        End Using
    End Sub

    Function ExecuteSF(ByVal sfName As String, ByVal parmNAMES() As String, ByVal ParamArray PARMs() As Object)
        Using cmd As New OracleCommand(sfName, oraCon)
            cmd.CommandType = CommandType.StoredProcedure
            If PARMs IsNot Nothing Then
                CreateParameters(cmd, parmNAMES, PARMs)
            End If
            cmd.Parameters.Add("returnValue", OracleDbType.Varchar2, 100, Nothing, ParameterDirection.ReturnValue)
            cmd.ExecuteNonQuery()
            ExecuteSF = cmd.Parameters("returnValue").Value.ToString
        End Using
    End Function

    Sub Update(ByVal tableName As String)
        _TDAs(tableName).Update(dst.Tables(tableName))
    End Sub

    Sub BeginTrans()
        If _T Is Nothing Then
            _T = oraCon.BeginTransaction()
        Else
            errorText = "Already in a transaction."
        End If
    End Sub

    Sub Commit()
        If _T IsNot Nothing Then
            _T.Commit()
            _T.Dispose()
            _T = Nothing
        Else
            errorText = "No transaction to commit."
        End If
    End Sub


    Private Sub CreateParameters( _
        ByRef da As OracleDataAdapter, _
        ByVal ParamArray PARMS() As Object)
        CreateParameters(da.SelectCommand, PARMS)
    End Sub

    Private Sub CreateParameters( _
        ByRef cmd As OracleCommand, _
        ByVal ParamArray PARMS() As Object)

        With cmd.Parameters
            For i As Integer = 1 To PARMS.Length
                Dim COLUMN_NAME As String = "PARM" & CStr(i)

                If PARMS(i - 1).GetType() Is Type.GetType("System.String") Then
                    .Add(COLUMN_NAME, OracleDbType.Varchar2, ParameterDirection.Input)
                ElseIf PARMS(i - 1).GetType() Is Type.GetType("System.Int64") Or PARMS(i - 1).GetType() Is Type.GetType("System.Int32") Or PARMS(i - 1).GetType() Is Type.GetType("System.Int16") Then
                    .Add(COLUMN_NAME, OracleDbType.Long, ParameterDirection.Input)
                ElseIf PARMS(i - 1).GetType() Is Type.GetType("System.Decimal") Then
                    .Add(COLUMN_NAME, OracleDbType.Double, ParameterDirection.Input)
                ElseIf PARMS(i - 1).GetType() Is Type.GetType("System.DateTime") Then
                    .Add(COLUMN_NAME, OracleDbType.Date, ParameterDirection.Input)
                Else
                    .Add(COLUMN_NAME, OracleDbType.Varchar2, ParameterDirection.Input)
                End If

                If PARMS Is Nothing OrElse PARMS.Length < i Then
                    .Item(COLUMN_NAME).Value = System.DBNull.Value
                Else
                    .Item(COLUMN_NAME).Value = PARMS(i - 1)
                End If
            Next
        End With
    End Sub

    Private Sub CreateParameters( _
     ByRef cmd As OracleCommand, _
     ByVal parmNAMES() As String, _
     ByVal ParamArray PARMS() As Object)

        cmd.BindByName = True
        With cmd.Parameters
            For i As Integer = 1 To PARMS.Length
                Dim COLUMN_NAME As String = parmNAMES(i - 1)

                If PARMS(i - 1).GetType() Is GetType(System.String) Then
                    .Add(COLUMN_NAME, OracleDbType.Varchar2, ParameterDirection.Input)
                ElseIf PARMS(i - 1).GetType() Is GetType(System.Int64) Or PARMS(i - 1).GetType() Is GetType(System.Int32) Or PARMS(i - 1).GetType() Is GetType(System.Int16) Then
                    .Add(COLUMN_NAME, OracleDbType.Long, ParameterDirection.Input)
                ElseIf PARMS(i - 1).GetType() Is GetType(System.Decimal) Then
                    .Add(COLUMN_NAME, OracleDbType.Double, ParameterDirection.Input)
                ElseIf PARMS(i - 1).GetType() Is GetType(System.DateTime) Then
                    .Add(COLUMN_NAME, OracleDbType.Date, ParameterDirection.Input)
                Else
                    .Add(COLUMN_NAME, OracleDbType.Varchar2, ParameterDirection.Input)
                End If

                If PARMS Is Nothing OrElse PARMS.Length < i Then
                    .Item(COLUMN_NAME).Value = System.DBNull.Value
                Else
                    .Item(COLUMN_NAME).Value = PARMS(i - 1)
                End If
            Next
        End With
    End Sub


End Class
