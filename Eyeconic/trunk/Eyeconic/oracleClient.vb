
Imports Oracle.DataAccess.Client
Imports System.Threading

Public Class oracleClient
    Implements IDisposable

#Region "Properties and Members"

    Private _connectionString As String
    Public Property connectionString() As String
        Get
            Return _connectionString
        End Get
        Set(ByVal value As String)
            _connectionString = value
        End Set
    End Property


    Private _oraCon As Oracle.DataAccess.Client.OracleConnection
    Public Property oraCon() As OracleConnection
        Get
            Return _oraCon
        End Get
        Set(ByVal value As OracleConnection)
            _oraCon = value
        End Set
    End Property

    Private _T As OracleTransaction
    Private _TDAs As New Dictionary(Of String, OracleDataAdapter)

#End Region


    'Create and open Oracle connection using optional provided connection string
    Sub New(Optional ByVal conString As String = "Data Source=TST;User ID=TST;Password=TST;pooling=true")
        connectionString = conString
        Dim x As String = System.Environment.UserName
        oraCon = New OracleConnection(connectionString)
        oraCon.Open()
    End Sub

    Public Overloads Sub Dispose() Implements IDisposable.Dispose
        If _TDAs IsNot Nothing Then
            For Each tda As OracleDataAdapter In _TDAs.Values
                tda.Dispose()
            Next
        End If
        If _T IsNot Nothing Then
            _T.Rollback()
            _T.Dispose()
        End If
        If oraCon IsNot Nothing Then
            oraCon.Close()
            oraCon.Dispose()
        End If
    End Sub

    Private Function GetDataAdapter(ByVal selectSQL As String, ByVal ParamArray PARMs() As Object) As OracleDataAdapter
        Dim da As OracleDataAdapter = New OracleDataAdapter(selectSQL, oraCon)

        If PARMs IsNot Nothing Then
            CreateParameters(da, PARMs)
        End If

        Return da
    End Function

    Function GetDataTable(ByVal selectSQL As String, ByVal ParamArray PARMs() As Object) As DataTable
        Dim tda As OracleDataAdapter = GetDataAdapter(selectSQL, PARMs)
        Dim dt As New DataTable()

        tda.Fill(dt)

        Return dt
    End Function

    Function GetDataValue(ByVal selectSQL As String, ByVal ParamArray PARMs() As Object) As Object
        Using cmd As New OracleCommand(selectSQL, oraCon)
            If PARMs IsNot Nothing Then
                CreateParameters(cmd, PARMs)
            End If
            GetDataValue = cmd.ExecuteScalar
        End Using
    End Function

    Public Sub ExecuteSQL(ByVal sqlToExecute As String, ByVal ParamArray PARMs() As Object)
        Using cmd As New OracleCommand(sqlToExecute, oraCon)
            If PARMs IsNot Nothing Then
                CreateParameters(cmd, PARMs)
            End If
            cmd.ExecuteNonQuery()
        End Using
    End Sub

    Public Sub ExecuteSP(ByVal spName As String, ByVal parmNAMES() As String, ByVal ParamArray PARMs() As Object)
        Using cmd As New OracleCommand(spName, oraCon)
            cmd.CommandType = CommandType.StoredProcedure
            If PARMs IsNot Nothing Then
                CreateParameters(cmd, parmNAMES, PARMs)
            End If
            cmd.ExecuteNonQuery()
        End Using
    End Sub

    Function ExecuteSF(ByVal sfName As String, ByVal parmNAMES() As String, ByVal ParamArray PARMs() As Object) As String
        Using cmd As New OracleCommand(sfName, oraCon)
            cmd.CommandType = CommandType.StoredProcedure
            If PARMs IsNot Nothing Then
                CreateParameters(cmd, parmNAMES, PARMs)
            End If

            cmd.Parameters.Add("returnValue", OracleDbType.Varchar2, 100, Nothing, ParameterDirection.ReturnValue)
            cmd.ExecuteNonQuery()
            If cmd.Parameters("returnValue").Value Is Nothing OrElse cmd.Parameters("returnValue").Value Is DBNull.Value OrElse DirectCast(cmd.Parameters("returnValue").Value, Oracle.DataAccess.Types.OracleString).IsNull Then
                ExecuteSF = ""
            Else
                ExecuteSF = cmd.Parameters("returnValue").Value.ToString
            End If
        End Using
    End Function

    Sub BeginTrans()
        _T = oraCon.BeginTransaction()
    End Sub

    Sub Commit()
        If _T IsNot Nothing Then
            _T.Commit()
            _T.Dispose()
            _T = Nothing
        Else
            Throw New Exception("No transaction to commit.")
        End If
    End Sub

    Sub Rollback()
        If _T IsNot Nothing Then
            _T.Rollback()
        Else
            Throw New Exception("No transaction to roll back.")
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
