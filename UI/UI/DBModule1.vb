Imports System.Data
Imports System.Data.Odbc
Imports System.Data.OleDb '关于数据的命名空间，在使用数据库的时候用到
Imports System.Data.SqlClient
Imports System.Text
Imports System.IO
Imports System.Data.Common



Module DBModule
#Region "定义变量,外部函数"

    Public tagIDarry() As Integer
    Public SVSUSER As String = "sa"
    Public SVSPSD As String = "123456"
    Public SVSNAME As String = "localhost"
    Public SVSPATH As String
    Public DBTYPE As String = "sql"
    Public Const null As Object = Nothing
    'Public accessPath As String
    'Public mysqlPath As String
    'Public sqlpath As String
#End Region

#Region "关于表的函数"
   


    Public Sub cbxAddItems(ByRef cbx As ComboBox, ByRef data As String)  '把数据库里面的一列放到一个下拉框ComboBox中

        cbx.Items.Clear()
        Dim mydatatable As DataTable '数据表，可以存放一个数据表
        Dim mydatarow As DataRow    '数据列，可以存放数据表中的一列
        Dim str As String
        Dim mystr() As String = data.Split(".") '把mystr分为几部分

        str = "select distinct " & mystr(1) & " from " & mystr(0) 'distinct就是选取过程中重复的部分只选取一次，在本程序里面是没有作用的

        mydatatable = GetData(str) '得到所需要的数据表

        For Each mydatarow In mydatatable.Rows '整体句型

            cbx.Items.Add(mydatarow.Item(mystr(1))) '把所得数据表里面的内容增加到combox框中
            ' Debug.WriteLine(mydatarow.Item("segment"))
        Next
    End Sub

    Public Sub cbxAddItems(ByRef cbx As ListBox, ByRef data As String)  '把数据库里面的一列放到一个下拉框ComboBox中

        cbx.Items.Clear()
        Dim mydatatable As DataTable '数据表，可以存放一个数据表
        Dim mydatarow As DataRow    '数据列，可以存放数据表中的一列
        Dim str As String
        Dim mystr() As String = data.Split(".") '把mystr分为几部分

        str = "select distinct " & mystr(1) & " from " & mystr(0) 'distinct就是选取过程中重复的部分只选取一次，在本程序里面是没有作用的

        mydatatable = GetData(str) '得到所需要的数据表

        For Each mydatarow In mydatatable.Rows '整体句型

            cbx.Items.Add(mydatarow.Item(mystr(1))) '把所得数据表里面的内容增加到combox框中
            ' Debug.WriteLine(mydatarow.Item("segment"))
        Next
    End Sub

    '可以查 整形？文本型？
    Public Function IsIntable(ByVal id As String, ByVal TableColumn As String, Optional ByVal isInt As Boolean = False) As Boolean
        '此函数可优化，datatable类？
        Dim mystr() As String = TableColumn.Split(".")
        Dim dt As New DataTable
        If isInt Then
            dt = GetData("select distinct " & mystr(1) & " from " & mystr(0) & " where " & mystr(1) & " = " & id & " order by " & mystr(1))
        Else
            dt = GetData("select distinct " & mystr(1) & " from " & mystr(0) & " where " & mystr(1) & " = '" & id & "' order by " & mystr(1))
        End If

        If dt.Rows.Count > 0 Then
            Return True
        End If
        Return False

    End Function

    Public Function getEN(ByVal data As Integer, ByVal table As String, ByVal iniCol As String, Optional ByVal finalCol As String = "id") As String
        Dim err As New Exception
        Dim dt As DataTable = GetData("select " & finalCol & " from " & table & " where " & iniCol & "=" & data)
        ' MsgBox(dt.Rows.Count & "  " & dt.Columns.Count)
        If dt.Rows.Count <> 1 Then
            Throw err
        End If
        Return dt.Rows(0)(0)
    End Function

    Public Function getEN(ByVal data As String, ByVal table As String, ByVal iniCol As String, Optional ByVal finalCol As String = "id") As String
        '
        Dim err As New Exception
        Dim dt As DataTable = GetData("select " & finalCol & " from " & table & " where " & iniCol & "='" & data & "'")
        ' MsgBox(dt.Rows.Count & "  " & dt.Columns.Count)
        If dt Is Nothing Then Throw New NullReferenceException()
        If dt.Rows.Count <> 1 OrElse dt.Columns.Count <> 1 Then
            Throw err
        End If
        Return dt.Rows(0)(0)

    End Function

    Public Function notIntable(ByVal data As String, ByVal TableColumn As String) As Boolean
        Return Not IsIntable(data, TableColumn)
    End Function

#End Region

#Region "SQL增删改查"

#Region "单表操作"


#End Region


#Region " 数据库增删改查 "

   
#End Region











#End Region

#Region "基础函数"

    Public Function ExcSQL(ByVal sqlCmd As String) As Integer
        Dim Conn As DbConnection = Nothing
        Dim Cmd As DbCommand = Nothing
        Dim ret As Integer
        Try
            Select Case DBTYPE.ToLower
                Case "access", "accdb"
                    Dim connectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & _
                     SVSPATH & ";Mode=ReadWrite|Share Deny None;Persist Security Info=False"
                    Conn = New OleDbConnection(connectionString)
                    Conn.Open()
                    Cmd = New OleDbCommand(sqlCmd, Conn)
                Case "mdb"

                    Dim connectionString As String = "Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=" & SVSPATH
                    Conn = New OdbcConnection(connectionString)
                    Conn.Open()
                    Cmd = New OdbcCommand(sqlCmd, Conn)

                Case "sql"
                    Dim connectionString As String = "Data Source=" & SVSNAME & ";Initial Catalog=report;User ID=" & SVSUSER & ";Password=" & SVSPSD
                    Conn = New SqlConnection(connectionString)
                    Conn.Open()
                    Cmd = New SqlCommand(sqlCmd, Conn)
            End Select
            ret = Cmd.ExecuteNonQuery
            Cmd.Dispose() : Cmd = Nothing
            Conn.Close() : Conn.Dispose() : Conn = Nothing
            Return ret
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "系统消息")
            'MsgBox(Err.Description)
            If Conn.State = ConnectionState.Open Then Conn.Close()
            Conn.Dispose()
            Return -1
        End Try



    End Function
    '获取数据库中的数据

    Public Function GetData(ByVal sqlCmd As String) As DataTable

        Dim Conn As DbConnection = Nothing
        Dim Cmd As DbCommand = Nothing
        Dim Adapter As DbDataAdapter = Nothing
        Dim table As DataTable = Nothing

        Try
            Select Case DBTYPE.ToLower
                Case "access", "accdb"
                    Dim connectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & _
                          SVSPATH & ";Mode=ReadWrite|Share Deny None;Persist Security Info=False"
                    Conn = New OleDbConnection(connectionString)
                    Conn.Open()
                    Cmd = New OleDbCommand(sqlCmd, Conn)
                    Adapter = New OleDbDataAdapter(Cmd)
                Case "mdb"
                    '              Dim connectionString As String = "Provider=Microsoft.ACE.OLEDB.4.0;Data Source=" & _
                    'SVSPATH & ";Mode=ReadWrite|Share Deny None;Persist Security Info=False"
                    '              Conn = New OleDbConnection(connectionString)
                    '              Conn.Open()
                    '              Cmd = New OleDbCommand(sqlCmd, Conn)
                    '              Adapter = New OleDbDataAdapter(Cmd)
                    'Dim connectionString As String = "Driver={Microsoft Access Driver (*.mdb)};Dbq=" & _
                    '        APP_PATH() & "report.mdb"
                    Dim connectionString As String = "Driver={ Microsoft Access Driver (*.mdb, *.accdb)};Dbq =" & SVSPATH
                    'Dim connectionString As String = "Driver={Microsoft Access Driver (*.mdb ,*.accdb)};Dbq=" & _
                    'SVSPATH
                    Conn = New OdbcConnection(connectionString)
                    Conn.Open()
                    Cmd = New OdbcCommand(sqlCmd, Conn)
                    Adapter = New OdbcDataAdapter(Cmd)
                Case "sql"
                    Dim connectionString As String = "Data Source=" & SVSNAME & ";Initial Catalog=report;User ID=" & SVSUSER & ";Password=" & SVSPSD
                    Conn = New SqlConnection(connectionString)
                    Conn.Open()
                    Cmd = New SqlCommand(sqlCmd, Conn)
                    Adapter = New SqlDataAdapter(Cmd)
            End Select
            table = New DataTable
            Adapter.Fill(table)
            If Conn.State = ConnectionState.Open Then
                Conn.Close() : Conn.Dispose() : Conn = Nothing
            End If
            Cmd.Dispose() : Cmd = Nothing
            Adapter.Dispose() : Adapter = Nothing
            Return table

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "系统消息")
            'MsgBox(Err.Description)
            If Conn.State = ConnectionState.Open Then
                Conn.Close() : Conn.Dispose() : Conn = Nothing
                Cmd.Dispose() : Cmd = Nothing
                Adapter.Dispose() : Adapter = Nothing
            End If
            Return Nothing
        Finally
            If table IsNot Nothing Then
                table.Dispose()
            End If
            table = Nothing
        End Try

    End Function
    'current path
    Public Function APP_PATH() As String
        Return System.AppDomain.CurrentDomain.BaseDirectory
    End Function

#End Region



End Module