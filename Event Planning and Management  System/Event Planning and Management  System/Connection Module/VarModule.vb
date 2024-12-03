Imports System.Data.OleDb
Module VarModule
    Public dbcommand As OleDbCommand
    Public dbconnect As OleDbConnection
    Public dbInsertQuery As String
    Public dbReader As OleDbDataReader
    Public dbAdapter As OleDbDataAdapter
    Public dbtable As DataTable
    Public dbset As DataSet
    Public countrow As Integer
    Public dbdt As DataTable
    Public dbds As DataSet
    Public rec_found As Boolean = False


    Public Sub GetConnection()
        'create a new connection string
        dbconnect = New OleDbConnection
        dbconnect.ConnectionString = AtteachDb
        If dbconnect.ConnectionString <> AtteachDb Then
            MsgBox("database connection failed or does not exist...", MsgBoxStyle.Critical + MsgBoxStyle.Information, "Try Again")
            Exit Sub

        ElseIf dbconnect.State = ConnectionState.Closed Then
            ' Open dbconnection

            ' MsgBox("connected to db", MsgBoxStyle.Exclamation)
            dbconnect.Open()
        Else
            MsgBox("Not Connected")
        End If


    End Sub

    Public Sub TerminateConn()

        If dbconnect.State = ConnectionState.Open Then
            'close the db connection
            dbconnect.Close()
        End If

    End Sub


End Module
