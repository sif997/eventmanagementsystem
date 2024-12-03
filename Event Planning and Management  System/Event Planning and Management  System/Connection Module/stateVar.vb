Imports System.Data.OleDb
Module stateVar
    Dim mystate As String
    Dim stateConn As OleDbConnection
    Dim stateCmd As OleDbCommand
    Dim stateAdapt As OleDbDataAdapter
    Dim ReadState As OleDbDataReader
    Dim stateds As DataSet



    Public Sub getstateRow(ByVal mystate As String)

        stateConn = New OleDbConnection
        stateConn.ConnectionString = AtteachDb
        stateConn.Open()
        stateCmd = New OleDbCommand
        With stateCmd
            .CommandText = mystate
            .CommandType = CommandType.Text
            .Connection = stateConn
        End With
        stateAdapt = New OleDbDataAdapter
        stateAdapt.SelectCommand = stateCmd
        dbds = New DataSet
        stateAdapt.Fill(dbds, "")
       
    End Sub

End Module
