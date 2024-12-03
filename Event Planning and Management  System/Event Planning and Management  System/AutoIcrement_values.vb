Imports System.Data.OleDb
Module AutoIcrement_values
    Dim i As Integer
    Dim reducme As Integer
    Dim getidValue As Integer
    Dim R As OleDbDataReader
    Dim IDIncremen As String = "0000"


    Public Sub Readmy_id_Row(ByVal getmy_ID_values As String, ByVal txt As TextBox)

        Try
            GetConnection()
            dbcommand = New OleDbCommand(getmy_ID_values, dbconnect)
            dbAdapter = New OleDbDataAdapter
            dbAdapter.SelectCommand = dbcommand
            dbdt = New DataTable
            dbAdapter.Fill(dbdt)

            Dim maxrow As Integer = dbdt.Rows.Count - 1

            ' i = dbdt.Rows(0).Item("customerID") + 1

            'reducme = dbdt.Rows(0).Item("customerID") - 1
            With txt

                If maxrow > -1 Then
                    .Text = "000" & dbdt.Rows(maxrow).Item(0)

                Else
                    .Text = "000"
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        TerminateConn()

    End Sub


End Module
