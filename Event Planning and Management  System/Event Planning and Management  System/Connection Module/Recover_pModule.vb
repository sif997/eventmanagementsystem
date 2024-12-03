Imports System.Data.OleDb
Module Recover_pModule


    Public Sub SQLFetchRowQuery(ByVal SqlFetchRows As String)
        'If MsgBox("Are You Sure to Update this Record...?", MsgBoxStyle.YesNo + MsgBoxStyle.Question, "Update Record") = MsgBoxResult.Yes Then
        Try

            GetConnection()
            dbcommand = New OleDbCommand
            With dbcommand
                .Connection = dbconnect
                .CommandText = SqlFetchRows
                .CommandType = CommandType.Text

            End With
            dbAdapter = New OleDbDataAdapter
            dbAdapter.SelectCommand = dbcommand
            dbtable = New DataTable
            dbAdapter.Fill(dbtable)
            'display a successfull message if record was saved successfull


        Catch ex As Exception
            'Display a failure messages if record was not saved
            MsgBox("Record Not Found Or Does Not Exist...!!!", MsgBoxStyle.Information)
            rec_found = False
        Finally
            dbcommand.Dispose()
            dbAdapter.Dispose()
            TerminateConn()

        End Try
        'Else
        'Exit Sub

        ' End If
    End Sub


End Module
