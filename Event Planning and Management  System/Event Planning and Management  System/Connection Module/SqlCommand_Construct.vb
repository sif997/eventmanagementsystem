Imports System.Data.OleDb
Module SqlCommand_Construct

    Public Sub sqlEnsertCommand(ByVal SqlEnter As String)
        If MsgBox("Are You Sure You want to Save...?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Save Record") = MsgBoxResult.Yes Then
            Try

                GetConnection()
                'dbconnect = New OleDbConnection
                'dbconnect.ConnectionString = AtteachDb
                'dbconnect.Open()
                dbcommand = New OleDbCommand
                With dbcommand
                    .Connection = dbconnect
                    .CommandText = SqlEnter
                    .CommandType = CommandType.Text
                End With
                dbAdapter = New OleDbDataAdapter
                dbAdapter.SelectCommand = dbcommand
                dbtable = New DataTable
                dbAdapter.Fill(dbtable)
                'countrow = dbtable.Rows.Count - 1

                MsgBox("Record Successfully Saved...!!!", MsgBoxStyle.Information)

                TerminateConn()

            Catch ex As Exception
                MsgBox("Error, Record not Successfully Saved...!!!", MsgBoxStyle.Information)
                ' MsgBox(ex.Message)

            Finally


            End Try
        Else

            Exit Sub
        End If

    End Sub

    Public Sub sqlDeleteCommand(ByVal SqlDeleter As String)
        If MsgBox("Are You Sure You Want to Delete This Record...?", MsgBoxStyle.YesNo + MsgBoxStyle.Question, "Deleting Record") = MsgBoxResult.Yes Then

            Try

                GetConnection()
                dbcommand = New OleDbCommand
                ' Create sqlcommand object
                With dbcommand
                    .Connection = dbconnect
                    .CommandText = SqlDeleter
                    .CommandType = CommandType.Text

                End With
                'creating a new dataAdapter
                dbAdapter = New OleDbDataAdapter
                dbAdapter.SelectCommand = dbcommand
                dbtable = New DataTable
                dbAdapter.Fill(dbtable)

                MsgBox("Record Successfully Deleted...!!!", MsgBoxStyle.Information)

            Catch ex As Exception

                MsgBox("Record Not Successfully Deleted...!!!", MsgBoxStyle.Information)

            Finally
                'dispose the connection and command objects
                dbcommand.Dispose()
                dbAdapter.Dispose()
                TerminateConn()

            End Try

        Else

            Exit Sub
        End If

    End Sub

    Public Sub sqlUpdateCommand(ByVal SqlUpdater As String)
        If MsgBox("Are You Sure to Update this Record...?", MsgBoxStyle.YesNo + MsgBoxStyle.Question, "Update Record") = MsgBoxResult.Yes Then
            Try

                GetConnection()
                dbcommand = New OleDbCommand
                With dbcommand
                    .Connection = dbconnect
                    .CommandText = SqlUpdater
                    .CommandType = CommandType.Text

                End With
                dbAdapter = New OleDbDataAdapter
                dbAdapter.SelectCommand = dbcommand
                dbtable = New DataTable
                dbAdapter.Fill(dbtable)
                'display a successfull message if record was saved successfull
                MsgBox("Record Successfully Updated...!!!", MsgBoxStyle.Information)

            Catch ex As Exception
                'Display a failure messages if record was not saved
                MsgBox("Record Not Successfully Updated...!!!", MsgBoxStyle.Information)

            Finally
                dbcommand.Dispose()
                dbAdapter.Dispose()
                TerminateConn()

            End Try
        Else
            Exit Sub

        End If


    End Sub


    Public Sub sqlFillQueryCommand(ByVal SqlFilledGrid As String)

        Try

            GetConnection()
            dbcommand = New OleDbCommand
            With dbcommand
                .Connection = dbconnect
                .CommandText = SqlFilledGrid
                .CommandType = CommandType.Text

            End With
            dbAdapter = New OleDbDataAdapter
            dbAdapter.SelectCommand = dbcommand
            dbdt = New DataTable
            dbAdapter.Fill(dbdt)
            
            'MsgBox("Record Successfully Saved...!!!", MsgBoxStyle.Information)
            TerminateConn()

        Catch ex As Exception

            MsgBox("Could not load data to the grid", MsgBoxStyle.Critical)
            'MsgBox("Error, Record not Successfully Saved...!!!", MsgBoxStyle.Information)
        Finally

        End Try

    End Sub

    Public Sub sqlgetQueryCommand(ByVal SqlfetcdGrid As String)

        Try

            GetConnection()
            dbcommand = New OleDbCommand
            With dbcommand
                .Connection = dbconnect
                .CommandText = SqlfetcdGrid
                .CommandType = CommandType.Text

            End With
            dbAdapter = New OleDbDataAdapter
            dbAdapter.SelectCommand = dbcommand
            dbdt = New DataTable
            dbAdapter.Fill(dbdt)

            TerminateConn()

        Catch ex As Exception

            MsgBox("Could not load data to the grid", MsgBoxStyle.Critical)
            'MsgBox("Error, Record not Successfully Saved...!!!", MsgBoxStyle.Information)
        Finally

        End Try

    End Sub

   


End Module

