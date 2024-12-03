Imports System.Data.OleDb
Module sqlLoginModule
    Public errormgs As ErrorProvider

    Public Sub sqlIogin_Con(ByVal SqlLoginString As String)

        Try

            GetConnection()
            dbcommand = New OleDbCommand
            With dbcommand
                .Connection = dbconnect
                .CommandText = SqlLoginString
                .CommandType = CommandType.Text
            End With

            dbAdapter = New OleDbDataAdapter
            dbAdapter.SelectCommand = dbcommand
            dbdt = New DataTable
            dbAdapter.Fill(dbdt)

            'display a successfull message if record was saved successfull
            If frmlogin.txtUsername.Text = dbdt.Rows(0).Item("userName") And frmlogin.txtPassword.Text = dbdt.Rows(0).Item("ePassword") Then
                MsgBox("Welcome " & " " & dbdt.Rows(0).Item("userName"), MsgBoxStyle.Information, "Access Granted")

                Dim frmControl_Panel As New frmControl_Panel
                frmControl_Panel.Show()
                frmlogin.Hide()
            End If
        Catch ex As Exception
            'Display a failure messages if record was not saved
            MsgBox("Wrong Username and Password, Please Try Again ...!!!", MsgBoxStyle.Information, "Access Denied !!!")
            frmlogin.txtUsername.Focus()
            Exit Sub
            'rec_found = False
        Finally
            'dbcommand.Dispose()
            'dbAdapter.Dispose()
            'TerminateConn()

        End Try
        'Else


    End Sub

End Module
