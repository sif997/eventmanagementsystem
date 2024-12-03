Module ModuleCheckfields

    Public Sub clearfield(ByVal obj As Object, ByVal obj1 As Object)
        For Each txt As Control In obj.controls
            If txt.GetType Is GetType(TextBox) Then
                txt.Text = ""
            End If

        Next

        For Each cbo As Control In obj1.controls

            If cbo.GetType Is GetType(ComboBox) Then
                cbo.Text = ""

            End If

        Next

    End Sub

    Public Sub clearAllCombo(ByVal objcbo As Object)

        For Each cbo As Control In objcbo.controls

            If cbo.GetType Is GetType(ComboBox) Then
                cbo.Text = ""

            End If

        Next

    End Sub

    Public Sub clearAllTxtbo(ByVal txtobj As Object)

        For Each txt As Control In txtobj.controls
            If txt.GetType Is GetType(TextBox) Then
                txt.Text = ""
            End If

        Next

    End Sub

End Module
