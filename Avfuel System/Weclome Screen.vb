Public Class welcomeScreen
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        If TextBox2.Text = "customer" And TextBox3.Text = "customer" Then
            fuelRequest.Show()
            Me.Hide()
        ElseIf TextBox2.Text = "dispatch" And TextBox3.Text = "dispatch" Then
            dispatchTerminal.Show()
            Me.Hide()
        ElseIf TextBox2.Text = "hauler" And TextBox3.Text = "hauler" Then
            orderTerminal.WebBrowser1.Hide()
            orderTerminal.Show()
            Me.Hide()
        Else
            MsgBox("Wrong Username or Password, Please Try Again")
            TextBox3.Text = ""
        End If




    End Sub



End Class