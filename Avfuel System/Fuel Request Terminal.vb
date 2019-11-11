Public Class fuelRequest
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        fuelOrderForm.Show()
        Me.Hide()
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Me.Hide()
        welcomeScreen.Show()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim Outl As Object
        Outl = CreateObject("Outlook.Application")
        If Outl IsNot Nothing Then
            Dim omsg As Object
            omsg = Outl.CreateItem(0) '=Outlook.OlItemType.olMailItem'
            'set message properties here...'
            omsg.to = "dispatch@avfuel.com"
            omsg.subject = "UPDATE to My Order"
            omsg.body = "Dear Dispatch," & vbNewLine & vbNewLine & "Order Number:" & vbNewLine & "Customer Name:" & vbNewLine & "Customer Phone Number:" & vbNewLine & vbNewLine & "Question:"
            omsg.Display(True) 'will display message to user
        End If
    End Sub
End Class