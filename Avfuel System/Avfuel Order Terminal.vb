Imports Excel = Microsoft.Office.Interop.Excel
Imports Microsoft.Office

Public Class orderTerminal
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        WebBrowser1.Hide()
        'Text Box Layout
        '2
        '1
        '4
        '6
        '5
        '3
        '7

        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet

        Label2.ForeColor = Color.Black
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox4.Text = ""
        TextBox6.Text = ""
        TextBox5.Text = ""
        TextBox3.Text = ""
        TextBox7.Text = ""

        Dim Row As Integer

        xlApp = New Excel.Application
        xlWorkBook = xlApp.Workbooks.Open("C:\Users\Sean\Documents\Visual Studio 2017\Projects\Avfuel System\Avfuel System\Resources\Avfuel_Database.xlsx")
        xlWorkSheet = xlWorkBook.Worksheets("sheet1")
        xlApp.DisplayAlerts = False

        Row = xlWorkSheet.UsedRange.Rows.Count
        For i = 1 To Row
            If xlWorkSheet.Cells(i + 1, 11).Value = "No" Then
                Label2.ForeColor = Color.DarkGray
                Label2.Text = "CURRENT ASSIGNED ORDER"
                'Destination
                TextBox2.Text = "Destination: " & xlWorkSheet.Cells(i + 1, 4).Value
                TextBox2.ForeColor = Color.White
                'Order Details
                TextBox1.Text = "Order Details: #" & xlWorkSheet.Cells(i + 1, 1).Value & " | " & xlWorkSheet.Cells(i + 1, 6).Value & " Gallons of " & xlWorkSheet.Cells(i + 1, 5).Value
                TextBox1.ForeColor = Color.White
                'Customer Information
                TextBox4.Text = "Customer: " & xlWorkSheet.Cells(i + 1, 2).Value & " " & xlWorkSheet.Cells(i + 1, 3).Value
                TextBox4.ForeColor = Color.White
                'Delivery Date
                TextBox6.Text = "Delivery Date: " & xlWorkSheet.Cells(i + 1, 7).Value
                TextBox6.ForeColor = Color.White
                'Email
                TextBox5.Text = "Email Address: " & xlWorkSheet.Cells(i + 1, 13).Value
                TextBox5.ForeColor = Color.White
                'Email
                TextBox3.Text = "Phone Number: " & xlWorkSheet.Cells(i + 1, 12).Value
                TextBox3.ForeColor = Color.White
                Exit For
            End If
        Next


        xlWorkBook.Close()
        xlApp.Quit()
        releaseObject(xlApp)
        releaseObject(xlWorkBook)
        releaseObject(xlWorkSheet)

    End Sub



    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Me.Hide()
        welcomeScreen.Show()
    End Sub

    Private Sub releaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
            GC.WaitForPendingFinalizers()
        End Try
    End Sub


    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim zip As Integer
        Dim state As String
        Dim city As String

        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet

        Dim Row As Integer

        xlApp = New Excel.Application
        xlWorkBook = xlApp.Workbooks.Open("C:\Users\Sean\Documents\Visual Studio 2017\Projects\Avfuel System\Avfuel System\Resources\Avfuel_Database.xlsx")
        xlWorkSheet = xlWorkBook.Worksheets("sheet1")
        xlApp.DisplayAlerts = False

        Row = xlWorkSheet.UsedRange.Rows.Count
        For i = 1 To Row
            If xlWorkSheet.Cells(i + 1, 11).Value = "No" Then
                city = xlWorkSheet.Cells(i + 1, 14).Value
                state = xlWorkSheet.Cells(i + 1, 15).Value
                city = city.Replace(" ", "-")
                zip = xlWorkSheet.Cells(i + 1, 16).Value
                Exit For
            End If
        Next



        xlWorkBook.Close()
        xlApp.Quit()
        releaseObject(xlApp)
        releaseObject(xlWorkBook)
        releaseObject(xlWorkSheet)

        WebBrowser1.Show()
        WebBrowser1.Navigate("https://www.wunderground.com/weather/us/" & state & "/" & city & "/" & zip)
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        WebBrowser1.Hide()

        Dim zip As Integer
        Dim state As String
        Dim city As String

        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet

        Dim Row As Integer

        xlApp = New Excel.Application
        xlWorkBook = xlApp.Workbooks.Open("C:\Users\Sean\Documents\Visual Studio 2017\Projects\Avfuel System\Avfuel System\Resources\Avfuel_Database.xlsx")
        xlWorkSheet = xlWorkBook.Worksheets("sheet1")
        xlApp.DisplayAlerts = False

        Row = xlWorkSheet.UsedRange.Rows.Count
        For i = 1 To Row
            If xlWorkSheet.Cells(i + 1, 11).Value = "No" Then
                city = xlWorkSheet.Cells(i + 1, 14).Value
                state = xlWorkSheet.Cells(i + 1, 15).Value
                city = city.Replace(" ", "+")
                zip = xlWorkSheet.Cells(i + 1, 16).Value
                Exit For
            End If
        Next

        xlWorkBook.Close()
        xlApp.Quit()
        releaseObject(xlApp)
        releaseObject(xlWorkBook)
        releaseObject(xlWorkSheet)

        googleMaps.WebBrowser1.Show()
        googleMaps.WebBrowser1.Navigate("https://www.google.com/maps/dir//" & city & ",+" & state)

        googleMaps.Show()

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        WebBrowser1.Hide()

        Dim email As String
        Dim firstName As String
        Dim lastName As String
        Dim order As Integer
        Dim ddate As String

        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet

        Dim Row As Integer
        firstName = ""
        lastName = ""
        lastName = ""
        email = ""
        ddate = ""

        xlApp = New Excel.Application
        xlWorkBook = xlApp.Workbooks.Open("C:\Users\Sean\Documents\Visual Studio 2017\Projects\Avfuel System\Avfuel System\Resources\Avfuel_Database.xlsx")
        xlWorkSheet = xlWorkBook.Worksheets("sheet1")
        xlApp.DisplayAlerts = False

        Row = xlWorkSheet.UsedRange.Rows.Count
        For i = 1 To Row
            If xlWorkSheet.Cells(i + 1, 11).Value = "No" Then
                email = xlWorkSheet.Cells(i + 1, 13).Value
                order = xlWorkSheet.Cells(i + 1, 1).Value
                firstName = xlWorkSheet.Cells(i + 1, 2).Value
                lastName = xlWorkSheet.Cells(i + 1, 3).Value
                ddate = xlWorkSheet.Cells(i + 1, 7).Value
            End If
        Next

        xlWorkBook.Close()
        xlApp.Quit()
        releaseObject(xlApp)
        releaseObject(xlWorkBook)
        releaseObject(xlWorkSheet)

        Dim Outl As Object
        Outl = CreateObject("Outlook.Application")
        If Outl IsNot Nothing Then
            Dim omsg As Object
            omsg = Outl.CreateItem(0) '=Outlook.OlItemType.olMailItem'
            'set message properties here...'
            omsg.to = email
            omsg.subject = "UPDATE to Order# " & order & " for " & ddate
            omsg.body = "Dear " & firstName & " " & lastName & ","
            omsg.Display(True) 'will display message to user
        End If
    End Sub
End Class