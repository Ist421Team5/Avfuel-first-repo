Imports Excel = Microsoft.Office.Interop.Excel
Imports Microsoft.Office

Public Class dispatchTerminal
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Me.Hide()
        welcomeScreen.Show()
    End Sub



    'Text Box Layout
    '1
    '4

    '6
    '5

    '3
    '7

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet

        Label2.ForeColor = Color.Black
        Label3.ForeColor = Color.Black
        Label4.ForeColor = Color.Black
        TextBox1.Text = ""
        TextBox4.Text = ""
        TextBox6.Text = ""
        TextBox5.Text = ""
        TextBox3.Text = ""
        TextBox7.Text = ""

        Dim Row As Integer
        Dim t1 As Integer
        Dim t2 As Integer
        t1 = 2000
        t2 = 2000

        xlApp = New Excel.Application
        xlWorkBook = xlApp.Workbooks.Open("C:\Users\Sean\Documents\Visual Studio 2017\Projects\Avfuel System\Avfuel System\Resources\Avfuel_Database.xlsx")
        xlWorkSheet = xlWorkBook.Worksheets("sheet1")
        xlApp.DisplayAlerts = False

        Row = xlWorkSheet.UsedRange.Rows.Count
        For i = 1 To Row
            If xlWorkSheet.Cells(i + 1, 11).Value = "No" Then
                Label2.ForeColor = Color.DarkGray
                TextBox1.Text = xlWorkSheet.Cells(i + 1, 1).Value & " - " & xlWorkSheet.Cells(i + 1, 4).Value & " | " & xlWorkSheet.Cells(i + 1, 6).Value & " Gallons of " & xlWorkSheet.Cells(i + 1, 5).Value
                TextBox4.Text = "Customer: " & xlWorkSheet.Cells(i + 1, 2).Value & " " & xlWorkSheet.Cells(i + 1, 3).Value & " | " & xlWorkSheet.Cells(i + 1, 12).Value
                TextBox4.ForeColor = Color.DarkRed
                t1 = (i + 1)
                Exit For
            End If
        Next

        For i = t1 To Row
            If xlWorkSheet.Cells(i + 1, 11).Value = "No" Then
                Label3.ForeColor = Color.DarkGray
                TextBox6.Text = xlWorkSheet.Cells(i + 1, 1).Value & " - " & xlWorkSheet.Cells(i + 1, 4).Value & " | " & xlWorkSheet.Cells(i + 1, 6).Value & " Gallons of " & xlWorkSheet.Cells(i + 1, 5).Value
                TextBox5.Text = "Customer: " & xlWorkSheet.Cells(i + 1, 2).Value & " " & xlWorkSheet.Cells(i + 1, 3).Value & " | " & xlWorkSheet.Cells(i + 1, 12).Value
                TextBox5.ForeColor = Color.DarkRed
                t2 = (i + 1)
                Exit For
            End If
        Next

        For i = t2 To Row
            If xlWorkSheet.Cells(i + 1, 11).Value = "No" Then
                Label4.ForeColor = Color.DarkGray
                TextBox3.Text = xlWorkSheet.Cells(i + 1, 1).Value & " - " & xlWorkSheet.Cells(i + 1, 4).Value & " | " & xlWorkSheet.Cells(i + 1, 6).Value & " Gallons of " & xlWorkSheet.Cells(i + 1, 5).Value
                TextBox7.Text = "Customer: " & xlWorkSheet.Cells(i + 1, 2).Value & " " & xlWorkSheet.Cells(i + 1, 3).Value & " | " & xlWorkSheet.Cells(i + 1, 12).Value
                TextBox7.ForeColor = Color.DarkRed
                Exit For
            End If
        Next

        xlWorkBook.Close()
        xlApp.Quit()
        releaseObject(xlApp)
        releaseObject(xlWorkBook)
        releaseObject(xlWorkSheet)


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

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click

        ' Me.Hide()
        orderApproval.Show()

        orderApproval.Label1.Text = TextBox1.Text
        orderApproval.Label2.Text = TextBox4.Text

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click


        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet

        Dim Row As Integer
        Dim t1 As Integer
        Dim t2 As Integer
        t1 = 2000
        t2 = 2000

        xlApp = New Excel.Application
        xlWorkBook = xlApp.Workbooks.Open("C:\Users\Sean\Documents\Visual Studio 2017\Projects\Avfuel System\Avfuel System\Resources\Avfuel_Database.xlsx")
        xlWorkSheet = xlWorkBook.Worksheets("sheet1")
        xlApp.DisplayAlerts = False

        Row = xlWorkSheet.UsedRange.Rows.Count
        For i = 1 To Row
            If xlWorkSheet.Cells(i + 1, 11).Value = "Yes" Then
                Label2.ForeColor = Color.DarkGray
                TextBox1.Text = xlWorkSheet.Cells(i + 1, 1).Value & " - " & xlWorkSheet.Cells(i + 1, 4).Value & " | " & xlWorkSheet.Cells(i + 1, 6).Value & " Gallons of " & xlWorkSheet.Cells(i + 1, 5).Value
                TextBox4.Text = "Customer: " & xlWorkSheet.Cells(i + 1, 2).Value & " " & xlWorkSheet.Cells(i + 1, 3).Value & " | " & xlWorkSheet.Cells(i + 1, 12).Value
                TextBox4.ForeColor = Color.Green
                t1 = (i + 1)
                Exit For
            End If
        Next

        For i = t1 To Row
            If xlWorkSheet.Cells(i + 1, 11).Value = "Yes" Then
                Label3.ForeColor = Color.DarkGray
                TextBox6.Text = xlWorkSheet.Cells(i + 1, 1).Value & " - " & xlWorkSheet.Cells(i + 1, 4).Value & " | " & xlWorkSheet.Cells(i + 1, 6).Value & " Gallons of " & xlWorkSheet.Cells(i + 1, 5).Value
                TextBox5.Text = "Customer: " & xlWorkSheet.Cells(i + 1, 2).Value & " " & xlWorkSheet.Cells(i + 1, 3).Value & " | " & xlWorkSheet.Cells(i + 1, 12).Value
                TextBox5.ForeColor = Color.Green
                t2 = (i + 1)
                Exit For
            End If
        Next

        For i = t2 To Row
            If xlWorkSheet.Cells(i + 1, 11).Value = "Yes" Then
                Label4.ForeColor = Color.DarkGray
                TextBox3.Text = xlWorkSheet.Cells(i + 1, 1).Value & " - " & xlWorkSheet.Cells(i + 1, 4).Value & " | " & xlWorkSheet.Cells(i + 1, 6).Value & " Gallons of " & xlWorkSheet.Cells(i + 1, 5).Value
                TextBox7.Text = "Customer: " & xlWorkSheet.Cells(i + 1, 2).Value & " " & xlWorkSheet.Cells(i + 1, 3).Value & " | " & xlWorkSheet.Cells(i + 1, 12).Value
                TextBox7.ForeColor = Color.Green
                Exit For
            End If
        Next

        xlWorkBook.Close()
        xlApp.Quit()
        releaseObject(xlApp)
        releaseObject(xlWorkBook)
        releaseObject(xlWorkSheet)




    End Sub
End Class