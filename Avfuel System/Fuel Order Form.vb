Imports Excel = Microsoft.Office.Interop.Excel
Imports Microsoft.Office

Public Class fuelOrderForm
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        '    Dim objApp As Object
        '    Dim ResourcePath As String = "C:\Users\Sean\Documents\Visual Studio 2017\Projects\Avfuel System\Avfuel System\Resources\Avfuel_Database.xlsx"
        '    objApp = CreateObject("Excel.Application")
        '    objApp.WorkBooks.Open(ResourcePath)
        '    objApp.visible = False
        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet

        Dim Row As Integer
        ' Dim Column As Integer

        Dim s1 As String
        s1 = TextBox1.Text  'firstname
        Dim s2 As String
        s2 = TextBox2.Text  'lastname
        Dim s3 As String
        s3 = TextBox3.Text  'phone#
        Dim s4 As String
        s4 = TextBox4.Text  'email
        Dim s5 As String
        s5 = TextBox5.Text 'airport
        Dim s6 As String
        s6 = TextBox6.Text 'city
        Dim s7 As String
        s7 = TextBox7.Text 'state
        Dim s8 As String
        s8 = TextBox8.Text 'zip
        Dim s9 As String
        s9 = TextBox9.Text  'fuel type
        Dim s10 As String
        s10 = TextBox10.Text  'amount

        Dim d1 As String
        d1 = DateTimePicker1.Text

        xlApp = New Excel.Application
        xlWorkBook = xlApp.Workbooks.Open("C:\Users\Sean\Documents\Visual Studio 2017\Projects\Avfuel System\Avfuel System\Resources\Avfuel_Database.xlsx")
        xlWorkSheet = xlWorkBook.Worksheets("sheet1")
        xlApp.DisplayAlerts = False

        '  Column = xlWorkSheet.UsedRange.Rows.Count
        Row = xlWorkSheet.UsedRange.Rows.Count
        xlWorkSheet.Cells(Row + 1, 1) = ("00" + Row)  'order number
        xlWorkSheet.Cells(Row + 1, 2) = s1  'first name
        xlWorkSheet.Cells(Row + 1, 3) = s2  'last name
        xlWorkSheet.Cells(Row + 1, 12) = s3  'phone number
        xlWorkSheet.Cells(Row + 1, 13) = s4  'email
        xlWorkSheet.Cells(Row + 1, 4) = s5  'airport
        xlWorkSheet.Cells(Row + 1, 14) = s6  'city
        xlWorkSheet.Cells(Row + 1, 15) = s7  'state
        xlWorkSheet.Cells(Row + 1, 16) = s8  'zip
        xlWorkSheet.Cells(Row + 1, 5) = s9  'fuel type
        xlWorkSheet.Cells(Row + 1, 6) = s10  'amount in gallons
        xlWorkSheet.Cells(Row + 1, 7) = d1  'delivery date
        xlWorkSheet.Cells(Row + 1, 11) = "No"

        xlWorkBook.SaveAs("C:\Users\Sean\Documents\Visual Studio 2017\Projects\Avfuel System\Avfuel System\Resources\Avfuel_Database.xlsx")
        xlWorkBook.Close()
        xlApp.Quit()

        releaseObject(xlApp)
        releaseObject(xlWorkBook)
        releaseObject(xlWorkSheet)
        Me.Hide()
        MsgBox("Order Submitted Successfully")
        fuelRequest.Show()
    End Sub

    Private Sub releaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub


    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Hide()
        fuelRequest.Show()
    End Sub

    Private Sub Label10_Click(sender As Object, e As EventArgs) Handles Label10.Click

    End Sub

    Private Sub TextBox6_TextChanged(sender As Object, e As EventArgs) Handles TextBox6.TextChanged

    End Sub

    Private Sub Label7_Click(sender As Object, e As EventArgs) Handles Label7.Click

    End Sub
End Class