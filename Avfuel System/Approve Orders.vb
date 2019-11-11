Imports Excel = Microsoft.Office.Interop.Excel
Imports Microsoft.Office

Public Class orderApproval


    Private Sub orderApproval_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

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
                xlWorkSheet.Cells(i + 1, 11).Value = "Yes"
                Exit For
            End If
        Next

        xlApp.DisplayAlerts = False
        xlWorkBook.SaveAs("C:\Users\Sean\Documents\Visual Studio 2017\Projects\Avfuel System\Avfuel System\Resources\Avfuel_Database.xlsx")

        xlWorkBook.Close()
        xlApp.Quit()

        releaseObject(xlApp)
        releaseObject(xlWorkBook)
        releaseObject(xlWorkSheet)

        Dim t1 As Integer
        Dim t2 As Integer
        t1 = 2000
        t2 = 2000
        'Text Box Layout
        '1
        '4

        '6
        '5

        '3
        '7
        xlApp = New Excel.Application
        xlWorkBook = xlApp.Workbooks.Open("C:\Users\Sean\Documents\Visual Studio 2017\Projects\Avfuel System\Avfuel System\Resources\Avfuel_Database.xlsx")
        xlWorkSheet = xlWorkBook.Worksheets("sheet1")
        xlApp.DisplayAlerts = False

        dispatchTerminal.Label2.ForeColor = Color.Black
        dispatchTerminal.Label3.ForeColor = Color.Black
        dispatchTerminal.Label4.ForeColor = Color.Black
        dispatchTerminal.TextBox1.Text = ""
        dispatchTerminal.TextBox4.Text = ""
        dispatchTerminal.TextBox6.Text = ""
        dispatchTerminal.TextBox5.Text = ""
        dispatchTerminal.TextBox3.Text = ""
        dispatchTerminal.TextBox7.Text = ""

        Row = xlWorkSheet.UsedRange.Rows.Count
        For i = 1 To Row
            If xlWorkSheet.Cells(i + 1, 11).Value = "No" Then
                dispatchTerminal.Label2.ForeColor = Color.DarkGray
                dispatchTerminal.TextBox1.Text = xlWorkSheet.Cells(i + 1, 1).Value & " - " & xlWorkSheet.Cells(i + 1, 4).Value & " | " & xlWorkSheet.Cells(i + 1, 6).Value & " Gallons of " & xlWorkSheet.Cells(i + 1, 5).Value
                dispatchTerminal.TextBox4.Text = "Customer: " & xlWorkSheet.Cells(i + 1, 2).Value & " " & xlWorkSheet.Cells(i + 1, 3).Value & " | " & xlWorkSheet.Cells(i + 1, 12).Value
                dispatchTerminal.TextBox4.ForeColor = Color.DarkRed
                t1 = (i + 1)
                Exit For
            End If
        Next

        For i = t1 To Row
            If xlWorkSheet.Cells(i + 1, 11).Value = "No" Then
                dispatchTerminal.Label3.ForeColor = Color.DarkGray
                dispatchTerminal.TextBox6.Text = xlWorkSheet.Cells(i + 1, 1).Value & " - " & xlWorkSheet.Cells(i + 1, 4).Value & " | " & xlWorkSheet.Cells(i + 1, 6).Value & " Gallons of " & xlWorkSheet.Cells(i + 1, 5).Value
                dispatchTerminal.TextBox5.Text = "Customer: " & xlWorkSheet.Cells(i + 1, 2).Value & " " & xlWorkSheet.Cells(i + 1, 3).Value & " | " & xlWorkSheet.Cells(i + 1, 12).Value
                dispatchTerminal.TextBox5.ForeColor = Color.DarkRed
                t2 = (i + 1)
                Exit For
            End If
        Next

        For i = t2 To Row
            If xlWorkSheet.Cells(i + 1, 11).Value = "No" Then
                dispatchTerminal.Label4.ForeColor = Color.DarkGray
                dispatchTerminal.TextBox3.Text = xlWorkSheet.Cells(i + 1, 1).Value & " - " & xlWorkSheet.Cells(i + 1, 4).Value & " | " & xlWorkSheet.Cells(i + 1, 6).Value & " Gallons of " & xlWorkSheet.Cells(i + 1, 5).Value
                dispatchTerminal.TextBox7.Text = "Customer: " & xlWorkSheet.Cells(i + 1, 2).Value & " " & xlWorkSheet.Cells(i + 1, 3).Value & " | " & xlWorkSheet.Cells(i + 1, 12).Value
                dispatchTerminal.TextBox7.ForeColor = Color.DarkRed
                Exit For
            End If
        Next

        xlWorkBook.Close()
        xlApp.Quit()

        releaseObject(xlApp)
        releaseObject(xlWorkBook)
        releaseObject(xlWorkSheet)

        Me.Hide()

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        Me.Hide()

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

End Class