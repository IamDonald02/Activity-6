Imports MySql.Data.MySqlClient
Imports Excel = Microsoft.Office.Interop.Excel

Module ExcelExportModule

    Public Sub importToExcel()
        Dim myConnectionString As String = "server=127.0.0.1;uid=kopi;pwd=kopi;database=coffee_shop"
        Dim strSQL As String = "SELECT * FROM products"
        Dim xlsPath As String = System.IO.Directory.GetCurrentDirectory & "\..\..\dataXls\TEMPLATE\"
        Dim xlsFiles As String = System.IO.Directory.GetCurrentDirectory & "\..\..\dataXls\"

        Dim xlsApp As Excel.Application
        Dim xlsWB As Excel.Workbook
        Dim xlsSheet As Excel.Worksheet

        Try
            ' Connect to the database
            Dim myconn As New MySqlConnection(myConnectionString)
            myconn.Open()

            ' Execute the SQL query
            Dim mycmd As New MySqlCommand(strSQL, myconn)
            Dim myda As New MySqlDataAdapter(mycmd)
            Dim mydt As New DataTable
            myda.Fill(mydt)

            ' Create a new Excel workbook and worksheet
            xlsApp = New Excel.Application
            xlsApp.Visible = False
            xlsWB = xlsApp.Workbooks.Add()
            xlsSheet = xlsWB.ActiveSheet

            ' Write the data to the worksheet
            Dim x, y As Integer
            For x = 0 To mydt.Rows.Count - 1
                For y = 0 To mydt.Columns.Count - 1
                    xlsSheet.Cells(x + 1, y + 1) = mydt.Rows(x)(y)
                Next
            Next

            ' Prompt the user for the file name and location
            Dim saveFileDialog As New SaveFileDialog()
            saveFileDialog.Filter = "Excel Workbook (*.xlsx)|*.xlsx"
            saveFileDialog.Title = "Save Excel Workbook"
            saveFileDialog.FileName = "products"
            saveFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
            If saveFileDialog.ShowDialog() = DialogResult.OK Then
                ' Save the workbook and quit Excel
                xlsWB.SaveAs(saveFileDialog.FileName)
                xlsWB.Close()
                xlsApp.Quit()

                ' Release the COM objects and show a message box
                releaseObject(xlsSheet)
                releaseObject(xlsWB)
                releaseObject(xlsApp)

                MessageBox.Show("Data exported successfully to: " & saveFileDialog.FileName, "Export Successful", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If

        Catch ex As Exception
            MessageBox.Show("Error exporting data to Excel: " & ex.Message, "Export Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
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

End Module