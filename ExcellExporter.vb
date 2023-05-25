Imports MySql.Data.MySqlClient
Imports Excel = Microsoft.Office.Interop.Excel

Module ExcelExportModule

    Public Sub ExportToExcel()
        Dim myConnectionString As String = "server=127.0.0.1;uid=kopi;pwd=kopi;database=coffee_shop"
        Dim strSQL As String = "SELECT * FROM products"
        Dim xlsFiles As String = System.IO.Directory.GetCurrentDirectory & "\..\..\dataXls\"
        Dim templatefilename As String = "C:\Users\user\Documents\Bicol University\3rd Year College\2nd Semester\IT 120 - Event Driven Programming\Main\Template.xlsx" ' Update with the complete file path of your template banner

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
            xlsWB = xlsApp.Workbooks.Open(templatefilename)
            xlsSheet = xlsWB.Worksheets(1)

            ' Write the data to the worksheet starting from cell A9
            Dim x, y As Integer
            For x = 0 To mydt.Rows.Count - 1
                For y = 0 To mydt.Columns.Count - 1
                    xlsSheet.Cells(x + 9, y + 1) = mydt.Rows(x)(y)
                Next
            Next

            ' Apply formatting if needed
            ' With xlsSheet.Range("A9", ConvertToLetters(mydt.Columns.Count) & (mydt.Rows.Count + 8))
            '     .Borders.LineStyle = Excel.XlLineStyle.xlContinuous
            ' End With

            ' Save the workbook and prompt the user for the file name and location
            Dim saveFileDialog As New SaveFileDialog()
            saveFileDialog.Filter = "Excel Workbook (*.xlsx)|*.xlsx"
            saveFileDialog.Title = "Save Excel Workbook"
            saveFileDialog.FileName = "products"
            saveFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
            If saveFileDialog.ShowDialog() = DialogResult.OK Then
                ' Save the workbook
                xlsWB.SaveAs(saveFileDialog.FileName)

                ' Close the workbook and quit Excel
                xlsWB.Close()
                xlsApp.Quit()

                ' Release the COM objects
                releaseObject(xlsSheet)
                releaseObject(xlsWB)
                releaseObject(xlsApp)

                ' Show a message box indicating successful export
                MessageBox.Show("Data exported successfully to: " & saveFileDialog.FileName, "Export Successful", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If

        Catch ex As Exception
            MessageBox.Show("Error exporting data to Excel: " & ex.Message, "Export Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Public Function ConvertToLetters(ByVal number As Integer) As String
        number -= 1
        Dim result As String = String.Empty

        If 26 > number Then
            result = Chr(number + 65)
        Else
            Dim column As Integer

            Do
                column = number Mod 26
                number = (number \ 26) - 1
                result = Chr(column + 65) & result
            Loop Until number < 0
        End If

        Return result
    End Function

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