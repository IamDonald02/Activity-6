Imports System.Data.OleDb
Imports System.IO
Imports MySql.Data.MySqlClient
Imports Excel = Microsoft.Office.Interop.Excel
Public Class Form4product

    Private dtDisplay As New DataTable()

    Private Sub Import_Click(sender As Object, e As EventArgs) Handles Import.Click

        'Open file dialog to select CSV file
        Dim openFileDialog As New OpenFileDialog()
        openFileDialog.Filter = "CSV Files (*.csv)|*.csv|All files (*.*)|*.*"
        openFileDialog.FilterIndex = 1
        openFileDialog.RestoreDirectory = True

        If openFileDialog.ShowDialog() = DialogResult.OK Then

            Dim filePath As String = openFileDialog.FileName

            'Create connection string for OleDB driver to read CSV file
            Dim connString As String = String.Format("Provider=Microsoft.Jet.OleDb.4.0;Data Source={0};Extended Properties=""Text;HDR=YES;FMT=Delimited""", Path.GetDirectoryName(filePath))

            'Create connection to CSV file and read its contents into a DataTable
            Dim dt As New DataTable()
            Using conn As New OleDbConnection(connString)
                conn.Open()
                Dim cmd As New OleDbCommand(String.Format("SELECT * FROM [{0}]", Path.GetFileName(filePath)), conn)
                Dim adapter As New OleDbDataAdapter(cmd)
                adapter.Fill(dt)
            End Using

            'Create connection string for MySQL database
            Dim myConnectionString As String = "server=127.0.0.1;uid=kopi;pwd=kopi;database=coffee_shop"

            'Create connection to MySQL database and insert DataTable contents into a table
            Using conn As New MySqlConnection(myConnectionString)
                conn.Open()
                Dim cmd As New MySqlCommand()
                cmd.Connection = conn
                For Each row As DataRow In dt.Rows
                    Try
                        cmd.CommandText = String.Format("INSERT INTO products (product_id, product_name, product_category, product_price, product_quantity) VALUES ('{0}', '{1}', '{2}', '{3}', '{4}')", row("Id"), row("Product"), row("Category"), row("Price"), row("Quantity"))
                        cmd.ExecuteNonQuery()
                    Catch ex As MySqlException
                        If ex.Number = 1062 Then '1062 is the error number for a duplicate key violation
                            MessageBox.Show(String.Format("Product with ID '{0}' already exists in the database.", row("product_id")), "Duplicate Product", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                        Else
                            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        End If
                    End Try
                Next
            End Using

            'Create connection to MySQL database and retrieve the inserted data into a DataTable for display
            Dim dtDisplay As New DataTable()
            Using conn As New MySqlConnection(myConnectionString)
                conn.Open()
                Dim cmd As New MySqlCommand("SELECT * FROM products", conn)
                Dim adapter As New MySqlDataAdapter(cmd)
                adapter.Fill(dtDisplay)
            End Using

            ' Retrieve the updated data from the database and display it in the DataGridView
            RefreshDataGrid()

        End If
    End Sub

    Private Sub RefreshDataGrid()
        ' Create connection to MySQL database and retrieve the data into a DataTable for display
        Using conn As New MySqlConnection("server=127.0.0.1;uid=kopi;pwd=kopi;database=coffee_shop")
            conn.Open()
            Dim cmd As New MySqlCommand("SELECT * FROM products", conn)
            Dim adapter As New MySqlDataAdapter(cmd)
            dtDisplay.Clear()
            adapter.Fill(dtDisplay)
        End Using

        ' Display the data in the DataGridView
        DataGridView1.DataSource = dtDisplay
    End Sub

    Private Sub Form4product_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        RefreshDataGrid()
    End Sub

    Private Sub Guna2Button7_Click(sender As Object, e As EventArgs) Handles Export.Click
        ExportToExcel()
    End Sub

    Private Sub Backup_Click(sender As Object, e As EventArgs) Handles Backup.Click
        'Create a new instance of the SaveFileDialog class
        Dim backup As New SaveFileDialog()
        'Locate in local drive
        backup.InitialDirectory = "C:\"
        backup.Title = "Database Backup"
        backup.DefaultExt = "sql"
        'Set file format
        backup.Filter = "sql files (.sql)|.sql|All files (.)|*.*"
        backup.RestoreDirectory = True

        If backup.ShowDialog() = DialogResult.OK Then
            'Connection to DB
            Using conn As New MySqlConnection("server=127.0.0.1;uid=kopi;pwd=kopi;database=coffee_shop")
                Using cmd As New MySqlCommand()
                    conn.Open()
                    cmd.Connection = conn
                    Dim mybu As New MySqlBackup(cmd)
                    mybu.ExportToFile(backup.FileName)
                    'Mesagebox
                    MessageBox.Show("Database Backup Successful!")
                End Using
            End Using
        End If
    End Sub

    Private Sub Add_Button_Click(sender As Object, e As EventArgs) Handles Add_Button.Click
        ' Call to connect to the database
        Connect_to_DB()

        Dim prodName As String = prodName_TextBox.Text
        Dim category As String = category_TextBox.Text
        Dim price As String = price_TextBox.Text
        Dim quantity As String = quantity_TextBox.Text

        ' Check if any of the input fields are empty
        If String.IsNullOrWhiteSpace(prodName) OrElse String.IsNullOrWhiteSpace(category) OrElse String.IsNullOrWhiteSpace(price) OrElse String.IsNullOrWhiteSpace(quantity) Then
            MessageBox.Show("Please input data in all fields before clicking the Add button.")
        Else
            Dim strSQL As String = "INSERT INTO products (product_name, product_category, product_price, product_quantity) VALUES (@prodName, @category, @price, @quantity)"
            Using conn As New MySqlConnection("server=127.0.0.1;uid=kopi;pwd=kopi;database=coffee_shop")
                conn.Open()
                Using cmd As New MySqlCommand(strSQL, conn)
                    cmd.Parameters.AddWithValue("@prodName", prodName)
                    cmd.Parameters.AddWithValue("@category", category)
                    cmd.Parameters.AddWithValue("@price", price)
                    cmd.Parameters.AddWithValue("@quantity", quantity)
                    cmd.ExecuteNonQuery()
                End Using
            End Using

            MessageBox.Show("Record Added")

            ' Clear the text boxes
            prodName_TextBox.Clear()
            category_TextBox.Clear()
            price_TextBox.Clear()
            quantity_TextBox.Clear()

            ' Refresh the DataGridView to display the updated data
            RefreshDataGrid()
        End If

        Disconnect_to_DB()
    End Sub

    Private Sub Delete_Click(sender As Object, e As EventArgs) Handles Delete.Click
        ' Prompt the user to enter the product ID to be deleted
        Dim productId As String = InputBox("Enter the Product ID to delete:", "Delete Record")

        ' Check if the user entered a value
        If Not String.IsNullOrEmpty(productId) Then
            ' Connect to the database
            Connect_to_DB()

            ' Delete the record from the database
            Dim strSQL As String = "DELETE FROM products WHERE product_id = @productId"
            Using conn As New MySqlConnection("server=127.0.0.1;uid=kopi;pwd=kopi;database=coffee_shop")
                conn.Open()
                Using cmd As New MySqlCommand(strSQL, conn)
                    cmd.Parameters.AddWithValue("@productId", productId)
                    Dim rowsAffected As Integer = cmd.ExecuteNonQuery()

                    ' Check if any rows were affected
                    If rowsAffected > 0 Then
                        ' Display success message
                        MessageBox.Show("Record Deleted")

                        ' Refresh the DataGridView to reflect the changes
                        RefreshDataGrid()
                    Else
                        MessageBox.Show("No record found with the provided Product ID.", "Record Not Found", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    End If
                End Using
            End Using

            ' Disconnect from the database
            Disconnect_to_DB()
        End If
    End Sub

    Private Sub search_TextBox_TextChanged(sender As Object, e As EventArgs) Handles search_TextBox.TextChanged
        ' Clear the DataGridView when the search textbox is empty
        If String.IsNullOrEmpty(search_TextBox.Text) Then
            ' Display all data in the DataGridView
            RefreshDataGrid()
        End If
    End Sub

    Private Sub Search_Click(sender As Object, e As EventArgs) Handles Search.Click
        ' Get the search ID from the search textbox
        Dim searchId As String = search_TextBox.Text

        ' Check if the search ID is provided
        If Not String.IsNullOrEmpty(searchId) Then
            ' Connect to the database
            Connect_to_DB()

            ' Search for the record in the database
            Dim strSQL As String = "SELECT * FROM products WHERE product_id = @id"
            Using conn As New MySqlConnection("server=127.0.0.1;uid=kopi;pwd=kopi;database=coffee_shop")
                conn.Open()
                Using cmd As New MySqlCommand(strSQL, conn)
                    cmd.Parameters.AddWithValue("@id", searchId)

                    ' Create a DataTable to hold the search result
                    Dim dt As New DataTable()
                    Dim adapter As New MySqlDataAdapter(cmd)
                    adapter.Fill(dt)

                    ' Display the search result in the DataGridView
                    DataGridView1.DataSource = dt
                End Using
            End Using

            ' Disconnect from the database
            Disconnect_to_DB()
        End If
    End Sub

    Private Sub Edit_Click(sender As Object, e As EventArgs) Handles Edit.Click
        ' Prompt the user to enter the details of the row to be edited
        Dim id As String = InputBox("Enter the Product ID:", "Edit Record")
        If String.IsNullOrEmpty(id) Then
            ' Exit the function if the user clicked "Cancel" or closed the input box
            Return
        End If

        Dim product As String = InputBox("Enter the Product:", "Edit Record")
        If String.IsNullOrEmpty(product) Then
            ' Exit the function if the user clicked "Cancel" or closed the input box
            Return
        End If

        Dim category As String = InputBox("Enter the Category:", "Edit Record")
        If String.IsNullOrEmpty(category) Then
            ' Exit the function if the user clicked "Cancel" or closed the input box
            Return
        End If

        Dim price As String = InputBox("Enter the Price:", "Edit Record")
        If String.IsNullOrEmpty(price) Then
            ' Exit the function if the user clicked "Cancel" or closed the input box
            Return
        End If

        Dim quantity As String = InputBox("Enter the Quantity:", "Edit Record")
        If String.IsNullOrEmpty(quantity) Then
            ' Exit the function if the user clicked "Cancel" or closed the input box
            Return
        End If

        ' Connect to the database
        Connect_to_DB()

        ' Update the record in the database
        Dim strSQL As String = "UPDATE products SET product_name = @product, product_category = @category, product_price = @price, product_quantity = @quantity WHERE product_id = @id"
        Using conn As New MySqlConnection("server=127.0.0.1;uid=kopi;pwd=kopi;database=coffee_shop")
            conn.Open()
            Using cmd As New MySqlCommand(strSQL, conn)
                cmd.Parameters.AddWithValue("@product", product)
                cmd.Parameters.AddWithValue("@category", category)
                cmd.Parameters.AddWithValue("@price", price)
                cmd.Parameters.AddWithValue("@quantity", quantity)
                cmd.Parameters.AddWithValue("@id", id)
                Dim rowsAffected As Integer = cmd.ExecuteNonQuery()

                ' Check if any rows were affected
                If rowsAffected > 0 Then
                    ' Display success message
                    MessageBox.Show("Record Updated")

                    ' Refresh the DataGridView to reflect the changes
                    RefreshDataGrid()
                Else
                    MessageBox.Show("No record found with the provided Product ID.", "Record Not Found", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
            End Using
        End Using

        ' Disconnect from the database
        Disconnect_to_DB()
    End Sub

    Private Sub Guna2TextBox5_TextChanged(sender As Object, e As EventArgs)

    End Sub
    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub Guna2TextBox3_TextChanged(sender As Object, e As EventArgs) Handles prodName_TextBox.TextChanged

    End Sub

    Private Sub Guna2TextBox2_TextChanged(sender As Object, e As EventArgs) Handles category_TextBox.TextChanged

    End Sub

    Private Sub Guna2TextBox4_TextChanged(sender As Object, e As EventArgs) Handles price_TextBox.TextChanged

    End Sub

    Private Sub Guna2TextBox5_TextChanged_1(sender As Object, e As EventArgs) Handles quantity_TextBox.TextChanged

    End Sub

    Private Sub Guna2TextBox2_TextChanged_1(sender As Object, e As EventArgs)

    End Sub

    Private Sub Guna2Button1_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Guna2Button1_Click_1(sender As Object, e As EventArgs) Handles Back.Click
        Me.Hide()
        Form3dashbord.Show()
    End Sub
End Class