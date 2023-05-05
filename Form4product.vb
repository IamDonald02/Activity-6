Imports System.Data.OleDb
Imports System.IO
Imports MySql.Data.MySqlClient
Imports Excel = Microsoft.Office.Interop.Excel
Public Class Form4product


    Private Sub Guna2Button1_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Guna2Button1_Click_1(sender As Object, e As EventArgs) Handles Guna2Button1.Click, Guna2Button4.Click
        Me.Hide()
        Form3dashbord.Show()
    End Sub

    Private Sub Guna2Button2_Click(sender As Object, e As EventArgs) Handles Guna2Button2.Click

    End Sub

    Private Sub Guna2TextBox5_TextChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub Guna2TextBox2_TextChanged(sender As Object, e As EventArgs) Handles Guna2TextBox2.TextChanged

    End Sub

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

            'Display the imported data in a DataGridView
            DataGridView1.DataSource = dtDisplay

        End If
    End Sub

    Private Sub Guna2Button7_Click(sender As Object, e As EventArgs) Handles Export.Click
        importToExcel()
    End Sub

    Private Sub Backup_Click(sender As Object, e As EventArgs) Handles Backup.Click
        'Create a new instance of the SaveFileDialog class
        Dim backup As New SaveFileDialog
        'Locate in local drive
        backup.InitialDirectory = "C:\"
        backup.Title = "Database Backup"
        backup.DefaultExt = "sql"
        'Set file format
        backup.Filter = "sql files (.sql)|.sql|All files (.)|*.*"
        backup.RestoreDirectory = True

        If backup.ShowDialog = Windows.Forms.DialogResult.OK Then
            'Connection to DB
            Call Connect_to_DB()
            Dim mycmd As MySqlCommand = New MySqlCommand
            mycmd.Connection = myconn
            Dim mybu As MySqlBackup = New MySqlBackup(mycmd)
            mybu.ExportToFile(backup.FileName)
            myconn.Close()
            'Mesagebox
            MessageBox.Show("Database Back-up Successful!")
        ElseIf backup.ShowDialog = Windows.Forms.DialogResult.Cancel Then
            Return
        End If
    End Sub
End Class