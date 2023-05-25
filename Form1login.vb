Imports MySql.Data.MySqlClient
Public Class Form1login
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Guna2PictureBox1_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Guna2PictureBox1_Click_1(sender As Object, e As EventArgs)

    End Sub

    Private Sub Guna2TextBox1_TextChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub Guna2TextBox1_TextChanged_1(sender As Object, e As EventArgs) Handles username.TextChanged, password.TextChanged

    End Sub

    Private Sub Guna2HtmlLabel1_Click(sender As Object, e As EventArgs) Handles helloLabel.Click, rememLabel.Click, signinLabel.Click

    End Sub

    Private Sub Guna2GradientButton1_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Guna2RadioButton1_CheckedChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub Guna2ControlBox1_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub LoginPanel_Paint(sender As Object, e As PaintEventArgs)

    End Sub

    Private Sub registerButton_Click(sender As Object, e As EventArgs) Handles registerButton.Click
        Me.Hide()
        Form2Register.Show()
    End Sub

    Private Sub Guna2ShadowPanel1_Paint(sender As Object, e As PaintEventArgs)

    End Sub

    Private Sub rememberMeToggleSwitch_CheckedChanged(sender As Object, e As EventArgs) Handles rememberMeToggleSwitch.CheckedChanged

    End Sub

    Private Sub Guna2GradientButton1_Click_1(sender As Object, e As EventArgs) Handles Guna2GradientButton1.Click
        With Me
            'Call to connect in the db kopi
            Call Connect_to_DB()
            Dim mycmd As New MySqlCommand
            Dim myreader As MySqlDataReader
            strSQL = "Select * from accounts where acct_username = '" _
                & username.Text & "' and acct_password = md5('" _
                & password.Text & "')"
            'MsgBox(strSQL)
            mycmd.CommandText = strSQL
            mycmd.Connection = myconn

            myreader = mycmd.ExecuteReader
            If myreader.HasRows Then
                .Hide()
                Form3dashbord.Show()
            Else
                MessageBox.Show("Invalid username or password")
            End If
            Call Disconnect_to_DB()
        End With

    End Sub

    Private Sub bgPictureBox_Click(sender As Object, e As EventArgs) Handles bgPictureBox.Click

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub UserManualButton_Click(sender As Object, e As EventArgs) Handles UserManualButton.Click
        Try
            Dim userManualFilePath As String = "C:\Users\user\Desktop\UserManual.pdf" ' Replace with the actual file path of the user manual

            ' Open the user manual file using the default application associated with PDF files
            System.Diagnostics.Process.Start(userManualFilePath)
        Catch ex As Exception
            MessageBox.Show("Error opening user manual: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

End Class
