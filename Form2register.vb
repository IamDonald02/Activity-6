Imports MySql.Data.MySqlClient
Public Class Form2Register
    Public DBConnection As New MySqlConnection
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Guna2PictureBox1_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Guna2PictureBox1_Click_1(sender As Object, e As EventArgs)

    End Sub

    Private Sub Guna2TextBox1_TextChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub Guna2TextBox1_TextChanged_1(sender As Object, e As EventArgs) Handles firstname.TextChanged, lastname.TextChanged, age.TextChanged, sex.TextChanged

    End Sub

    Private Sub Guna2HtmlLabel1_Click(sender As Object, e As EventArgs) Handles helloLabel.Click, signinLabel.Click, Guna2HtmlLabel1.Click

    End Sub

    Private Sub Guna2GradientButton1_Click(sender As Object, e As EventArgs) Handles loginGradientButton.Click

        With Me
            Call Connect_to_DB()
            Dim mycmd As New MySqlCommand
            Try
                If confirm_password.Text = confirm_password.Text Then
                    strSQL = "INSERT INTO accounts (acct_id, acct_firstname, acct_lastname, acct_age, acct_sex, acct_username, acct_password) VALUES ('" _
                    & username.Text & "', '" _
                    & firstname.Text & "', '" _
                    & lastname.Text & "', '" _
                    & age.Text & "', '" _
                    & sex.Text & "', '" _
                    & username.Text & "', md5('" _
                    & confirm_password.Text & "'))"

                    mycmd.CommandText = strSQL
                    mycmd.Connection = myconn
                    mycmd.ExecuteNonQuery()
                    MsgBox("Record Successfully Added")
                End If

            Catch ex As MySqlException
                MsgBox(ex.Number & " " & ex.Message)

            End Try
            Disconnect_to_DB()
        End With

    End Sub

    Private Sub Guna2RadioButton1_CheckedChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub Guna2ControlBox1_Click(sender As Object, e As EventArgs) Handles Guna2ControlBox1.Click

    End Sub

    Private Sub LoginPanel_Paint(sender As Object, e As PaintEventArgs) Handles LoginPanel.Paint

    End Sub

    Private Sub registerButton_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Guna2ShadowPanel1_Paint(sender As Object, e As PaintEventArgs)

    End Sub

    Private Sub Guna2ImageButton1_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub registerButton_Click_1(sender As Object, e As EventArgs) Handles registerButton.Click
        Me.Hide()
        Form1login.Show()
    End Sub

    Private Sub Guna2TextBox2_TextChanged(sender As Object, e As EventArgs) Handles username.TextChanged

    End Sub

    Private Sub bgPictureBox_Click(sender As Object, e As EventArgs) Handles bgPictureBox2.Click

    End Sub

    Private Sub con_pass_TextChanged(sender As Object, e As EventArgs) Handles confirm_pass.TextChanged

    End Sub

    Private Sub enter_password_TextChanged(sender As Object, e As EventArgs) Handles confirm_password.TextChanged

    End Sub
End Class
