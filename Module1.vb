Imports MySql.Data.MySqlClient
Imports Excel = Microsoft.Office.Interop.Excel
Module Module1
    Public myconn As New MySql.Data.MySqlClient.MySqlConnection
    Public myConnectionString As String
    Public strSQL As String

    Public Sub Connect_to_DB()
        myConnectionString = "server=127.0.0.1;" _
                    & "uid=kopi;" _
                    & "pwd=kopi;" _
                    & "database=coffee_shop"
        Try
            myconn.ConnectionString = myConnectionString
            myconn.Open()

        Catch ex As MySql.Data.MySqlClient.MySqlException
            Select Case ex.Number
                Case 1
                    MsgBox("Cannot Connect to Server")
                Case 1067
                    MsgBox("Invalid Username or password!")
            End Select

        End Try
    End Sub

    Public Sub Disconnect_to_DB()
        myconn.Close()
        myconn.Dispose()
    End Sub

End Module