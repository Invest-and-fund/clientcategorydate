Imports System.Globalization
Imports System.Configuration
Imports System.Collections.Specialized
Imports System.IO
Imports System.Data
Imports System.Data.SqlClient

Public Class Form1


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Dim i, j As Integer
        Dim dr As DataRow
        Dim MyConn As FirebirdSql.Data.FirebirdClient.FbConnection
        Dim strConn As String
        Dim ds As New DataSet

        Dim sSQL As String

        Dim Adaptor As FirebirdSql.Data.FirebirdClient.FbDataAdapter
        sSQL = "select n.userid, min(n.history_datecreated) as ccd
                from new_users_history n
                where activated_cert = 5
                group by n.userid "

        strConn = System.Configuration.ConfigurationManager.ConnectionStrings("FBConnectionString").ConnectionString
        MyConn = New FirebirdSql.Data.FirebirdClient.FbConnection(strConn)
        MyConn.Open()
        Adaptor = New FirebirdSql.Data.FirebirdClient.FbDataAdapter(sSQL, MyConn)

        Adaptor.Fill(ds)
        MyConn.Close()

        j = ds.Tables(0).Rows.Count

        For i = 0 To j - 1
            dr = ds.Tables(0).Rows(i)
            Dim ccdate As Date = dr("ccd")
            Dim userid As Integer = dr("userid")
            Dim Cmd As FirebirdSql.Data.FirebirdClient.FbCommand
            sSQL = "update users
                    set client_categorisation_date = @ccdate
                    where userid = @userID"

            strConn = System.Configuration.ConfigurationManager.ConnectionStrings("FBConnectionString").ConnectionString
            MyConn = New FirebirdSql.Data.FirebirdClient.FbConnection(strConn)

            MyConn.Open()
            Cmd = New FirebirdSql.Data.FirebirdClient.FbCommand
            Cmd.Connection = MyConn
            Cmd.CommandText = sSQL
            Cmd.Parameters.Clear()
            Cmd.Parameters.Add("userID", FirebirdSql.Data.FirebirdClient.FbDbType.Integer).Value = userid
            Cmd.Parameters.Add("ccDate", FirebirdSql.Data.FirebirdClient.FbDbType.Date).Value = ccdate
            Cmd.ExecuteNonQuery()

            MyConn.Close()

        Next

        Message.Text = "update complete"

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Message.Text = "click to start"
    End Sub
End Class
