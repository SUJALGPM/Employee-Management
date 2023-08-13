Imports System.Data.OleDb
Imports System.Data
Public Class Form1
    Sub view()
        Dim con As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\SUJAL\Documents\Database2.mdb")
        con.Open()
        Dim adp As New OleDbDataAdapter("Select * from department", con)
        Dim ds As New DataSet
        adp.Fill(ds, "department")
        DataGridView1.DataSource = ds.Tables("department")
        con.Close()
    End Sub
    Public Sub ShowData(ByVal CurrentRow)
        Try
            TextBox1.Text = ds.Tables("department").Rows(CurrentRow)("id")
            TextBox2.Text = ds.Tables("department").Rows(CurrentRow)("name")
            TextBox3.Text = ds.Tables("department").Rows(CurrentRow)("course")
            TextBox4.Text = ds.Tables("department").Rows(CurrentRow)("marks")
            TextBox5.Text = ds.Tables("department").Rows(CurrentRow)("place")
        Catch ex As Exception
            MessageBox.Show(ex.Message, "error")
        End Try
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim con As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\SUJAL\Documents\Database2.mdb")
        con.Open()
        Dim cmd As New OleDbCommand()
        cmd.Connection = con
        cmd.CommandType = CommandType.Text
        cmd.CommandText = "Insert into department values(" & TextBox1.Text & ",'" & TextBox2.Text & "','" & TextBox3.Text & "'," & TextBox4.Text & ",'" & TextBox5.Text & "')"
        cmd.ExecuteNonQuery()
        con.Close()
        MessageBox.Show("RECORD SAVE")
        view()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        view()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim con As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\SUJAL\Documents\Database2.mdb")
        con.Open()
        Dim cmd As New OleDbCommand()
        cmd.Connection = con
        cmd.CommandType = CommandType.Text
        cmd.CommandText = "Update department set name='" & TextBox2.Text & "',course='" & TextBox3.Text & "',marks=" & TextBox4.Text & ",place='" & TextBox5.Text & "' where id=" & TextBox1.Text & ""
        cmd.ExecuteNonQuery()
        con.Close()
        MessageBox.Show("RECORD IS UPDATED")
        view()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim con As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\SUJAL\Documents\Database2.mdb")
        con.Open()
        Dim cmd As New OleDbCommand()
        cmd.Connection = con
        cmd.CommandType = CommandType.Text
        cmd.CommandText = "Delete from department where id=" & TextBox1.Text & ""
        cmd.ExecuteNonQuery()
        con.Close()
        MessageBox.Show("RECORD IS DELETED")
        view()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        If CurrentRow <> 0 Then
            CurrentRow -= 1
            ShowData(CurrentRow)
        Else
            MsgBox("1ST Row Reached", MsgBoxStyle.Exclamation)
        End If
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        If CurrentRow = ds.Tables("department").Rows.Count - 1 Then
            MsgBox("Last Row Reached", MsgBoxStyle.Exclamation)
        Else
            CurrentRow += 1
            ShowData(CurrentRow)
        End If
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim con As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\SUJAL\Documents\Database2.mdb")
        CurrentRow = 0
        con.Open()
        adp = New OleDbDataAdapter("Select * from department", con)
        adp.Fill(ds, "department")
        ShowData(CurrentRow)
        con.Close()
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        TextBox1.Clear()
        TextBox2.Clear()
        TextBox3.Clear()
        TextBox4.Clear()
        TextBox5.Clear()
    End Sub
End Class
