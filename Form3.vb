Imports System.Data.SqlClient

Public Class Form3
    Dim objForm1 As New Form1
    Dim dr As SqlDataReader
    Dim arrays(500) As String
    Private Sub Form3_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        objForm1.getConn().Open()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        DataGridView1.ColumnCount = 4
        DataGridView1.Columns(0).Name = "Number"
        DataGridView1.Columns(1).Name = "char"
        DataGridView1.Columns(2).Name = "conuter"
        DataGridView1.Columns(3).Name = "AVG"
        Dim i As Integer = 0
        Dim conuter As Integer = 0
        Dim sql As String = "select sum(conu),char,avg(avg) from MoreCharacters group by char order by sum(conu) DESC  "
        Dim cmd As New SqlCommand(sql, objForm1.conn)
        cmd = New System.Data.SqlClient.SqlCommand(sql, objForm1.conn)
        dr = cmd.ExecuteReader
        If dr.HasRows = True Then
            Do While (dr.Read)
                DataGridView1.RowCount += 1
                DataGridView1.Item(0, i).Value = i
                DataGridView1.Item(1, i).Value = dr.Item("char")
                DataGridView1.Item(2, i).Value = dr(0) ' dr(1)
                DataGridView1.Item(3, i).Value = dr(2) ' dr(1)
                i += 1

            Loop
        End If
        dr.Close()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Try ' crate table 
            Dim sql As String = "CREATE TABLE MAXMoreCharacters (number int ,char VARCHAR(30) ,conuter int ,avg int, ) "
            Dim cmd As New SqlCommand(sql, objForm1.conn)
            cmd.Connection = objForm1.conn
            cmd.ExecuteNonQuery()
            '  MsgBox("ADD")
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        ' MsgBox(fiter_user)
        Dim i As Integer ' add data form DataGridView2 to BD
        For i = 0 To DataGridView1.Rows.Count - 2
            Dim row As DataGridViewRow = DataGridView1.Rows(i)
            Dim number As DataGridViewTextBoxCell = row.Cells(0)
            Dim charr As DataGridViewTextBoxCell = row.Cells(1)
            Dim conuter As DataGridViewTextBoxCell = row.Cells(2)
            Dim avg As DataGridViewTextBoxCell = row.Cells(3)
            Dim sql As String = "insert into MAXMoreCharacters  (number,char,conuter,avg)values('" & number.Value & "','" & charr.Value & "','" & conuter.Value & "','" & avg.Value & "')" ' " & fiter_user & "
            Dim cmd As New SqlCommand(sql, objForm1.conn)
            cmd.CommandText = sql
            cmd.Connection = objForm1.conn
            cmd.ExecuteNonQuery()
        Next
    End Sub
End Class