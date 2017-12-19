Imports System
Imports System.Data.SqlClient
Imports System.IO

Public Class Group1
    Dim dr As SqlDataReader
    Dim cmd As New SqlCommand
    Dim nameUser(21) As String
    Dim group(21) As Integer
    Dim Account As Integer
    Dim clo As Integer = 0
    Public conn As New SqlConnection("Data Source=.\SQLEXPRESS;AttachDbFilename=" & Application.StartupPath & "\output.mdf;Integrated Security=True;Connect Timeout=30;User Instance=True")

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try 'Connection DB
            conn.Open()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        dd()
    End Sub

    Sub dd()
        Dim i As Integer = 0
        Try
            Dim sql As String = "select nameUser from AllUser  "
            cmd = New System.Data.SqlClient.SqlCommand(sql, conn)
            dr = cmd.ExecuteReader
            If dr.HasRows = True Then
                Do While (dr.Read)
                    nameUser(i) = dr(0)
                    i += 1
                Loop
            End If
            dr.Close()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        'For i = 1 To 2
        '    MsgBox(" index = " & i & " name = " & nameUser(i))
        'Next
    End Sub
    Sub xx(ByVal a As String, ByVal loaction As Integer)
        CreateDataGridView1()
        Dim i As Integer = 0
        Try
            Dim sql As String = "select * from char130 where Observation like '" & a & "' and class =  " & loaction & " "
            Dim cmd As New SqlCommand(sql, conn)
            dr = cmd.ExecuteReader
            If dr.HasRows = True Then
                Dim ii As Integer = 0
                While (dr.Read())
                    DataGridView1.RowCount += 1
                    DataGridView1.Item(clo, i).Value = dr(0)
                    ' MsgBox("the method xx" & dr(0))
                    ' MsgBox(1)
                    i += 1
                End While
                Account = Account + i
            End If
            dr.Close()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        clo += 1
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        'Dim i As Integer
        'Dim PercentError As Integer = 0
        'For i = 0 To 21
        '    ' MsgBox(arrError(i))
        '    PercentError = PercentError + arrError(i)
        'Next
        'MsgBox("the re = " & PercentError)
        'MsgBox(315 - PercentError)
        'PercentError = (315 - PercentError)
        'PercentError = PercentError * 100 / 315
        'MsgBox("%" & PercentError)

        MsgBox("The correct Samples = " & Account & vbNewLine & " The Percent = " & Int(Account * 100 / 315) _
      & vbNewLine & " The Error Samples = " & Account - 315 & vbNewLine & " The Percent = " & Int((Account - 315) * 100 / 315))
    End Sub
    Sub CreateDataGridView1()
        DataGridView1.ColumnCount = 22
        DataGridView1.Columns(0).Name = "user 1"
        DataGridView1.Columns(1).Name = "user 2"
        DataGridView1.Columns(2).Name = "user 3"
        DataGridView1.Columns(3).Name = "user 4"
        DataGridView1.Columns(4).Name = "user 5"
        DataGridView1.Columns(5).Name = "user 6"
        DataGridView1.Columns(6).Name = "user 7"
        DataGridView1.Columns(7).Name = "user 8"
        DataGridView1.Columns(8).Name = "user 9"
        DataGridView1.Columns(9).Name = "user 10"
        DataGridView1.Columns(10).Name = "user 11"
        DataGridView1.Columns(11).Name = "user 12"
        DataGridView1.Columns(12).Name = "user 13"
        DataGridView1.Columns(13).Name = "user 14"
        DataGridView1.Columns(14).Name = "user 15"
        DataGridView1.Columns(15).Name = "user 16"
        DataGridView1.Columns(16).Name = "user 17"
        DataGridView1.Columns(17).Name = "user 18"
        DataGridView1.Columns(18).Name = "user 19"
        DataGridView1.Columns(19).Name = "user 20"
        DataGridView1.Columns(20).Name = "user 21"
    End Sub

    Private Sub Button3_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        addgroup()
        Dim i As Integer
        For i = 0 To 20
            xx(nameUser(i) & "%", group(i))
        Next
    End Sub
    Sub addgroup()
        Dim i As Integer
        Try
            Dim sql As String = "select * from finalOutput130 "
            Dim cmd As New SqlCommand(sql, conn)
            dr = cmd.ExecuteReader
            If dr.HasRows = True Then
                Dim ii As Integer = 0
                While (dr.Read())
                    group(i) = dr(1)
                    i += 1
                End While
            End If
            dr.Close()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub


End Class
