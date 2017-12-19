Imports System.Data.SqlClient
Imports Excel = Microsoft.Office.Interop.Excel
Public Class Form2
    Dim objForm1 As New Form1
    Dim dr As SqlDataReader
    Dim arrays() As String
    Dim siize As Integer = 0
    Private Sub Form2_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        objForm1.getConn().Open()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
       
        Try


            siize = Int(TextBox1.Text)
            ReDim arrays(siize)
            showMoreChar()
            Dim o As Integer = 0
            Dim i As Integer = 0
            Dim arrClo(siize) As Integer
            Dim c As Integer = 0
            Dim flag As Boolean = True
            For i = 0 To arrClo.Length - 1
                arrClo(i) = 0
            Next
            Dim bb As Integer = 1
            i = 0
            For o = 0 To siize - 1
                Dim sql As String = "select tablename,avg from MoreCharacters where char = '" & arrays(o) & "' order by tablename ASC  " 'where char ='" & arrays(o) & "'
                Dim cmd As New SqlCommand(sql, objForm1.conn)
                cmd = New System.Data.SqlClient.SqlCommand(sql, objForm1.conn)
                dr = cmd.ExecuteReader
                If dr.HasRows = True Then
                    Do While (dr.Read)
                        Do While (flag)
                            '  MsgBox("1 " & dr(0) & " = " & DataGridView1.Item(0, arrClo(o)).Value.ToString)
                            If dr(0) = DataGridView1.Item(0, arrClo(o)).Value.ToString Then
                                DataGridView1.Item(1 + o, arrClo(o)).Value = dr(1) ' dr(1)
                                flag = False
                            End If
                            i += 1
                            arrClo(o) = arrClo(o) + 1
                        Loop
                        flag = True
                    Loop
                End If
                dr.Close()
            Next
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        SwitchValueEmpty()


    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Form3.Show()
    End Sub

    Sub showMoreCharInDataGridView1()
        DataGridView1.ColumnCount = siize + 1
        DataGridView1.Columns(0).Name = "table name"

        Dim ii As Integer = 0
        For ii = 1 To siize
            DataGridView1.Columns(ii).Name = """" & arrays(ii - 1) & """"
        Next
        Dim i As Integer = 0
        Dim sql As String = "select tablename from MoreCharacters group by tablename order by tablename ASC  "
        Dim cmd As New SqlCommand(sql, objForm1.conn)
        cmd = New System.Data.SqlClient.SqlCommand(sql, objForm1.conn)
        dr = cmd.ExecuteReader
        If dr.HasRows = True Then
            Do While (dr.Read)
                DataGridView1.RowCount += 1
                DataGridView1.Item(0, i).Value = dr(0)
                i += 1
            Loop
        End If
        dr.Close()
    End Sub
    Sub showMoreChar()
        Dim i As Integer = 0
        Dim conuter As Integer = 0
        Dim sql As String = "select sum(conu),char,avg(avg) from MoreCharacters group by char order by sum(conu) DESC  "
        Dim cmd As New SqlCommand(sql, objForm1.conn)
        cmd = New System.Data.SqlClient.SqlCommand(sql, objForm1.conn)
        dr = cmd.ExecuteReader
        If dr.HasRows = True Then
            Do While (dr.Read)
                i += 1
                arrays(conuter) = dr.Item("char")
                conuter = conuter + 1
                If conuter >= siize Then
                    Exit Do
                End If
            Loop
        End If
        dr.Close()
        showMoreCharInDataGridView1()
    End Sub

    Private Sub Button3_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub
    Sub SwitchValueEmpty()
        Dim r As Integer = 0
        Dim i As Integer = 0
        Dim j As Integer = 0
        For i = 1 To (DataGridView1.Columns.Count - 1)
            For j = 0 To (DataGridView1.Rows.Count - 2)
                If DataGridView1.Item(i, j).Value Is Nothing Then
                    Dim sql As String = "select avg(avg) from MoreCharacters where char = '" & arrays(i - 1) & "'  "
                    Dim cmd As New SqlCommand(sql, objForm1.conn)
                    cmd = New System.Data.SqlClient.SqlCommand(sql, objForm1.conn)
                    cmd.CommandType = CommandType.Text
                    Dim value As Integer = cmd.ExecuteScalar.ToString
                    DataGridView1.Item(i, j).Value = value
                End If
            Next
        Next i
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim rowsTotal, colsTotal As Short
        Dim I, j, iC As Short
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        Dim xlApp As New Excel.Application
        Try
            Dim excelBook As Excel.Workbook = xlApp.Workbooks.Add
            Dim excelWorksheet As Excel.Worksheet = CType(excelBook.Worksheets(1), Excel.Worksheet)
            xlApp.Visible = True
            rowsTotal = DataGridView1.RowCount - 1
            colsTotal = DataGridView1.Columns.Count - 1
            With excelWorksheet
                .Cells.Select()
                .Cells.Delete()
                For iC = 0 To colsTotal
                    .Cells(1, iC + 1).Value = DataGridView1.Columns(iC).HeaderText
                Next
                For I = 0 To rowsTotal
                    For j = 0 To colsTotal
                        .Cells(I + 2, j + 1).value = DataGridView1.Rows(I).Cells(j).Value
                    Next j
                Next I
                .Rows("1:1").Font.FontStyle = "Bold"
                .Rows("1:1").Font.Size = 10
                .Cells.Columns.AutoFit()
                .Cells.Select()
                .Cells.EntireColumn.AutoFit()
                .Cells(1, 1).Select()
            End With
        Catch ex As Exception
            MsgBox("Export Excel Error " & ex.Message)
        Finally
            'RELEASE ALLOACTED RESOURCES
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
            xlApp = Nothing
        End Try
    End Sub

    Private Sub Panel1_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Panel1.Paint

    End Sub
End Class