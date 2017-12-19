Imports System
Imports System.Data.SqlClient
Imports System.IO
Imports Excel = Microsoft.Office.Interop.Excel

Public Class AnalyzingData-Phase2
    Dim dr As SqlDataReader
    Dim cmd As New SqlCommand
    Dim nameUser(21) As String
    Dim arrError(21) As Integer
    Dim arrMaxC(21, 6)
    Dim group(21) As Integer
    Dim clo As Integer = 0
    Public conn As New SqlConnection("Data Source=.\SQLEXPRESS;AttachDbFilename=" & Application.StartupPath & "\output.mdf;Integrated Security=True;Connect Timeout=30;User Instance=True")

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try 'Connection DB
            conn.Open()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        nameAllUser()
    End Sub

    Sub nameAllUser()
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

    Sub cc()
        DataGridView1.ColumnCount = 7
        DataGridView1.Columns(0).Name = "user"
        DataGridView1.Columns(1).Name = "group"
        DataGridView1.Columns(2).Name = "count"
        DataGridView1.Columns(3).Name = "group"
        DataGridView1.Columns(4).Name = "count"
        DataGridView1.Columns(5).Name = "group"
        DataGridView1.Columns(6).Name = "count"
        Dim i As Integer
        Dim ii As Integer = 0
        For i = 0 To 20
            'MsgBox(arrMaxC(i, 0))
            DataGridView1.RowCount += 1
            DataGridView1.Item(0, ii).Value = nameUser(i)
            DataGridView1.Item(1, ii).Value = arrMaxC(i, 0)
            DataGridView1.Item(2, ii).Value = arrMaxC(i, 1)
            DataGridView1.Item(3, ii).Value = arrMaxC(i, 2)
            DataGridView1.Item(4, ii).Value = arrMaxC(i, 3)
            DataGridView1.Item(5, ii).Value = arrMaxC(i, 4)
            DataGridView1.Item(6, ii).Value = arrMaxC(i, 5)
            ii += 1
        Next
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
    Sub printGroup()
        DataGridView2.ColumnCount = 2
        DataGridView2.Columns(0).Name = " user "
        DataGridView2.Columns(1).Name = " Group "
        Dim i As Integer
        Dim ii As Integer
        For i = 0 To 20
            '   If group(i) > 0 Then
            '  MsgBox("the i = " & i + 1 & "  " & group(i))
            DataGridView2.RowCount += 1
            DataGridView2.Item(0, ii).Value = nameUser(i)
            DataGridView2.Item(1, ii).Value = group(i)
            ii += 1
            ' End If
        Next
    End Sub
    Function test(ByVal x As Integer) As Boolean
        Dim i As Integer
        For i = 0 To 20
            If group(i) = x Then
                'MsgBox("the false = " & x)
                test = True
                Exit Function
            End If
        Next
        test = False

    End Function


    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Dim i As Integer = 0
        Dim conut As Integer = 15
        Dim j As Integer
        For j = 0 To 20
            For i = 0 To 20
                If arrMaxC(i, 1) = conut Then
                    If test(arrMaxC(i, 0)) Then
                        group(i) = 0
                    Else
                        group(i) = arrMaxC(i, 0)
                    End If
                End If
            Next
            conut = conut - 1
        Next


        i = 0
        conut = 15

        For j = 0 To 20
            conut = 5
            For i = 0 To 20
                If group(i) = 0 Then
                    If arrMaxC(i, 3) = conut Then ' 4 Or 2 Or 3 Or 1
                        If test(arrMaxC(i, 2)) Then
                            group(i) = 0
                        Else
                            group(i) = arrMaxC(i, 2)
                        End If
                    End If
                End If
                conut = conut - 1
            Next
        Next


        i = 0
        conut = 15
        For j = 0 To 20
            conut = 5
            For i = 0 To 20
                If group(i) = 0 Then
                    If arrMaxC(i, 5) = 4 Or 2 Or 3 Or 1 Or 5 Then '4 Or 2 Or 3 Or 1 Or 5
                        If test(arrMaxC(i, 4)) Then
                            group(i) = 0
                        Else
                            group(i) = arrMaxC(i, 4)
                        End If
                    End If
                End If
                conut = conut - 1
            Next
        Next
        printGroup()
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Dim rowsTotal, colsTotal As Short
        Dim I, j, iC As Short
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        Dim xlApp As New Excel.Application
        Try
            Dim excelBook As Excel.Workbook = xlApp.Workbooks.Add
            Dim excelWorksheet As Excel.Worksheet = CType(excelBook.Worksheets(1), Excel.Worksheet)
            xlApp.Visible = True
            rowsTotal = DataGridView2.RowCount - 1
            colsTotal = DataGridView2.Columns.Count - 1
            With excelWorksheet
                .Cells.Select()
                .Cells.Delete()
                For iC = 0 To colsTotal
                    .Cells(1, iC + 1).Value = DataGridView2.Columns(iC).HeaderText
                Next
                For I = 0 To rowsTotal
                    For j = 0 To colsTotal
                        .Cells(I + 2, j + 1).value = DataGridView2.Rows(I).Cells(j).Value
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


    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim arr(21) As Integer
        Dim index As Integer
        Dim i As Integer
        Dim conuterGroup As Integer
        Dim a As String
        For i = 0 To 20 '21
            a = nameUser(i) & "%"
            Try
                Dim sql As String = "select * from char130 where Observation like '" & a & "'"  ' "select class from char100 where Observation like 'user1-%' group by class order by class ASC    "
                Dim cmd As New SqlCommand(sql, conn)
                dr = cmd.ExecuteReader
                If dr.HasRows = True Then
                    Dim ii As Integer = 0
                    While (dr.Read())
                        index = dr.Item("class")
                        arr(index) = arr(index) + 1
                    End While
                End If
                dr.Close()
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
            Dim max As Integer = arr(0)
            Dim location1 As Integer
            Dim location2 As Integer
            Dim location3 As Integer
            For index = 1 To arr.Length - 1
                If arr(index) > max Then
                    max = arr(index)
                    location1 = index
                End If
            Next
            conuterGroup = conuterGroup + arr(location1)
            arrMaxC(i, 0) = location1
            arrMaxC(i, 1) = arr(location1)
            arr(location1) = 0
            max = 0
            For index = 1 To arr.Length - 1
                If arr(index) > max Then
                    max = arr(index)
                    location2 = index
                End If
            Next
            arrMaxC(i, 2) = location2
            arrMaxC(i, 3) = arr(location2)

            arr(location2) = 0
            max = 0
            For index = 1 To arr.Length - 1
                If arr(index) > max Then
                    max = arr(index)
                    location3 = index
                End If
            Next
            arrMaxC(i, 4) = location3
            arrMaxC(i, 5) = arr(location3)


            ' all = 0
            max = 0
            location1 = 0
            location2 = 0
            location3 = 0
            index = 0
            For index = 0 To arr.Length - 1
                arr(index) = 0
            Next
        Next
        'Dim j As Integer = 0
        'For i = 0 To 20
        '    MsgBox("the i = " & i & "   /  " & arrMaxC(i, 0) & arrMaxC(i, 1) & arrMaxC(i, 2) & arrMaxC(i, 3))
        'Next()
        '   MsgBox(conuterGroup)
        cc()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Button2_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Group1.Show()
    End Sub
End Class
