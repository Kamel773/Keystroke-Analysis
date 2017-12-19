Imports System
Imports System.Data.SqlClient
Imports System.IO
Imports Excel = Microsoft.Office.Interop.Excel

Public Class Form1
    Dim NameFile As String
    Dim fiter_user As String
    Dim dr As SqlDataReader
    Dim cmd As New SqlCommand
    Dim arrnameTable(4000) As String
    Dim nametable(sizeNametable) As String
    Dim sizeNametable As Integer
    Public conn As New SqlConnection("Data Source=.\SQLEXPRESS;AttachDbFilename=" & Application.StartupPath & "\dr.man.mdf;Integrated Security=True;Connect Timeout=30;User Instance=True")
    Function getConn() As SqlConnection
        getConn = conn
    End Function

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try 'Connection DB
            conn.Open()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        '   addDataOnDB1("oo")
    End Sub
    Public Sub getdataFromTxtFile()
        Try ' imports data from text file and sent to DB
            Dim ioFile As New StreamReader(OpenFileDialog1.FileName) '"d:\user1-2.txt"
            Dim milli As String ' Going to hold one line at a time
            Dim ASCII As String ' Going to hold whole file
            Dim nan As String
            Dim chaar As String
            '       Dim timen As String
            chaar = ioFile.ReadLine
            ASCII = ioFile.ReadLine
            milli = ioFile.ReadLine
            nan = ioFile.ReadLine
            '   MsgBox(ioFile.ReadToEnd)
            '  timen = ioFile.ReadLine
            While Not milli = ""
                '  MsgBox("the ASCII " & ASCII & "  the mliil  " & milli)
                addDataOnDB(milli, ASCII, nan)
                chaar = ioFile.ReadLine
                ASCII = ioFile.ReadLine
                milli = ioFile.ReadLine
                nan = ioFile.ReadLine
                '  timen = ioFile.ReadLine
            End While
            MsgBox("yes .the data add in databeas")
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

    End Sub

    Public Sub addDataOnDB(ByVal milli As String, ByVal ASCII As String, ByVal nan As String)

        Try ' insert data in table
            Dim sql As String = "insert into " & NameFile & " (ASCII,milli,nan) values('" & ASCII & "','" & milli & "','" & nan & "')"  ' values('" & milli & "','" & ASCII & "')"
            Dim cmd As New SqlCommand(sql, conn)
            cmd.Connection = conn
            cmd.ExecuteNonQuery()
            '  MsgBox("ADD")
        Catch ex As Exception
            '    MsgBox(ex.Message)
        End Try




    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        OpenFileDialog1.ShowDialog()
        Dim T() As String
        T = Split(OpenFileDialog1.FileName, "\", )
        NameFile = T(UBound(T))
        ' getdataFromTxtFile()
        Dim count As Integer = NameFile.Length
        '  MsgBox(NameFile.Substring(0, NameFile.Length - 4))
        NameFile = NameFile.Substring(0, NameFile.Length - 4)
        TextBox1.Text = NameFile
        fiter_user = "filter_" + NameFile
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try ' crate table 
            '  MsgBox(NameFile.ToString)
            NameFile = TextBox1.Text
            Dim sql As String = "CREATE TABLE " & NameFile & " ( ASCII VARCHAR(30),milli  VARCHAR(30),nan  VARCHAR(30),) "
            Dim cmd As New SqlCommand(sql, conn)
            cmd.Connection = conn
            cmd.ExecuteNonQuery()
            '  MsgBox("ADD")
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        getdataFromTxtFile()
    End Sub




    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        DataGridView1.Rows.Clear()
        DataGridView2.Rows.Clear()
        DataGridView3.Rows.Clear()
        DataGridView4.Rows.Clear()
        DataGridView5.Rows.Clear()
        DataGridView6.Rows.Clear()

    End Sub
    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        '  MsgBox(Asc(65))
        'MsgBox("|" & ChrW(10) & "|  ")
        Dim i As Integer = 0
        Try
            Dim sql As String = "SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE (TABLE_NAME LIKE 'user%')" ' user%
            Dim cmd As New SqlCommand(sql, conn)
            dr = cmd.ExecuteReader
            If dr.HasRows = True Then
                While (dr.Read())
                    arrnameTable(i) = dr(0)
                    i += 1
                End While
            End If
            dr.Close()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        'For i = 0 To arrnameTable.Length - 2
        '    If arrnameTable(i) Is Nothing Then
        '        Exit For
        '    End If
        '    MsgBox(" i = " & i & "  /  " & arrnameTable(i))
        'Next

        i = 0
        For i = 0 To arrnameTable.Length - 2
            If arrnameTable(i) Is Nothing Then
                Exit For
            End If
            NameFile = arrnameTable(i)
            showDataAndFilter1()
            SaveDataAndFilter1()
            showDataAndFilter2()
            SaveDataAndFilter2()
            MoreCharacters()
            DataGridView1.Rows.Clear()
            DataGridView2.Rows.Clear()
            DataGridView3.Rows.Clear()
            DataGridView4.Rows.Clear()
        Next

    End Sub
    Sub showDataAndFilter1()
        Dim exitloop As Integer = 0
        Dim conuArr As Integer = 0
        DataGridView1.ColumnCount = 4
        DataGridView1.Columns(0).Name = "char"
        DataGridView1.Columns(1).Name = "ASCIIcode"
        DataGridView1.Columns(2).Name = "milli"
        DataGridView1.Columns(3).Name = "nano"
        Dim rr As Integer = 0
        Try
            Dim dr As SqlDataReader
            Dim sql As String = "select * from " & NameFile & " " '"select * from " & NameFile & "  "
            Dim cmd As New SqlCommand(sql, conn)
            dr = cmd.ExecuteReader
            If dr.HasRows = True Then
                Dim flag As Boolean
                Dim char1 As String = ""
                Dim timemili1 As String = 0
                Dim timenano1 As String = 0
                While (dr.Read)
                    Me.DataGridView1.Rows.Add(ChrW(dr.Item("ASCII")), dr.Item("ASCII"), dr.Item("milli"), dr.Item("nan"))
                    If flag Then
                        ' MsgBox(char1 & dr.Item("char"))
                        '  addDataOnDB1(char1 & dr.Item("char"))
                        DataGridView2.ColumnCount = 3
                        DataGridView2.Columns(0).Name = "char"
                        DataGridView2.Columns(1).Name = "milli"
                        DataGridView2.Columns(2).Name = "nano"
                        ' DataGridView2.Rows.Add(char1 & dr.Item("char"))
                        Dim Filtermilli As Integer = dr.Item("milli") - timemili1
                        Dim Filternan As Long = dr.Item("nan") - timenano1
                        '  fliter
                        If Filtermilli > 500 Then
                            Filtermilli = 0
                            Filternan = 0
                        End If
                        'conuArr = conuArr + 1
                        'If conuArr > 0 Then
                        Me.DataGridView2.Rows.Add(char1 & ChrW(dr.Item("ASCII")), Filtermilli, Filternan)
                        'End If
                        'If conuArr > 5 Then
                        '    Exit While
                        'End If
                    End If
                    flag = True
                    char1 = ChrW(dr.Item("ASCII"))
                    timemili1 = dr.Item("milli")
                    timenano1 = dr.Item("nan")

                    'exitloop = exitloop + 1
                    'If exitloop = 100 Then
                    '    Exit While

                    'End If

                End While
            Else
                MsgBox("Not Found")
            End If
            dr.Close()
        Catch ex As Exception
            MsgBox(ex.ToString)
            MsgBox("show filter 1 " & NameFile)
        End Try

    End Sub
    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click

    End Sub
    Sub SaveDataAndFilter1()
        Try


            Try ' crate table 
                '  MsgBox(NameFile.ToString)
                ' NameFile = TextBox1.Text
                fiter_user = "filter_" + NameFile
                '  MsgBox(fiter_user)
                Dim sql As String = "CREATE TABLE " & fiter_user & " (char VARCHAR(30) ,milli int, ) " '" & fiter_user & "
                Dim cmd As New SqlCommand(sql, conn)
                cmd.Connection = conn
                cmd.ExecuteNonQuery()
                '  MsgBox("ADD")
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try

            ' MsgBox(fiter_user)
            Dim i As Integer ' add data form DataGridView2 to BD
            For i = 0 To DataGridView2.Rows.Count - 2
                Dim row As DataGridViewRow = DataGridView2.Rows(i)
                Dim chr As DataGridViewTextBoxCell = row.Cells(0)
                Dim milli As DataGridViewTextBoxCell = row.Cells(1)
                Dim sql As String = "insert into " & fiter_user & "  (char,milli)values('" & chr.Value & "','" & milli.Value & "')" ' " & fiter_user & "
                Dim cmd As New SqlCommand(sql, conn)
                cmd.CommandText = sql
                cmd.Connection = conn
                cmd.ExecuteNonQuery()
            Next
        Catch ex As Exception
            MsgBox(ex.Message)
            '      MsgBox(Chr.Value & "  " & milli.Value & "  " & nan.Value)
        End Try
    End Sub
    Sub showDataAndFilter2()
        Try
            DataGridView3.ColumnCount = 3
            DataGridView3.Columns(0).Name = "char"
            DataGridView3.Columns(1).Name = "conuter"
            DataGridView3.Columns(2).Name = "milli AVG"
            Dim i As Integer = 0
            fiter_user = "filter_" + NameFile
            Dim sql As String = "select count(milli),avg(milli),char from " & fiter_user & " group by char  "
            'Dim cmd As New SqlCommand(sql, conn)
            cmd = New System.Data.SqlClient.SqlCommand(sql, conn)
            'conn.Open()
            dr = cmd.ExecuteReader
            If dr.HasRows = True Then
                Do While (dr.Read)
                    '  Me.DataGridView3.Rows.Add(dr.Item("char"), dr.Item("time"))
                    'If Find(dr.Item("char")) Then
                    'Else
                    ' MsgBox(dr.Item("time"))
                    'searchChar(dr.Item("char"))
                    DataGridView3.RowCount += 1
                    DataGridView3.Item(0, i).Value = dr.Item("char")
                    DataGridView3.Item(1, i).Value = dr(0) ' dr(1)
                    DataGridView3.Item(2, i).Value = dr(1) ' dr(1)
                    ' DataGridView3.Item(1, i).Value = dr.Item("char")
                    i += 1
                    ' Me.DataGridView3.Rows.Add(dr.Item("char")) ' , searchChar(dr.Item("char"))
                    ' dr = cmd.ExecuteReader
                    'End If
                Loop
            End If
            dr.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Sub SaveDataAndFilter2()
        Try


            fiter_user = "filter2_" + NameFile
            Try
                Dim sql As String = "CREATE TABLE " & fiter_user & " (char VARCHAR(30) ,conuter int, milli int ,) " '" & fiter_user & "
                Dim cmd As New SqlCommand(sql, conn)
                cmd.Connection = conn
                cmd.ExecuteNonQuery()
                '  MsgBox("ADD")
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try


            Dim i As Integer ' add data form DataGridView3 to BD
            For i = 0 To DataGridView3.Rows.Count - 2
                Dim row As DataGridViewRow = DataGridView3.Rows(i)
                Dim chr As DataGridViewTextBoxCell = row.Cells(0)
                Dim conuter As DataGridViewTextBoxCell = row.Cells(1)
                Dim milli As DataGridViewTextBoxCell = row.Cells(2)
                Dim sql As String = "insert into " & fiter_user & "  (char,conuter,milli)values('" & chr.Value & "','" & conuter.Value & "','" & milli.Value & "')" ' " & fiter_user & "
                Dim cmd As New SqlCommand(sql, conn)
                cmd.CommandText = sql
                cmd.Connection = conn
                cmd.ExecuteNonQuery()
            Next
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Dim rowsTotal, colsTotal As Short
        Dim I, j, iC As Short
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        Dim xlApp As New Excel.Application
        Try
            Dim excelBook As Excel.Workbook = xlApp.Workbooks.Add
            Dim excelWorksheet As Excel.Worksheet = CType(excelBook.Worksheets(1), Excel.Worksheet)
            xlApp.Visible = True
            rowsTotal = DataGridView3.RowCount - 1
            colsTotal = DataGridView3.Columns.Count - 1
            With excelWorksheet
                .Cells.Select()
                .Cells.Delete()
                For iC = 0 To colsTotal
                    .Cells(1, iC + 1).Value = DataGridView3.Columns(iC).HeaderText
                Next
                For I = 0 To rowsTotal
                    For j = 0 To colsTotal
                        .Cells(I + 2, j + 1).value = DataGridView3.Rows(I).Cells(j).Value
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

    Sub MoreCharacters()
        Try


            DataGridView4.ColumnCount = 5
            DataGridView4.Columns(0).Name = "name table"
            DataGridView4.Columns(1).Name = "char"
            DataGridView4.Columns(2).Name = "conuter"
            DataGridView4.Columns(3).Name = "milli"
            DataGridView4.Columns(4).Name = "nan"
            ' dr.Close()
            fiter_user = "filter2_" + NameFile
            Try
                '  Dim dr As SqlDataReader
                Dim sql As String = "select * from " & fiter_user & " order by conuter DESC " '"select * from " & NameFile & "  "
                Dim cmd As New SqlCommand(sql, conn)
                dr = cmd.ExecuteReader
                If dr.HasRows = True Then
                    Dim ii As Integer = 0
                    While (dr.Read())
                        If (dr.Item("milli") > 0) Then
                            Me.DataGridView4.Rows.Add(fiter_user, dr.Item("char"), dr.Item("conuter"), dr.Item("milli"))
                            ii = ii + 1
                            If ii = 100 Then
                                Exit While
                            End If
                        End If
                    End While
                End If
                dr.Close()
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
            Dim i As Integer ' add data form DataGridView4 to BD
            For i = 0 To DataGridView4.Rows.Count - 2
                Dim row As DataGridViewRow = DataGridView4.Rows(i)
                Dim tablename As DataGridViewTextBoxCell = row.Cells(0)
                Dim chr As DataGridViewTextBoxCell = row.Cells(1)
                Dim conu As DataGridViewTextBoxCell = row.Cells(2)
                Dim milli As DataGridViewTextBoxCell = row.Cells(3)
                Dim nan As DataGridViewTextBoxCell = row.Cells(4)
                Dim sql As String = "insert into MoreCharacters  (tableName,char,conu,avg)values('" & tablename.Value.substring(8) & "','" & chr.Value & "','" & conu.Value & "','" & milli.Value & "')" ' " & fiter_user & "
                Dim cmd As New SqlCommand(sql, conn)
                cmd.CommandText = sql
                cmd.Connection = conn
                cmd.ExecuteNonQuery()
            Next
        Catch ex As Exception
            'MsgBox("more char " & fiter_user)
            MsgBox(ex.Message)
        End Try
    End Sub


End Class

