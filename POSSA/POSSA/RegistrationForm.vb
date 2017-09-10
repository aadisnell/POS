Imports System.Runtime.InteropServices
Imports System.Data.OleDb
Imports iTextSharp.text.pdf
Imports iTextSharp.text
Imports System.IO
Imports Excel = Microsoft.Office.Interop.Excel


Public Class RegistrationForm

    ' Declare User32 constants and methods.
    Private Const MF_BYPOSITION As Integer = &H400

    <DllImport("user32.dll", CallingConvention:=CallingConvention.Cdecl)> _
    Private Shared Function GetSystemMenu(ByVal hWnd As IntPtr, ByVal bRevert As Boolean) As IntPtr
    End Function

    <DllImport("user32.dll")> _
    Public Shared Function GetMenuItemCount(ByVal hMenu As IntPtr) As Int32
    End Function

    <DllImport("user32.dll")> _
    Public Shared Function RemoveMenu(ByVal hMenu As IntPtr, ByVal nPosition As Int32, ByVal wFlags As Int32) As Int32
    End Function

    'provider is the technology used to connect to access database
    Dim provider As String
    'datafile is the data source of the database
    Dim datafile As String
    'connstring is the provider and the datafile
    Dim connstring As String = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=D:\POSSA\POSSA\POSSA\POSSA.accdb"
    'myConnection is an instance of the object oledbconnection
    Dim myConnection As OleDbConnection = New OleDbConnection
    Dim MyConn As OleDbConnection
    Dim da As OleDbDataAdapter
    Dim ds As DataSet
    Dim tables As DataTableCollection
    Dim source1 As New BindingSource

    Private Sub RegistrationForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.TextBox1.Select()

        Dim hMenu As IntPtr = GetSystemMenu(Me.Handle, False)
        Dim num_menu_items As Integer = GetMenuItemCount(hMenu)

        'For i As Integer = num_menu_items - 1 To 0 Step -1
        '    RemoveMenu(hMenu, i, MF_BYPOSITION)
        'Next i
        RemoveMenu(hMenu, num_menu_items - 1, MF_BYPOSITION)
        RemoveMenu(hMenu, num_menu_items - 2, MF_BYPOSITION)
    End Sub





    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Application.Restart()
    End Sub

    Private Sub REGISTATIONFORMBindingSource_CurrentChanged(sender As Object, e As EventArgs)

    End Sub

    '    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
    '        REGISTATIONFORMBindingSource.MovePrevious()
    '    End Sub

    '    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
    '        REGISTATIONFORMBindingSource.AddNew()
    '    End Sub

    '    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
    '        REGISTATIONFORMBindingSource.MoveNext()
    '    End Sub

    '    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click

    '        On Error GoTo SaveErr

    '        REGISTATIONFORMBindingSource.EndEdit()
    '        REGISTATION_FORMTableAdapter.Update(POSDataSet.REGISTATION_FORM)
    '        MessageBox.Show("Successfully saved")

    'SaveErr:

    '        Exit Sub
    '    End Sub

    '    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
    '        REGISTATIONFORMBindingSource.RemoveCurrent()
    '    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        If (TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox10.Text = "" Or TextBox8.Text = "" _
            Or TextBox5.Text = "" Or TextBox9.Text = "") Then
            MsgBox("YOU HAVE NOT PROVIDED ALL THE REQUIRED INFO FOR A SUCCESFUL REGISTRATION")
            Me.TextBox1.Select()
        Else
            provider = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source="
            datafile = "D:\POSSA\POSSA\POSSA\POSSA.accdb"
            connstring = provider & datafile
            myConnection.ConnectionString = connstring

            'connection is open 
            myConnection.Open()


            'Str stands for the variable for inserting data into the registration form table 
            Dim str As String
            str = "insert into REGISTATION_FORM([GENDER],[SURNAME],[FIRSTNAME],[OTHERNAME],[PHONE NUM], [ID NUM], [AMOUNT PAID],[LEVEL],[DATE PAID]) values(?,?,?,?,?,?,?,?,?)"
            Dim cmd As OleDbCommand = New OleDbCommand(str, myConnection)
            cmd.Parameters.Add(New OleDbParameter("GENDER", CType(TextBox1.Text, String)))
            cmd.Parameters.Add(New OleDbParameter("SURNAME", CType(TextBox2.Text, String)))
            cmd.Parameters.Add(New OleDbParameter("FIRSTNAME", CType(TextBox3.Text, String)))
            cmd.Parameters.Add(New OleDbParameter("OTHERNAME", CType(TextBox4.Text, String)))
            cmd.Parameters.Add(New OleDbParameter("PHONE NUM", CType(TextBox6.Text, String)))
            cmd.Parameters.Add(New OleDbParameter("ID NUM", CType(TextBox5.Text, String)))
            cmd.Parameters.Add(New OleDbParameter("AMOUNT PAID", CType(TextBox8.Text, String)))
            cmd.Parameters.Add(New OleDbParameter("LEVEL", CType(TextBox9.Text, String)))
            cmd.Parameters.Add(New OleDbParameter("DATE PAID", CType(TextBox10.Text, String)))
            Try
                cmd.ExecuteNonQuery()
                cmd.Dispose()
                myConnection.Close()

            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
            MsgBox("REGISTRATION SUCESSFUL")
            TextBox1.Text = ""
            TextBox2.Clear()
            TextBox3.Clear()
            TextBox4.Clear()
            TextBox5.Clear()
            TextBox6.Clear()
            TextBox8.Clear()
            TextBox9.Text = ""
            TextBox10.Clear()
            Me.TextBox1.Select()

        End If
    End Sub

    Private Sub TextBox2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox8.KeyPress

        'This section of code allows only numeric numbers to be inputed into the textbox. It also makes exception to the key backspace(Asc 8)
        'and dot(.)(Asc 46)

        If (Asc(e.KeyChar) >= 48 And Asc(e.KeyChar) <= 57) Or Asc(e.KeyChar) = 8 Or Asc(e.KeyChar) = 46 Then
            e.Handled = False
        Else
            e.Handled = True
        End If
    End Sub





    Private Sub TextBox5_KeyPress(sender As Object, e As KeyPressEventArgs)

        'This section of code allows only numeric numbers to be inputed into the textbox. It also makes exception to the key backspace(Asc 8)

        If (Asc(e.KeyChar) >= 48 And Asc(e.KeyChar) <= 57) Or Asc(e.KeyChar) = 8 Then
            e.Handled = False
        Else
            e.Handled = True
        End If

    End Sub



    Private Sub TextBox7_KeyPress(sender As Object, e As KeyPressEventArgs)
        If Asc(e.KeyChar) < 65 And Asc(e.KeyChar) > 90 Or
        Asc(e.KeyChar) > 97 And Asc(e.KeyChar) < 122 Then
            e.Handled = True
            MessageBox.Show("YOU CAN ONLY INPUT UPPERCASE LETTERS")
        End If
    End Sub

    Private Sub TextBox6_KeyPress(sender As Object, e As KeyPressEventArgs)
        If Asc(e.KeyChar) < 65 And Asc(e.KeyChar) > 90 Or
        Asc(e.KeyChar) > 97 And Asc(e.KeyChar) < 122 Then
            e.Handled = True
            MessageBox.Show("YOU CAN ONLY INPUT UPPERCASE LETTERS")
        End If
    End Sub




    Private Sub TextBox6_TextChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub TextBox9_KeyPress(sender As Object, e As KeyPressEventArgs)
        If (Asc(e.KeyChar) >= 48 And Asc(e.KeyChar) <= 57) Or Asc(e.KeyChar) = 8 Then
            e.Handled = False
        Else
            e.Handled = True
        End If
    End Sub


    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        MyConn = New OleDbConnection
        MyConn.ConnectionString = connstring
        ds = New DataSet
        tables = ds.Tables
        da = New OleDbDataAdapter("Select * from [REGISTATION_FORM]", MyConn) 'Change items to your database name
        da.Fill(ds, "REGISTATION_FORM") 'Change items to your database name
        Dim view As New DataView(tables(0))
        source1.DataSource = view
        DataGridView1.DataSource = view

        'configuration for save file dialog







    End Sub
    'Private Sub Search()
    '    Dim connstring As String = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=C:\POS\POS\POS.accdb"
    '    'myConnection is an instance of the object oledbconnection
    '    Dim myConnection As OleDbConnection = New OleDbConnection

    '    Dim dt As New DataTable("REGISTATION_FORM")
    '    Dim da As New OleDbDataAdapter("Select * from REGISTATION_FORM where ID NUM='" & "'", myConnection)
    '    da.Fill(dt)

    '    DataGridView1.DataSource = dt
    '    DataGridView1.Refresh()

    '    ds.Dispose()

    '    myConnection.Close()




    'End Sub

    'exporting data to text document for easy viewing cos i dont want the client to temper the acces database file .....
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click


        'This section of code prompts the user for and administrative password when the Export to text Button is Clicked.

        'Dim writer As TextWriter = New StreamWriter("D:\POSSA\POSSA\TextGenerated\Exported_Data.text")
        'For i As Integer = 0 To DataGridView1.Rows.Count - 2 Step +1

        'For j As Integer = 0 To DataGridView1.Columns.Count - 1 Step +1

        'writer.Write(vbTab & DataGridView1.Rows(i).Cells(j).Value.ToString() & vbTab & "|")

        'Next
        'writer.WriteLine("")
        'writer.WriteLine(".......................................................................................................................................................................................")

        'Next
        'writer.Close()
        'MessageBox.Show("DATA EXPORTED")


        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet
        Dim misValue As Object = System.Reflection.Missing.Value

        Dim i As Integer
        Dim j As Integer


        xlApp = New Microsoft.Office.Interop.Excel.Application
        xlWorkBook = xlApp.Workbooks.Add(misValue)
        xlWorkSheet = xlWorkBook.Sheets("sheet1")

        For i = 0 To DataGridView1.RowCount - 2
            For j = 0 To DataGridView1.ColumnCount - 1
                For k As Integer = 1 To DataGridView1.Columns.Count
                    xlWorkSheet.Cells(1, k) = DataGridView1.Columns(k - 1).HeaderText
                    'xlWorkSheet.Cells(i + 2, j + 1) = DataGridView1(j, i).Value.ToString()


                    xlWorkSheet.Cells(i + 2, j + 1) = _
                    DataGridView1(j, i).Value.ToString()

                Next

            Next
        Next

        xlWorkSheet.SaveAs("D:\POSSA\EXCEL_FILE\EXPORTED_DATA.xlsx")
        xlWorkBook.Close()
        xlApp.Quit()

        releaseObject(xlApp)
        releaseObject(xlWorkBook)
        releaseObject(xlWorkSheet)

        Dim res As MsgBoxResult
        res = MsgBox("DATA IS EXPORTED...WOULD YOU LIKE TO OPEN THE FILE?", MsgBoxStyle.YesNo)
        If (res = MsgBoxResult.Yes) Then
            Process.Start("D:\POSSA\EXCEL_FILE\EXPORTED_DATA.xlsx")
        End If



    End Sub
    Private Sub releaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing

        Finally
            GC.Collect()

        End Try
    End Sub

    Private Sub TextBox10_KeyPress(sender As Object, e As KeyPressEventArgs)

        'This section of code allows only numeric numbers to be inputed into the textbox. It also makes exception to the key backspace(Asc 8)
        'and the forward slash sign(/)(Asc 47)

        If (Asc(e.KeyChar) >= 48 And Asc(e.KeyChar) <= 57) Or Asc(e.KeyChar) = 8 Or Asc(e.KeyChar) = 47 Then
            e.Handled = False
        Else
            e.Handled = True
        End If

    End Sub



    Private Sub TextBox2_KeyPress_1(sender As Object, e As KeyPressEventArgs) Handles TextBox2.KeyPress
        If Asc(e.KeyChar) < 65 And Asc(e.KeyChar) > 90 Or
           Asc(e.KeyChar) > 97 And Asc(e.KeyChar) < 122 Then
            e.Handled = True
            MessageBox.Show("YOU CAN ONLY INPUT UPPERCASE LETTERS")
        End If
    End Sub

    Private Sub TextBox3_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox3.KeyPress
        If Asc(e.KeyChar) < 65 And Asc(e.KeyChar) > 90 Or
           Asc(e.KeyChar) > 97 And Asc(e.KeyChar) < 122 Then
            e.Handled = True
            MessageBox.Show("YOU CAN ONLY INPUT UPPERCASE LETTERS")
        End If
    End Sub

    Private Sub TextBox4_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox4.KeyPress
        If Asc(e.KeyChar) < 65 And Asc(e.KeyChar) > 90 Or
        Asc(e.KeyChar) > 97 And Asc(e.KeyChar) < 122 Then
            e.Handled = True
            MessageBox.Show("YOU CAN ONLY INPUT UPPERCASE LETTERS")
        End If
    End Sub

    Private Sub TextBox5_KeyPress_1(sender As Object, e As KeyPressEventArgs) Handles TextBox5.KeyPress
        If (Asc(e.KeyChar) >= 48 And Asc(e.KeyChar) <= 57) Or Asc(e.KeyChar) = 8 Then
            e.Handled = False
        Else
            e.Handled = True
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        TextBox1.Text = ""
        TextBox2.Clear()
        TextBox3.Clear()
        TextBox4.Clear()
        TextBox5.Clear()
        TextBox6.Clear()
        TextBox8.Clear()
        TextBox9.Text = ""
        TextBox10.Clear()
        Me.TextBox1.Select()
    End Sub
End Class