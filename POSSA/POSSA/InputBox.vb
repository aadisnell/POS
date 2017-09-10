Imports System.Data.OleDb

Public Class InputBox
    'provider is the technology used to connect to access database 
    Dim provider As String

    'datafile is the data source of the databasae 

    Dim datafile As String
    Dim connstring As String

    'myConnection is an instance of the object oledbconnection 
    Dim myConnection As OleDbConnection = New OleDbConnection

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            TextBox1.PasswordChar = "*"

        Else
            TextBox1.PasswordChar = ""
        End If
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'This section of code prompts the user for and administrative password when the Create New User Button is Clicked.
        provider = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source="
        datafile = "D:\POSSA\POSSA\POSSA\POSSA.accdb"
        connstring = provider & datafile
        myConnection.ConnectionString = connstring

        'connection is open 
        myConnection.Open()
        Dim ds As New POSDataSet
        Dim cmd As OleDbCommand = New OleDbCommand("Select * from [ADMIN] where [PASSWORD] = '" & TextBox1.Text & "'", myConnection)

        Dim dr As OleDbDataReader = cmd.ExecuteReader

        'the following variables will hold the user first and last name if found.
        Dim userFound As Boolean = False
        Dim password As String = ""

        'if found
        While dr.Read
            userFound = True
            password = dr("PASSWORD").ToString
        End While

        If userFound = True Then
            Me.Hide()
            TextBox1.Clear()
            CreateUser.Show()
        Else
            MsgBox("SORRY, PASSWORD IS INCORRECT", MsgBoxStyle.OkOnly, "INVALID LOGIC, TRY AGAIN")
            TextBox1.Clear()
            Me.TextBox1.Select()
            myConnection.Close()
        End If


    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Application.Restart()
    End Sub

    Private Sub InputBox_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.TextBox1.Select()
    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        If Asc(e.KeyChar) < 65 And Asc(e.KeyChar) > 90 Or
       Asc(e.KeyChar) > 97 And Asc(e.KeyChar) < 122 Then
            e.Handled = True
            MessageBox.Show("YOU CAN ONLY INPUT UPPERCASE LETTERS")
        End If
    End Sub
End Class