Imports System.Data.OleDb


Public Class CreateUser
    'provider is the technology used to connect to access database
    Dim provider As String
    'datafile is the data source of the database
    Dim datafile As String
    'connstring is the provider and the datafile
    Dim connstring As String
    'myConnection is an instance of the object oledbconnection
    Dim myConnection As OleDbConnection = New OleDbConnection

    Private Sub Button2_Click(sender As Object, e As EventArgs)
        If (TextBox1.Text = "" Or TextBox2.Text = "") Then
            MsgBox("Please Enter Your Username And Password In The Textboxes Below")
        Else
            provider = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source="
            datafile = "D:\POSSA\POSSA\POSSA\POSSA.accdb"
            connstring = provider & datafile
            myConnection.ConnectionString = connstring

            'connection is open 
            myConnection.Open()


            'Str stands for the variable for inserting data into the farmer table 
            Dim str As String
            str = "insert into CREATE_USER([USERNAME],[PASSWORD]) values(?,?)"
            Dim cmd As OleDbCommand = New OleDbCommand(str, myConnection)

            cmd.Parameters.Add(New OleDbParameter("USERNAME", CType(TextBox1.Text, String)))
            cmd.Parameters.Add(New OleDbParameter("PASSWORD", CType(TextBox2.Text, String)))

            Try
                cmd.ExecuteNonQuery()
                cmd.Dispose()
                myConnection.Close()
                TextBox1.Clear()
                TextBox2.Clear()

            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
            MsgBox("USER CREATION SUCESSFUL")
            Me.Close()
            HomePage.Show()

        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs)
        Me.Close()
        HomePage.Show()
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs)
        'when checkbox is ticked, the password is displayed in thhe textbox and unticked the password will be encrypted

        If CheckBox1.Checked = True Then
            TextBox2.PasswordChar = "*"
        Else
            TextBox2.PasswordChar = ""


        End If
    End Sub

    Private Sub CheckBox1_CheckedChanged_1(sender As Object, e As EventArgs)
        If CheckBox1.Checked = True Then
            TextBox2.PasswordChar = "*"

        Else
            TextBox2.PasswordChar = ""
        End If
    End Sub

    Private Sub CreateUser_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.TextBox1.Select()
    End Sub

    Private Sub Button2_Click_1(sender As Object, e As EventArgs) Handles Button2.Click
        If (TextBox1.Text = "" Or TextBox2.Text = "") Then
            MsgBox("PLEASE PROVIDE THE USERNAME AND PASSWORD YOU WANT TO CREATE AS A USER")
            Me.TextBox1.Select()
        Else
            provider = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source="
            datafile = "D:\POSSA\POSSA\POSSA\POSSA.accdb"
            connstring = provider & datafile
            myConnection.ConnectionString = connstring

            'connection is open 
            myConnection.Open()



            Dim str As String
            str = "insert into CREATE_USER([USERNAME],[PASSWORD]) values(?,?)"
            Dim cmd As OleDbCommand = New OleDbCommand(str, myConnection)

            cmd.Parameters.Add(New OleDbParameter("USERNAME", CType(TextBox1.Text, String)))
            cmd.Parameters.Add(New OleDbParameter("PASSWORD", CType(TextBox2.Text, String)))

            Try
                cmd.ExecuteNonQuery()
                cmd.Dispose()
                myConnection.Close()
                TextBox1.Clear()
                TextBox2.Clear()

            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
            MsgBox("USER CREATION SUCESSFUL")
            Me.Close()
            HomePage.Show()
            HomePage.TextBox1.Select()

        End If
    End Sub

    Private Sub Button3_Click_1(sender As Object, e As EventArgs) Handles Button3.Click
        Application.Restart()
    End Sub

    Private Sub CheckBox1_CheckedChanged_2(sender As Object, e As EventArgs)
        If CheckBox1.Checked = True Then
            TextBox2.PasswordChar = "*"

        Else
            TextBox2.PasswordChar = ""
        End If
    End Sub

    Private Sub CheckBox1_CheckedChanged_3(sender As Object, e As EventArgs)
        If CheckBox1.Checked = True Then
            TextBox2.PasswordChar = "*"

        Else
            TextBox2.PasswordChar = ""
        End If
    End Sub

    Private Sub CheckBox1_CheckedChanged_4(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            TextBox2.PasswordChar = "*"

        Else
            TextBox2.PasswordChar = ""
        End If
    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        If Asc(e.KeyChar) < 65 And Asc(e.KeyChar) > 90 Or
       Asc(e.KeyChar) > 97 And Asc(e.KeyChar) < 122 Then
            e.Handled = True
            MessageBox.Show("YOU CAN ONLY INPUT UPPERCASE LETTERS")
        End If
    End Sub

    Private Sub TextBox2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox2.KeyPress
        If Asc(e.KeyChar) < 65 And Asc(e.KeyChar) > 90 Or
       Asc(e.KeyChar) > 97 And Asc(e.KeyChar) < 122 Then
            e.Handled = True
            MessageBox.Show("YOU CAN ONLY INPUT UPPERCASE LETTERS")
        End If
    End Sub
End Class