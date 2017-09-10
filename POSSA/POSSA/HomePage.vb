Imports System.Data.OleDb
Imports System.Runtime.InteropServices

Public Class HomePage


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

    'datafile is the data source of the databasae 

    Dim datafile As String
    Dim connstring As String

    'myConnection is an instance of the object oledbconnection 
    Dim myConnection As OleDbConnection = New OleDbConnection


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        provider = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source="
        datafile = "D:\POSSA\POSSA\POSSA\POSSA.accdb"
        connstring = provider & datafile
        myConnection.ConnectionString = connstring

        'connection is open 
        myConnection.Open()
        Dim ds As New POSDataSet
        Dim cmd As OleDbCommand = New OleDbCommand("Select * from [CREATE_USER] where [USERNAME] = '" & TextBox1.Text & "' AND [PASSWORD] = '" & TextBox2.Text & "'", myConnection)

        Dim dr As OleDbDataReader = cmd.ExecuteReader


        'the following variables will hold the user first and last name if found.
        Dim userFound As Boolean = False
        Dim username As String = ""
        Dim password As String = ""

        'if found
        While dr.Read
            userFound = True

            username = dr("USERNAME").ToString
            password = dr("PASSWORD").ToString
        End While

        If userFound = True Then
            Me.Hide()
            TextBox1.Clear()
            TextBox2.Clear()
            RegistrationForm.Show()

            'this line of code will display a welcome plus the first and last name of the user 
            MsgBox("WELCOME " & username)
        Else
            MsgBox("SORRY, USERNAME OR PASSWORD NOT FOUND", MsgBoxStyle.OkOnly, "INVALID LOGIN, TRY AGAIN")
            TextBox1.Clear()
            TextBox2.Clear()
            Me.TextBox1.Select()
            myConnection.Close()
        End If
    End Sub


    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Application.Exit()
        'Me.Close()
        'InputBox.Close()
        'CreateUser.Close()
        'RegistrationForm.Close()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Hide()
        InputBox.Show()
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
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

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Label4.Text = Mid(Label4.Text, 2) & Mid(Label4.Text, 1, 1)
    End Sub

    Private Sub HomePage_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Timer1.Enabled = True

        Dim hMenu As IntPtr = GetSystemMenu(Me.Handle, False)
        Dim num_menu_items As Integer = GetMenuItemCount(hMenu)

        'For i As Integer = num_menu_items - 1 To 0 Step -1
        '    RemoveMenu(hMenu, i, MF_BYPOSITION)
        'Next i
        RemoveMenu(hMenu, num_menu_items - 1, MF_BYPOSITION)
        RemoveMenu(hMenu, num_menu_items - 2, MF_BYPOSITION)
    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        ABOUT.Show()
    End Sub
End Class
