Public Class ABOUT

    Private Sub ABOUT_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Set the title of the form.
        Dim ApplicationTitle As String
        If My.Application.Info.Title <> "" Then
            ApplicationTitle = My.Application.Info.Title
        Else
            ApplicationTitle = System.IO.Path.GetFileNameWithoutExtension(My.Application.Info.AssemblyName)
        End If
        Me.Text = String.Format("About {0}", ApplicationTitle)
        ' Initialize all of the text displayed on the About Box.
        ' TODO: Customize the application's assembly information in the "Application" pane of the project 
        '    properties dialog (under the "Project" menu).
        Me.LabelProductName.Text = "POSSA"
        Me.LabelVersion.Text = String.Format("Version {0}", My.Application.Info.Version.ToString)
        Me.LabelCopyright.Text = My.Application.Info.Copyright
        Me.LabelCompanyName.Text = "Asirifi Charles (BSc Computer Science) && Sowah Richmond Doe (Bsc Statistics)"
        Me.TextBoxDescription.Text = "This is a local App designed for the Political Science Department UG, Legon. It's a Departmental Registration App. It allows User to register students and display the sucessful registrations in a Grid View. The User can import the sucessful registrations into Microsoft Excel for further analysis."
    End Sub

    Private Sub OKButton_Click(sender As Object, e As EventArgs) Handles OKButton.Click
        Me.Close()
    End Sub
End Class