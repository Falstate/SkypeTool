Imports SKYPE4COMLib

Public Class Form3

    Dim Skype As New SKYPE4COMLib.Skype
    Private Sub Form2_Load(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load


    End Sub


    Private Sub LogInButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LogInButton2.Click

        Try

            Skype.Attach()
            MessageBox.Show("Connected Successfully!", "Connected", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            MessageBox.Show("Unable To Connect!", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    Private Sub LogInButton4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LogInButton4.Click
        Skype.CurrentUserProfile.MoodText = TextBox11.Text 'Mood Change
    End Sub

    Private Sub LogInButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LogInButton3.Click
        Skype.CurrentUserProfile.FullName = TextBox12.Text 'Name Change
    End Sub

    Private Sub LogInButton5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LogInButton5.Click
        Skype.CurrentUserProfile.About = TextBox10.Text 'About Tab
    End Sub

    Private Sub LogInButton6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LogInButton6.Click
        Skype.CurrentUserProfile.PhoneMobile = TextBox9.Text 'Phone Number
    End Sub

    Private Sub LogInButton7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LogInButton7.Click
        Skype.CurrentUserProfile.Province = TextBox7.Text 'Province
    End Sub

    Private Sub LogInButton8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LogInButton8.Click
        Skype.CurrentUserProfile.City = TextBox7.Text 'City Change
    End Sub

    Private Sub LogInButton9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LogInButton9.Click
        Skype.CurrentUserProfile.Homepage = TextBox13.Text 'Homepage
    End Sub

    Private Sub LogInButton17_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button17.Click

        For Each user As SKYPE4COMLib.User In Skype.Friends

            Skype.SendMessage(user.Handle, TextBox1.Text)

        Next

    End Sub

    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        this.oSkype.CurrentUserStatus = TUserStatus.cusOnline
    End Sub

    Private Function this() As Object
        Throw New NotImplementedException
    End Function

End Class
