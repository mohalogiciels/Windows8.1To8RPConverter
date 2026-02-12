Public Class ContactMe

    Private Sub OKButton_Click(sender As Object, e As EventArgs) Handles OKButton.Click
        Me.Dispose()
    End Sub

    Private Sub GitHubRepoLinkLabel_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles GitHubRepoLinkLabel.LinkClicked
        Process.Start("https://github.com/mohalogiciels/Windows8.1To8RPConverter")
    End Sub

    Private Sub EmailLinkLabel_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles EmailLinkLabel.LinkClicked
        Process.Start("mailto:mohalogiciels@hotmail.com")
    End Sub
End Class