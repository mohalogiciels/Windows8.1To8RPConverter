﻿Public Class SetupToolStripInfo

    Private Sub SetupToolStripInfo_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles ThemeSourceLinkLabel.LinkClicked
        Process.Start("https://www.deviantart.com/yash12396/art/Windows-8-Release-Preview-VS-for-Windows-8-1-1-764774206")
    End Sub

    Private Sub OKButton_Click(sender As Object, e As EventArgs) Handles OKButton.Click
        Me.Dispose()
    End Sub
End Class