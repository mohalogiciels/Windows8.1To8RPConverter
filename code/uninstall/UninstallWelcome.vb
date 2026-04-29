Public Class UninstallWelcome

    Private Sub NextButton_Click(sender As Object, e As EventArgs) Handles NextButton.Click
        Me.Dispose()
        UninstallOptions.Show()
    End Sub

    Private Sub CloseButton_Click(sender As Object, e As EventArgs) Handles CloseButton.Click
        Application.Exit()
    End Sub

    Private Sub UninstallWelcome_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        If MessageBox.Show("Are you sure you want to cancel? The uninstaller will be closed.", "Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) = System.Windows.Forms.DialogResult.Yes Then
            Application.Exit()
        ElseIf System.Windows.Forms.DialogResult.No Then
            e.Cancel = True
        End If
    End Sub
End Class