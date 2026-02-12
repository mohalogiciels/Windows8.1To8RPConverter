Public Class SetupWelcome

    Private Sub NextButton_Click(sender As Object, e As EventArgs) Handles NextButton.Click
        Me.Dispose()
        SetupThemeSelect.Show()
    End Sub

    Private Sub CloseButton_Click(sender As Object, e As EventArgs) Handles CloseButton.Click
        Me.Close()
    End Sub

    Private Sub InfoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles InfoToolStripMenuItem.Click
        AboutProgram.ShowDialog()
    End Sub

    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()
    End Sub

    Private Sub ContactMeToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ContactMeToolStripMenuItem.Click
        ContactMe.ShowDialog()
    End Sub

    Private Sub SetupWelcome_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        If e.CloseReason = CloseReason.UserClosing Then
            If MessageBox.Show("Are you sure you want to cancel? This program will be closed.", "Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) = System.Windows.Forms.DialogResult.Yes Then
                SetupWizard.Close()
            ElseIf System.Windows.Forms.DialogResult.No Then
                e.Cancel = True
            End If
        End If
    End Sub
End Class