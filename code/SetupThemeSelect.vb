Imports System.Globalization

Public Class SetupThemeSelect

    Private Sub SetupThemeSelect_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' If already chosen style, load setting
        If SetupWizard.AddressBarStyle = "BlueAddressBar" Then
            BlueAddBarRadioButton.Checked = True
        ElseIf SetupWizard.AddressBarStyle = "WhiteAddressBar" Then
            WhiteAddBarRadioButton.Checked = True
        End If
    End Sub

    Private Sub BlueAddBarRadioButton_CheckedChanged(sender As Object, e As EventArgs) Handles BlueAddBarRadioButton.CheckedChanged
        NextButton.Enabled = True
        SetupWizard.AddressBarStyle = "BlueAddressBar"
    End Sub

    Private Sub WhiteAddBarRadioButton_CheckedChanged(sender As Object, e As EventArgs) Handles WhiteAddBarRadioButton.CheckedChanged
        NextButton.Enabled = True
        SetupWizard.AddressBarStyle = "WhiteAddressBar"
    End Sub

    Private Sub NextButton_Click(sender As Object, e As EventArgs) Handles NextButton.Click
        Me.Dispose()
        SetupOptions.Show()
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

    Private Sub SetupThemeSelect_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        If e.CloseReason = CloseReason.UserClosing Then
            If MessageBox.Show("Are you sure you want to cancel? This program will be closed.", "Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) = System.Windows.Forms.DialogResult.Yes Then
                SetupWizard.Close()
            ElseIf System.Windows.Forms.DialogResult.No Then
                e.Cancel = True
            End If
        End If
    End Sub
End Class