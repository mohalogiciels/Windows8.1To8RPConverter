Imports System.IO
Imports System.Xml

Public Class ModifyThemeSelect

    Private Sub ModifyThemeSelect_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Check installed theme address bar style
        If ModifyWizard.AddressBarStyle = "BlueAddressBar" Then
            BlueAddBarRadioButton.Enabled = False
        ElseIf ModifyWizard.AddressBarStyle = "WhiteAddressBar" Then
            WhiteAddBarRadioButton.Enabled = False
        End If
    End Sub

    Private Sub BlueAddBarRadioButton_CheckedChanged(sender As Object, e As EventArgs) Handles BlueAddBarRadioButton.CheckedChanged
        NextButton.Enabled = True
        ModifyWizard.ChangeAddBarStyle = True
        ModifyWizard.ChangeToSelectedStyle = "BlueAddressBar"
        ModifyWizard.IsThemeChangeNeeded = True
    End Sub

    Private Sub WhiteAddBarRadioButton_CheckedChanged(sender As Object, e As EventArgs) Handles WhiteAddBarRadioButton.CheckedChanged
        NextButton.Enabled = True
        ModifyWizard.ChangeAddBarStyle = True
        ModifyWizard.ChangeToSelectedStyle = "WhiteAddressBar"
        ModifyWizard.IsThemeChangeNeeded = True
    End Sub

    Private Sub NoChangeRadioButton_CheckedChanged(sender As Object, e As EventArgs) Handles NoChangeStyleRadioButton.CheckedChanged
        NextButton.Enabled = True
        ModifyWizard.ChangeAddBarStyle = False
        ModifyWizard.IsThemeChangeNeeded = False
    End Sub

    Private Sub NextButton_Click(sender As Object, e As EventArgs) Handles NextButton.Click
        Me.Dispose()
        ModifyOptions.Show()
    End Sub

    Private Sub CloseButton_Click(sender As Object, e As EventArgs) Handles CloseButton.Click
        Me.Close()
    End Sub

    Private Sub ModifyThemeSelect_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        If e.CloseReason = CloseReason.UserClosing Then
            If MessageBox.Show("Are you sure you want to cancel? This program will be closed.", "Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) = System.Windows.Forms.DialogResult.Yes Then
                ModifyWizard.Close()
            ElseIf System.Windows.Forms.DialogResult.No Then
                e.Cancel = True
            End If
        End If
    End Sub
End Class