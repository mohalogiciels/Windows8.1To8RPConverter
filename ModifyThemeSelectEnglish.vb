Imports System.IO
Imports System.Xml

Public Class ModifyThemeSelectEnglish
    Dim ProgramFiles = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles)

    Private Sub ModifyThemeSelectEnglish_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Check installed theme address bar style
        If ModifyWizard.AddressBarStyle = "blue" Then
            BlueAddBarRadioButton.Enabled = False
        ElseIf ModifyWizard.AddressBarStyle = "white" Then
            WhiteAddBarRadioButton.Enabled = False
        End If

        ' Load checked value
        If ModifyWizard.ChangeToSelected = "blue" Then
            BlueAddBarRadioButton.Checked = True
        ElseIf ModifyWizard.ChangeToSelected = "white" Then
            WhiteAddBarRadioButton.Checked = True
        ElseIf ModifyWizard.ChangeToSelected = "none" Then
            NoChangeStyleRadioButton.Checked = True
        End If
    End Sub

    Private Sub BlueAddBarRadioButton_CheckedChanged(sender As Object, e As EventArgs) Handles BlueAddBarRadioButton.CheckedChanged
        NextButton.Enabled = True
        ModifyWizard.ChangeAddBarStyle = True
        ModifyWizard.ChangeToSelected = "blue"
        ModifyWizard.IsThemeChangeNeeded = True
    End Sub

    Private Sub WhiteAddBarRadioButton_CheckedChanged(sender As Object, e As EventArgs) Handles WhiteAddBarRadioButton.CheckedChanged
        NextButton.Enabled = True
        ModifyWizard.ChangeAddBarStyle = True
        ModifyWizard.ChangeToSelected = "white"
        ModifyWizard.IsThemeChangeNeeded = True
    End Sub

    Private Sub NoChangeRadioButton_CheckedChanged(sender As Object, e As EventArgs) Handles NoChangeStyleRadioButton.CheckedChanged
        NextButton.Enabled = True
        ModifyWizard.ChangeAddBarStyle = False
        ModifyWizard.ChangeToSelected = "none"
        ModifyWizard.IsThemeChangeNeeded = False
    End Sub

    Private Sub NextButton_Click(sender As Object, e As EventArgs) Handles NextButton.Click
        Me.Dispose()
        ModifyOptionsEnglish.Show()
    End Sub

    Private Sub CloseButton_Click(sender As Object, e As EventArgs) Handles CloseButton.Click
        Me.Close()
    End Sub

    Private Sub ModifyThemeSelectEnglish_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        If e.CloseReason = CloseReason.UserClosing Then
            If MessageBox.Show("Are you sure you want to cancel? This program will be closed.", "Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) = System.Windows.Forms.DialogResult.Yes Then
                ModifyWizard.Close()
            ElseIf System.Windows.Forms.DialogResult.No Then
                e.Cancel = True
            End If
        End If
    End Sub
End Class