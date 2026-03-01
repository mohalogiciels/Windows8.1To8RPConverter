Imports Microsoft.Win32

Public Class UninstallOptionsEnglish
    Dim ProgramFiles As String = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles)

    Private Sub UninstallOptionsEnglish_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown
        ' Check if programs have been installed
        If UninstallWizard.InstAeroGlass = True Then
            RemoveAeroGlassCheckBox.Enabled = True
        End If

        If UninstallWizard.InstONE = True Then
            RemoveONECheckBox.Enabled = True
        End If

        If UninstallWizard.InstUX <> "none" Then
            RemoveUXThemePatchCheckBox.Enabled = True
        End If

        If UninstallWizard.InstQuero = True Then
            RemoveQueroCheckBox.Enabled = True
        End If

        If UninstallWizard.InstSysFiles = True Then
            RestoreSysFilesCheckBox.Enabled = True
        End If

        If UninstallWizard.InstSounds = True Then
            RemoveSoundsCheckBox.Enabled = True
        End If

        If UninstallWizard.InstGadgets = True Then
            RemoveGadgetsCheckBox.Enabled = True
        End If

        ' Choose remove all as standard
        RemoveEverythingRadioButton.Checked = True
    End Sub

    Private Sub RemoveEverythingRadioButton_CheckedChanged(sender As Object, e As EventArgs) Handles RemoveEverythingRadioButton.CheckedChanged
        ' Disable and check all checkboxes when chosen remove all
        ProgramListGroupBox.Visible = False
        InfoRemoveOnlyThemeRadioButton.Enabled = False

        If RemoveAeroGlassCheckBox.Enabled = True Then
            RemoveAeroGlassCheckBox.Checked = True
        Else

            RemoveAeroGlassCheckBox.Checked = False
        End If

        If RemoveONECheckBox.Enabled = True Then
            RemoveONECheckBox.Checked = True
        Else
            RemoveONECheckBox.Checked = False
        End If

        If RemoveUXThemePatchCheckBox.Enabled = True Then
            RemoveUXThemePatchCheckBox.Checked = True
        Else
            RemoveUXThemePatchCheckBox.Checked = False
        End If

        If RemoveQueroCheckBox.Enabled = True Then
            RemoveQueroCheckBox.Checked = True
        Else
            RemoveQueroCheckBox.Checked = False
        End If

        If RestoreSysFilesCheckBox.Enabled = True Then
            RestoreSysFilesCheckBox.Checked = True
        Else
            RestoreSysFilesCheckBox.Checked = False
        End If

        If RemoveSoundsCheckBox.Enabled = True Then
            RemoveSoundsCheckBox.Checked = True
        Else
            RemoveSoundsCheckBox.Checked = False
        End If

        If RemoveGadgetsCheckBox.Enabled = True Then
            RemoveGadgetsCheckBox.Checked = True
        Else
            RemoveGadgetsCheckBox.Checked = False
        End If
    End Sub

    Private Sub RemoveOnlyThemeRadioButton_CheckedChanged(sender As Object, e As EventArgs) Handles RemoveOnlyThemeRadioButton.CheckedChanged
        ' Disable and uncheck all checkboxes when chosen remove only theme
        ProgramListGroupBox.Visible = False
        InfoRemoveOnlyThemeRadioButton.Enabled = True
        RemoveAeroGlassCheckBox.Checked = False
        RemoveONECheckBox.Checked = False
        RemoveUXThemePatchCheckBox.Checked = False
        RemoveQueroCheckBox.Checked = False
        RestoreSysFilesCheckBox.Checked = False
        RemoveSoundsCheckBox.Checked = False
        RemoveGadgetsCheckBox.Checked = False
    End Sub

    Private Sub InfoRemoveOnlyThemeRadioButton_Click(sender As Object, e As EventArgs) Handles InfoRemoveOnlyThemeRadioButton.Click
        UninstallOptionsInfoRemoveOnlyThemeEnglish.ShowDialog()
    End Sub

    Private Sub RemoveChooseRadioButton_CheckedChanged(sender As Object, e As EventArgs) Handles RemoveChooseRadioButton.CheckedChanged
        ProgramListGroupBox.Visible = True
        InfoRemoveOnlyThemeRadioButton.Enabled = False
    End Sub

    Private Sub RemoveONECheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles RemoveONECheckBox.CheckedChanged
        If RemoveONECheckBox.Checked = False Then
            MessageBox.Show("Restoring the system files later manually needs extended knowledge of modification of system files, so choose it now to remove it if you are unsure of doing it.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End If
    End Sub

    Private Sub NextButton_Click(sender As Object, e As EventArgs) Handles NextButton.Click
        If MessageBox.Show("Are you sure you want to proceed?", "Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) = System.Windows.Forms.DialogResult.Yes Then
            If RemoveAeroGlassCheckBox.Checked = True Then
                UninstallWizard.UninstAero = True
            Else
                UninstallWizard.UninstAero = False
            End If

            If RemoveONECheckBox.Checked = True Then
                UninstallWizard.UninstONE = True
            Else
                UninstallWizard.UninstONE = False
            End If

            If RemoveUXThemePatchCheckBox.Checked = True Then
                UninstallWizard.UninstUX = True
            Else
                UninstallWizard.UninstUX = False
            End If

            If RemoveQueroCheckBox.Checked = True Then
                UninstallWizard.UninstQuero = True
            Else
                UninstallWizard.UninstQuero = False
            End If

            If RestoreSysFilesCheckBox.Checked = True Then
                UninstallWizard.UninstSysFiles = True
            Else
                UninstallWizard.UninstSysFiles = False
            End If

            If RemoveSoundsCheckBox.Checked = True Then
                UninstallWizard.UninstSounds = True
            Else
                UninstallWizard.UninstSounds = False
            End If

            If RemoveGadgetsCheckBox.Checked = True Then
                UninstallWizard.UninstGadgets = True
            Else
                UninstallWizard.UninstGadgets = False
            End If

            Me.Dispose()
            UninstallUninstallEnglish.Show()
        End If
    End Sub

    Private Sub CloseButton_Click(sender As Object, e As EventArgs) Handles CloseButton.Click
        Me.Close()
    End Sub

    Private Sub UninstallOptionsEnglish_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        If e.CloseReason = CloseReason.UserClosing Then
            If MessageBox.Show("Are you sure you want to cancel? The uninstaller will be closed.", "Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) = System.Windows.Forms.DialogResult.Yes Then
                UninstallWizard.Close()
            ElseIf System.Windows.Forms.DialogResult.No Then
                e.Cancel = True
            End If
        End If
    End Sub
End Class
