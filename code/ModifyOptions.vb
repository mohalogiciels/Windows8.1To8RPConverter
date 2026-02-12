Imports System.Xml

Public Class ModifyOptions
    Dim ProgramFiles As String = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles)
    Dim ProgramFilesX86 As String = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86)
    Dim RestartCount As Integer

    Private Sub ModifyOptionsEnglish_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        InitialiseOptions()
    End Sub

    Private Sub InitialiseOptions()
        If ModifyWizard.InstONE = True And FileIO.FileSystem.DirectoryExists(ProgramFiles & "\OldNewExplorer") Then
            ChangeONECheckBox.Text = "Uninstall OldNewExplorer"
            ModifyWizard.ModifyInstallONE = False
        ElseIf ModifyWizard.InstONE = False Or Not FileIO.FileSystem.DirectoryExists(ProgramFiles & "\OldNewExplorer") Then
            ChangeONECheckBox.Text = "Install OldNewExplorer"
            ModifyWizard.ModifyInstallONE = True
        End If

        If ModifyWizard.InstUX = "UltraUX" And (FileIO.FileSystem.DirectoryExists(ProgramFiles & "\UltraUXThemePatcher") Or FileIO.FileSystem.DirectoryExists(ProgramFilesX86 & "\UltraUXThemePatcher")) Then
            ChangeUXPatchComboBox.SelectedItem = "UXThemeSignatureBypass by Big Muscle"
            ChangeUXPatchComboBox.Items.Remove("UltraUXThemePatcher by Manuel Hoefs")
            ModifyWizard.ModifyChangeToUXMethod = "UXTSB"
        ElseIf ModifyWizard.InstUX = "UXTSB" And FileIO.FileSystem.FileExists(Environment.GetFolderPath(Environment.SpecialFolder.Windows) & "\System32\UxThemeSignatureBypass.dll") Then
            ChangeUXPatchComboBox.SelectedItem = "UltraUXThemePatcher by Manuel Hoefs"
            ChangeUXPatchComboBox.Items.Remove("UXThemeSignatureBypass by Big Muscle")
            ModifyWizard.ModifyChangeToUXMethod = "UltraUX"
        ElseIf ModifyWizard.InstUX = "NoUXPatch" Then
            ChangeUXPatchComboBox.Items.Remove("UltraUXThemePatcher by Manuel Hoefs")
            ChangeUXPatchComboBox.Items.Remove("UXThemeSignatureBypass by Big Muscle")
            ChangeUXCheckBox.Enabled = False
            ChangeUXCheckBox.Checked = False
            ChangeUXPatchComboBox.Enabled = False
        Else
            MessageBox.Show("Some installed files could not be found on this system!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            ChangeUXPatchComboBox.Items.Remove("UltraUXThemePatcher by Manuel Hoefs")
            ChangeUXPatchComboBox.Items.Remove("UXThemeSignatureBypass by Big Muscle")
            ChangeUXCheckBox.Enabled = False
            ChangeUXPatchComboBox.Enabled = False
            ChangeUXWarningLabel.Enabled = True
            ChangeUXWarningLabel.Visible = True
        End If

        If ModifyWizard.ChangeAddBarStyle = True And ModifyWizard.ChangeToSelectedStyle = "WhiteAddressBar" And ModifyWizard.InstQuero = True Then
            ChangeQueroCheckBox.Checked = True
            ChangeQueroCheckBox.Enabled = False
        ElseIf (ModifyWizard.ChangeAddBarStyle = True And ModifyWizard.ChangeToSelectedStyle = "WhiteAddressBar" And ModifyWizard.InstQuero = False) Or (ModifyWizard.ChangeAddBarStyle = False And ModifyWizard.AddressBarStyle = "WhiteAddressBar") Then
            ChangeQueroCheckBox.Enabled = False
        ElseIf ModifyWizard.ChangeToSelectedStyle = "BlueAddressBar" Or (ModifyWizard.AddressBarStyle = "BlueAddressBar" And ModifyWizard.ChangeAddBarStyle = False) Then
            ChangeQueroCheckBox.Enabled = True
        End If

        If ModifyWizard.InstQuero = True And FileIO.FileSystem.DirectoryExists(ProgramFiles & "\Quero Toolbar") Then
            ChangeQueroCheckBox.Text = "Uninstall Quero Toolbar"
            ModifyWizard.ModifyInstallQuero = False
        ElseIf ModifyWizard.InstQuero = False Or Not FileIO.FileSystem.DirectoryExists(ProgramFiles & "\Quero Toolbar") Then
            ChangeQueroCheckBox.Text = "Install Quero Toolbar"
            ModifyWizard.ModifyInstallQuero = True
        End If

        If ModifyWizard.IsSelectedLanguageDisplayLanguage = True Then
            ChangeSysFilesCheckBox.Enabled = True
        Else
            ChangeSysFilesCheckBox.Enabled = False
        End If

        If ModifyWizard.InstSysFiles = True Then
            ChangeSysFilesCheckBox.Text = "Restore UIRibbon.dll && UIRibbonRes.dll"
            ModifyWizard.ModifyInstallSysFiles = False
        Else
            ChangeSysFilesCheckBox.Text = "Replace UIRibbon.dll && UIRibbonRes.dll"
            ModifyWizard.ModifyInstallSysFiles = True
        End If

        If ModifyWizard.InstSounds = True Then
            InstallSoundsCheckBox.Enabled = False
        End If

        If ModifyWizard.InstGadgets = True And FileIO.FileSystem.FileExists(ProgramFiles & "\Windows Sidebar\8GadgetPack.exe") Then
            ChangeGadgetsCheckBox.Text = "Uninstall 8GadgetPack"
            ModifyWizard.ModifyInstallGadgets = False
        ElseIf ModifyWizard.InstGadgets = False Or Not FileIO.FileSystem.FileExists(ProgramFiles & "\Windows Sidebar\8GadgetPack.exe") Then
            ChangeGadgetsCheckBox.Text = "Install 8GadgetPack"
            ModifyWizard.ModifyInstallGadgets = True
        End If

        If ModifyWizard.Inst7TaskTw = True And FileIO.FileSystem.DirectoryExists(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) & "\Programs\7+ Taskbar Tweaker") Then
            Change7TaskTwCheckBox.Text = "Uninstall 7+ Taskbar Tweaker"
            ModifyWizard.ModifyInstall7TaskTw = False
        ElseIf ModifyWizard.Inst7TaskTw = False Or Not FileIO.FileSystem.DirectoryExists(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) & "\Programs\7+ Taskbar Tweaker") Then
            Change7TaskTwCheckBox.Text = "Install 7+ Taskbar Tweaker"
            ModifyWizard.ModifyInstall7TaskTw = True
        End If
    End Sub

    Private Sub NextButton_Click(sender As Object, e As EventArgs) Handles NextButton.Click
        If MessageBox.Show("Are you sure you want to proceed?", "Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) = System.Windows.Forms.DialogResult.Yes Then
            ' Change values (settings) of programs.xml file in program folder
            If ModifyWizard.ChangeAddBarStyle = False And ChangeONECheckBox.Checked = False And ChangeUXCheckBox.Checked = False And ChangeQueroCheckBox.Checked = False And ChangeSysFilesCheckBox.Checked = False And InstallSoundsCheckBox.Checked = False And ChangeGadgetsCheckBox.Checked = False And Change7TaskTwCheckBox.Checked = False Then
                MessageBox.Show("You have selected nothing to change. Please choose something to change, or close this program.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Else
                If ModifyWizard.ChangeAddBarStyle = True Then
                    If ModifyWizard.ChangeToSelectedStyle = "BlueAddressBar" Then
                        ModifyWizard.ConfigFile.SelectSingleNode("//setting[@name='AddressBarStyle']/value").InnerText = "BlueAddressBar"
                    ElseIf ModifyWizard.ChangeToSelectedStyle = "WhiteAddressBar" Then
                        ModifyWizard.ConfigFile.SelectSingleNode("//setting[@name='AddressBarStyle']/value").InnerText = "WhiteAddressBar"
                    End If
                End If

                If ChangeONECheckBox.Checked = True Then
                    If ModifyWizard.ModifyInstallONE = True Then
                        ModifyWizard.ConfigFile.SelectSingleNode("//setting[@name='InstONE']/value").InnerText = True
                    Else
                        ModifyWizard.ConfigFile.SelectSingleNode("//setting[@name='InstONE']/value").InnerText = False
                    End If
                End If

                If ChangeUXCheckBox.Checked = True Then
                    If ModifyWizard.ModifyChangeToUXMethod = "UXTSB" Then
                        ModifyWizard.ConfigFile.SelectSingleNode("//setting[@name='InstUX']/value").InnerText = "UXTSB"
                    ElseIf ModifyWizard.ModifyChangeToUXMethod = "UltraUX" Then
                        ModifyWizard.ConfigFile.SelectSingleNode("//setting[@name='InstUX']/value").InnerText = "UltraUX"
                    End If
                End If

                If ChangeQueroCheckBox.Checked = True Then
                    If ModifyWizard.ModifyInstallQuero = True Then
                        ModifyWizard.ConfigFile.SelectSingleNode("//setting[@name='InstQuero']/value").InnerText = True
                    Else
                        ModifyWizard.ConfigFile.SelectSingleNode("//setting[@name='InstQuero']/value").InnerText = False
                    End If
                End If

                If ChangeSysFilesCheckBox.Checked = True Then
                    If ModifyWizard.ModifyInstallSysFiles = True Then
                        ModifyWizard.ConfigFile.SelectSingleNode("//setting[@name='InstSysFiles']/value").InnerText = True
                    Else
                        ModifyWizard.ConfigFile.SelectSingleNode("//setting[@name='InstSysFiles']/value").InnerText = False
                    End If
                End If

                If InstallSoundsCheckBox.Checked = True Then
                    ModifyWizard.ConfigFile.SelectSingleNode("//setting[@name='InstSounds']/value").InnerText = True
                End If

                If ChangeGadgetsCheckBox.Checked = True Then
                    If ModifyWizard.ModifyInstallGadgets = True Then
                        ModifyWizard.ConfigFile.SelectSingleNode("//setting[@name='InstGadgets']/value").InnerText = True
                    Else
                        ModifyWizard.ConfigFile.SelectSingleNode("//setting[@name='InstGadgets']/value").InnerText = False
                    End If
                End If

                If Change7TaskTwCheckBox.Checked = True Then
                    If ModifyWizard.ModifyInstall7TaskTw = True Then
                        ModifyWizard.ConfigFile.SelectSingleNode("//setting[@name='Inst7TaskTw']/value").InnerText = True
                    Else
                        ModifyWizard.ConfigFile.SelectSingleNode("//setting[@name='Inst7TaskTw']/value").InnerText = False
                    End If
                End If

                ' Finally save changes to config.xml file
                ModifyWizard.ConfigFile.Save(ModifyWizard.ProgDir & "\config.xml")

                ' Load next screen
                Me.Dispose()
                ModifyModify.Show()
            End If
        End If
    End Sub

    Private Sub BackButton_Click(sender As Object, e As EventArgs) Handles BackButton.Click
        Me.Dispose()
        ModifyThemeSelect.Show()
    End Sub

    Private Sub CloseButton_Click(sender As Object, e As EventArgs) Handles CloseButton.Click
        Me.Close()
    End Sub

    Private Sub ChangeONECheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles ChangeONECheckBox.CheckedChanged
        If ChangeONECheckBox.Checked = True Then
            ModifyWizard.ChangeONEChecked = True
        Else
            ModifyWizard.ChangeONEChecked = False
        End If
    End Sub

    Private Sub ChangeUXCheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles ChangeUXCheckBox.CheckedChanged
        If ChangeUXCheckBox.Checked = True Then
            ModifyWizard.ChangeUXMethodChecked = True
            ModifyWizard.IsRestartNeeded = True
            RestartCount = RestartCount + 1
            ChangeUXPatchComboBox.Enabled = True
        Else
            ModifyWizard.ChangeUXMethodChecked = False
            RestartCount = RestartCount - 1
            If RestartCount < 1 Then
                ModifyWizard.IsRestartNeeded = False
            End If
            ChangeUXPatchComboBox.Enabled = False
        End If
    End Sub

    Private Sub ChangeUXWarningLabel_Click(sender As Object, e As EventArgs) Handles ChangeUXWarningLabel.Click
        MessageBox.Show("Some installed files could not be found on this system!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
    End Sub

    Private Sub ChangeQueroCheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles ChangeQueroCheckBox.CheckedChanged
        If ChangeQueroCheckBox.Checked = True Then
            ModifyWizard.ChangeQueroChecked = True
        Else
            ModifyWizard.ChangeQueroChecked = False
        End If
    End Sub

    Private Sub ChangeSysFilesCheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles ChangeSysFilesCheckBox.CheckedChanged
        If ChangeSysFilesCheckBox.Checked = True Then
            ModifyWizard.ChangeSysFilesChecked = True
            ModifyWizard.IsRestartNeeded = True
            RestartCount = RestartCount + 1
        Else
            ModifyWizard.ChangeSysFilesChecked = False
            RestartCount = RestartCount - 1
            If RestartCount < 1 Then
                ModifyWizard.IsRestartNeeded = False
            End If
        End If
    End Sub

    Private Sub InstallSoundsCheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles InstallSoundsCheckBox.CheckedChanged
        If InstallSoundsCheckBox.Checked = True Then
            ModifyWizard.InstallSoundsChecked = True
            ModifyWizard.IsThemeChangeNeeded = True
        Else
            ModifyWizard.InstallSoundsChecked = False
            If ModifyWizard.ChangeAddBarStyle = False Then
                ModifyWizard.IsThemeChangeNeeded = False
            End If
        End If
    End Sub

    Private Sub ChangeGadgetsCheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles ChangeGadgetsCheckBox.CheckedChanged
        If ChangeGadgetsCheckBox.Checked = True Then
            ModifyWizard.ChangeGadgetsChecked = True
        Else
            ModifyWizard.ChangeGadgetsChecked = False
        End If
    End Sub

    Private Sub Change7TaskTwCheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Change7TaskTwCheckBox.CheckedChanged
        If Change7TaskTwCheckBox.Checked = True Then
            ModifyWizard.Change7TaskTwChecked = True
        Else
            ModifyWizard.Change7TaskTwChecked = False
        End If
    End Sub

    Private Sub SelectAllButton_Click(sender As Object, e As EventArgs) Handles SelectAllButton.Click
        ChangeONECheckBox.Checked = True
        ChangeUXCheckBox.Checked = True
        If ChangeQueroCheckBox.Enabled = True Then
            ChangeQueroCheckBox.Checked = True
        End If
        If ChangeSysFilesCheckBox.Enabled = True Then
            ChangeSysFilesCheckBox.Checked = True
        End If
        If InstallSoundsCheckBox.Enabled = True Then
            InstallSoundsCheckBox.Checked = True
        End If
        ChangeGadgetsCheckBox.Checked = True
        Change7TaskTwCheckBox.Checked = True
    End Sub

    Private Sub ClearAllButton_Click(sender As Object, e As EventArgs) Handles ClearAllButton.Click
        ChangeONECheckBox.Checked = False
        ChangeUXCheckBox.Checked = False
        If ChangeQueroCheckBox.Enabled = True Then
            ChangeQueroCheckBox.Checked = False
        End If
        If ChangeSysFilesCheckBox.Enabled = True Then
            ChangeSysFilesCheckBox.Checked = False
        End If
        If InstallSoundsCheckBox.Enabled = True Then
            InstallSoundsCheckBox.Checked = False
        End If
        ChangeGadgetsCheckBox.Checked = False
        Change7TaskTwCheckBox.Checked = False
    End Sub

    Private Sub ModifyOptionsEnglish_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        If e.CloseReason = CloseReason.UserClosing Then
            If MessageBox.Show("Are you sure you want to cancel? This program will be closed.", "Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) = System.Windows.Forms.DialogResult.Yes Then
                ModifyWizard.Close()
            ElseIf System.Windows.Forms.DialogResult.No Then
                e.Cancel = True
            End If
        End If
    End Sub
End Class