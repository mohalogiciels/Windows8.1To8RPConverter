Imports System.Xml

Public Class ModifyOptionsEnglish
    Dim ProgramFiles = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles)
    Dim ProgramFilesX86 = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86)
    Dim WinDir = Environment.GetFolderPath(Environment.SpecialFolder.Windows)
    Dim RestartCount As Integer

    Private Sub ModifyOptionsEnglish_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If ModifyWizard.InstONE = True And FileIO.FileSystem.DirectoryExists(ProgramFiles & "\OldNewExplorer") Then
            ChangeONECheckBox.Text = "Uninstall OldNewExplorer"
        ElseIf ModifyWizard.InstONE = True And Not FileIO.FileSystem.DirectoryExists(ProgramFiles & "\OldNewExplorer") Then
            ChangeONECheckBox.Text = "Install OldNewExplorer"
        ElseIf ModifyWizard.InstONE = False Then
            ChangeONECheckBox.Text = "Install OldNewExplorer"
        End If

        If ModifyWizard.InstUX = "UltraUX" And (FileIO.FileSystem.DirectoryExists(ProgramFiles & "\UltraUXThemePatcher") Or FileIO.FileSystem.DirectoryExists(ProgramFilesX86 & "\UltraUXThemePatcher")) Then
            ChangeUXPatchComboBox.SelectedItem = "UXThemeSignatureBypass by Big Muscle"
            ChangeUXPatchComboBox.Items.Remove("UltraUXThemePatcher by Manuel Hoefs")
        ElseIf ModifyWizard.InstUX = "UXTSB" And FileIO.FileSystem.FileExists(FileIO.FileSystem.GetFiles(WinDir, FileIO.SearchOption.SearchTopLevelOnly, "UxThemeSignatureBypass*")(0)) Then
            ChangeUXPatchComboBox.SelectedItem = "UltraUXThemePatcher by Manuel Hoefs"
            ChangeUXPatchComboBox.Items.Remove("UXThemeSignatureBypass by Big Muscle")
        ElseIf ModifyWizard.InstUX = "none" Then
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
            ChangeUXCheckBox.Checked = False
            ChangeUXPatchComboBox.Enabled = False
            ChangeUXWarningLabel.Enabled = True
            ChangeUXWarningLabel.Visible = True
        End If

        If ModifyWizard.AddressBarStyle = "blue" And ModifyWizard.ChangeAddBarStyle = True And ModifyWizard.InstQuero = True Then
            ChangeQueroCheckBox.Checked = True
            ChangeQueroCheckBox.Enabled = False
        ElseIf ModifyWizard.AddressBarStyle = "blue" And ModifyWizard.ChangeAddBarStyle = True And ModifyWizard.InstQuero = False Then
            ChangeQueroCheckBox.Checked = False
            ChangeQueroCheckBox.Enabled = False
        ElseIf ModifyWizard.AddressBarStyle = "white" Or (ModifyWizard.AddressBarStyle = "blue" And ModifyWizard.ChangeAddBarStyle = False) Then
            ChangeQueroCheckBox.Enabled = True
            If ModifyWizard.ChangeQueroChecked = True Then
                ChangeQueroCheckBox.Checked = True
            ElseIf ModifyWizard.ChangeQueroChecked = False Then
                ChangeQueroCheckBox.Checked = False
            End If
        End If

        If ModifyWizard.InstQuero = True And FileIO.FileSystem.DirectoryExists(ProgramFiles & "\Quero Toolbar") Then
            ChangeQueroCheckBox.Text = "Uninstall Quero Toolbar"
        ElseIf ModifyWizard.InstQuero = True And Not FileIO.FileSystem.DirectoryExists(ProgramFiles & "\Quero Toolbar") Then
            ChangeQueroCheckBox.Text = "Install Quero Toolbar"
        ElseIf ModifyWizard.InstQuero = False Then
            ChangeQueroCheckBox.Text = "Install Quero Toolbar"
        End If

        If ModifyWizard.IsSelectedLanguageDisplayLanguage = False Then
            ChangeSysFilesCheckBox.Enabled = False
        ElseIf ModifyWizard.IsSelectedLanguageDisplayLanguage = True Then
            ChangeSysFilesCheckBox.Enabled = True
        End If

        If ModifyWizard.InstSysFiles = True Then
            ChangeSysFilesCheckBox.Text = "Restore UIRibbon.dll && UIRibbonRes.dll"
        ElseIf ModifyWizard.InstSysFiles = False Then
            ChangeSysFilesCheckBox.Text = "Replace UIRibbon.dll && UIRibbonRes.dll"
        End If

        If ModifyWizard.InstSounds = True Then
            InstallSoundsCheckBox.Checked = False
            InstallSoundsCheckBox.Enabled = False
        End If

        If ModifyWizard.InstGadgets = True And FileIO.FileSystem.FileExists(ProgramFiles & "\Windows Sidebar\8GadgetPack.exe") Then
            ChangeGadgetsCheckBox.Text = "Uninstall 8GadgetPack"
        ElseIf ModifyWizard.InstGadgets = True And Not FileIO.FileSystem.FileExists(ProgramFiles & "\Windows Sidebar\8GadgetPack.exe") Then
            ChangeGadgetsCheckBox.Text = "Install 8GadgetPack"
        ElseIf ModifyWizard.InstGadgets = False Then
            ChangeGadgetsCheckBox.Text = "Install 8GadgetPack"
        End If
    End Sub

    Private Sub ModifyOptionsEnglish_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown
        If ModifyWizard.ChangeONEChecked = True Then
            ChangeONECheckBox.Checked = True
        Else
            ChangeONECheckBox.Checked = False
        End If

        If ModifyWizard.ChangeUXMethodChecked = True Then
            ChangeUXCheckBox.Checked = True
        Else
            ChangeUXCheckBox.Checked = False
        End If

        If ModifyWizard.ChangeSysFilesChecked = True And ModifyWizard.IsSelectedLanguageDisplayLanguage = True Then
            ChangeSysFilesCheckBox.Checked = True
        ElseIf ModifyWizard.ChangeSysFilesChecked = False Then
            ChangeSysFilesCheckBox.Checked = False
        End If

        If ModifyWizard.ChangeGadgetsChecked = True Then
            ChangeGadgetsCheckBox.Checked = True
        Else
            ChangeGadgetsCheckBox.Checked = False
        End If
    End Sub

    Private Sub NextButton_Click(sender As Object, e As EventArgs) Handles NextButton.Click
        If MessageBox.Show("Are you sure you want to proceed?", "Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) = System.Windows.Forms.DialogResult.Yes Then
            ' Change values (settings) of programs.xml file in program folder
            If ModifyWizard.ChangeAddBarStyle = False And ChangeONECheckBox.Checked = False And ChangeUXCheckBox.Checked = False And ChangeQueroCheckBox.Checked = False And ChangeSysFilesCheckBox.Checked = False And InstallSoundsCheckBox.Checked = False And ChangeGadgetsCheckBox.Checked = False Then
                MessageBox.Show("You have selected nothing to change. Please choose something to change, or close this program.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Else
                Dim ConfigFile As New XmlDocument
                Dim ProgDir = ProgramFiles & "\Windows8.1to8RPConv"
                ConfigFile.Load(ProgDir & "\programs.xml")
                If ModifyWizard.ChangeAddBarStyle = True Then
                    If ModifyWizard.AddressBarStyle = "blue" Then
                        ConfigFile.SelectSingleNode("//setting[@name='AddressBarStyle']/value").InnerText = "white"
                    ElseIf ModifyWizard.AddressBarStyle = "white" Then
                        ConfigFile.SelectSingleNode("//setting[@name='AddressBarStyle']/value").InnerText = "blue"
                    End If
                End If
                If ChangeONECheckBox.Checked = True Then
                    If ModifyWizard.InstONE = True And FileIO.FileSystem.DirectoryExists(ProgramFiles & "\OldNewExplorer") Then
                        ConfigFile.SelectSingleNode("//setting[@name='InstONE']/value").InnerText = False
                    ElseIf ModifyWizard.InstONE = True And Not FileIO.FileSystem.DirectoryExists(ProgramFiles & "\OldNewExplorer") Then
                        ConfigFile.SelectSingleNode("//setting[@name='InstONE']/value").InnerText = True
                    ElseIf ModifyWizard.InstONE = False Then
                        ConfigFile.SelectSingleNode("//setting[@name='InstONE']/value").InnerText = True
                    End If
                End If
                If ChangeUXCheckBox.Checked = True Then
                    If ModifyWizard.InstUX = "UltraUX" Then
                        ConfigFile.SelectSingleNode("//setting[@name='InstUX']/value").InnerText = "UXTSB"
                    ElseIf ModifyWizard.InstUX = "UXTSB" Then
                        ConfigFile.SelectSingleNode("//setting[@name='InstUX']/value").InnerText = "UltraUX"
                    End If
                End If
                If ChangeQueroCheckBox.Checked = True Then
                    If ModifyWizard.InstQuero = False Then
                        ConfigFile.SelectSingleNode("//setting[@name='InstQuero']/value").InnerText = True
                    ElseIf ModifyWizard.InstQuero = True Then
                        ConfigFile.SelectSingleNode("//setting[@name='InstQuero']/value").InnerText = False
                    ElseIf ModifyWizard.InstQuero = True And Not FileIO.FileSystem.DirectoryExists(ProgramFiles & "\Quero Toolbar") Then
                        ConfigFile.SelectSingleNode("//setting[@name='InstQuero']/value").InnerText = True
                    End If
                End If
                If ChangeSysFilesCheckBox.Checked = True Then
                    If ModifyWizard.InstSysFiles = True Then
                        ConfigFile.SelectSingleNode("//setting[@name='InstSysFiles']/value").InnerText = False
                    ElseIf ModifyWizard.InstSysFiles = False Then
                        ConfigFile.SelectSingleNode("//setting[@name='InstSysFiles']/value").InnerText = True
                    End If
                End If
                If InstallSoundsCheckBox.Checked = True Then
                    ConfigFile.SelectSingleNode("//setting[@name='InstSounds']/value").InnerText = True
                End If
                If ChangeGadgetsCheckBox.Checked = True Then
                    If ModifyWizard.InstGadgets = True Then
                        ConfigFile.SelectSingleNode("//setting[@name='InstGadgets']/value").InnerText = False
                    ElseIf ModifyWizard.InstGadgets = True And Not FileIO.FileSystem.FileExists(ProgramFiles & "\Windows Sidebar\8GadgetPack.exe") Then
                        ConfigFile.SelectSingleNode("//setting[@name='InstGadgets']/value").InnerText = True
                    ElseIf ModifyWizard.InstGadgets = False Then
                        ConfigFile.SelectSingleNode("//setting[@name='InstGadgets']/value").InnerText = True
                    End If
                End If

                ' Finally save changes to programs.xml
                ConfigFile.Save(ProgDir & "\programs.xml")

                ' Load next screen
                Me.Dispose()
                ModifyModifyEnglish.Show()
            End If
        End If
    End Sub

    Private Sub BackButton_Click(sender As Object, e As EventArgs) Handles BackButton.Click
        Me.Dispose()
        ModifyThemeSelectEnglish.Show()
    End Sub

    Private Sub CloseButton_Click(sender As Object, e As EventArgs) Handles CloseButton.Click
        Me.Close()
    End Sub

    Private Sub ChangeONECheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles ChangeONECheckBox.CheckedChanged
        If ChangeONECheckBox.Checked = True Then
            ModifyWizard.ChangeONEChecked = True
        ElseIf ChangeONECheckBox.Checked = False Then
            ModifyWizard.ChangeONEChecked = False
        End If
    End Sub

    Private Sub ChangeUXCheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles ChangeUXCheckBox.CheckedChanged
        If ChangeUXCheckBox.Checked = True Then
            ModifyWizard.ChangeUXMethodChecked = True
            ModifyWizard.IsRestartNeeded = True
            RestartCount = RestartCount + 1
            ChangeUXPatchComboBox.Enabled = True
        ElseIf ChangeUXCheckBox.Checked = False Then
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
        ElseIf ChangeQueroCheckBox.Checked = False Then
            ModifyWizard.ChangeQueroChecked = False
        End If
    End Sub

    Private Sub ChangeSysFilesCheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles ChangeSysFilesCheckBox.CheckedChanged
        If ChangeSysFilesCheckBox.Checked = True Then
            ModifyWizard.ChangeSysFilesChecked = True
            ModifyWizard.IsRestartNeeded = True
            RestartCount = RestartCount + 1
        ElseIf ChangeSysFilesCheckBox.Checked = False Then
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
        ElseIf InstallSoundsCheckBox.Checked = False Then
            ModifyWizard.InstallSoundsChecked = False
            If ModifyWizard.ChangeAddBarStyle = False Then
                ModifyWizard.IsThemeChangeNeeded = False
            End If
        End If
    End Sub

    Private Sub ChangeGadgetsCheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles ChangeGadgetsCheckBox.CheckedChanged
        If ChangeGadgetsCheckBox.Checked = True Then
            ModifyWizard.ChangeGadgetsChecked = True
        ElseIf ChangeGadgetsCheckBox.Checked = False Then
            ModifyWizard.ChangeGadgetsChecked = False
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
