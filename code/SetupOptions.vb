Imports Microsoft.Win32
Imports System.Globalization

Public Class SetupOptionsEnglish
    Private Sub SetupOptionsEnglish_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Check language and region in menu when program loads
        If SetupProgram.Language = "en-GB" Then
            EnglishToolStripMenuItem.Checked = True
            EnglishUnitedKingdomToolStripMenuItem.Checked = True
            EnglishUnitedKingdomToolStripMenuItem.Enabled = False
        ElseIf SetupProgram.Language = "en-US" Then
            EnglishToolStripMenuItem.Checked = True
            EnglishUnitedStatesToolStripMenuItem.Checked = True
            EnglishUnitedStatesToolStripMenuItem.Enabled = False
        End If

        ' If language choice does not match display language
        If SetupProgram.Language <> CultureInfo.CurrentUICulture.ToString Then
            SetupProgram.IsSelectedLanguageDisplayLanguage = False
        Else
            SetupProgram.IsSelectedLanguageDisplayLanguage = True
        End If

        ' Check bitness of system and choose bitness automatically
        If Environment.Is64BitOperatingSystem = True Then
            CPU64BitRadioButton.Checked = True
            CPU32BitRadioButton.Enabled = False
        Else
            CPU32BitRadioButton.Checked = True
            CPU64BitRadioButton.Enabled = False
        End If

        ' Check if Aero Glass is already installed via uninstall registry key
        Dim UninstallRegKey As RegistryKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall", False)
        Dim AeroGlassRegKeys As New List(Of String)
        Dim AeroGlassPath As String = String.Empty
        For Each RegKey In UninstallRegKey.GetSubKeyNames
            If RegKey.StartsWith("Aero Glass for Win8.1") Then
                AeroGlassRegKeys.Add(RegKey)
            End If
        Next
        If AeroGlassRegKeys.Count > 0 Then
            Dim AeroGlassRegKey = UninstallRegKey.OpenSubKey(AeroGlassRegKeys(0))
            AeroGlassPath = AeroGlassRegKey.GetValue("InstallLocation")
            AeroGlassRegKey.Close()
            UninstallRegKey.Close()
            If FileIO.FileSystem.DirectoryExists(AeroGlassPath) Then
                SetupProgram.AeroGlassInstalled = True
                AeroGlassCheckBox.Text = "Reinstall Aero Glass by Big Muscle"
                AeroGlassCheckBox.Enabled = True
                InfoAeroGlassCheckBox.Visible = False
                If SetupProgram.InstAeroGlass = True Then
                    AeroGlassCheckBox.Checked = True
                ElseIf SetupProgram.InstAeroGlass = False Then
                    AeroGlassCheckBox.Checked = False
                End If
            Else
                SetupProgram.AeroGlassInstalled = False
                SetupProgram.InstAeroGlass = True
            End If
        End If

        ' Set selected settings when form loads
        If SetupProgram.InstONE = True Then
            ONECheckBox.Checked = True
        Else
            ONECheckBox.Checked = False
        End If
        If SetupProgram.InstUX = "UltraUX" Then
            UXPatchComboBox.SelectedIndex = 0
        ElseIf SetupProgram.InstUX = "UXTSB" Then
            UXPatchComboBox.SelectedIndex = 1
        ElseIf SetupProgram.InstUX = "none" Then
            UXPatchComboBox.SelectedIndex = 2
        End If
        If SetupProgram.AddressBarStyle = "blue" Then
            QueroCheckBox.Enabled = True
            If SetupProgram.InstQuero = True Then
                QueroCheckBox.Checked = True
            ElseIf SetupProgram.InstQuero = False Then
                QueroCheckBox.Checked = False
            End If
        ElseIf SetupProgram.AddressBarStyle = "white" Then
            SetupProgram.InstQuero = False
            QueroCheckBox.Enabled = False
            QueroCheckBox.Checked = False
        End If
        AddBarChoiceShowLabel.Text = SetupProgram.AddressBarStyle
        '' If program language is not UI language, disable system files replacement
        If SetupProgram.IsSelectedLanguageDisplayLanguage = False Then
            SysFilesCheckBox.Enabled = False
            SysFilesCheckBox.Checked = False
            InfoReplaceSysFilesCheckBox.Image = My.Resources.WarningSmall
            InfoReplaceSysFilesCheckBox.Cursor = Cursors.Help
            ToolTipSetup.SetToolTip(InfoReplaceSysFilesCheckBox, "The setup language does not match the system’s UI language, so replacing system files is disabled.")
        ElseIf SetupProgram.IsSelectedLanguageDisplayLanguage = True Then
            If SetupProgram.InstSysFiles = True Then
                SysFilesCheckBox.Checked = True
            ElseIf SetupProgram.InstSysFiles = False Then
                SysFilesCheckBox.Checked = False
            End If
        End If
        If SetupProgram.InstSounds = True Then
            SoundsCheckBox.Checked = True
        ElseIf SetupProgram.InstSounds = False Then
            SoundsCheckBox.Checked = False
        End If
        If SetupProgram.InstGadgets = True Then
            GadgetsCheckBox.Checked = True
        ElseIf SetupProgram.InstGadgets = False Then
            GadgetsCheckBox.Checked = False
        End If

        ' Set variables for program directory label
        Dim ProgramFiles = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles)
        Dim ProgDir As String = ProgramFiles & "\Windows8.1to8RPConv"
        InstDirectoryLabel.Text = InstDirectoryLabel.Text & ProgDir
    End Sub

    Private Sub CPU64BitRadioButton_CheckedChanged(sender As Object, e As EventArgs) Handles CPU64BitRadioButton.CheckedChanged
        SetupProgram.BitnessSystem = "64bit"
    End Sub

    Private Sub CPU32bitRadioButton_CheckedChanged(sender As Object, e As EventArgs) Handles CPU32BitRadioButton.CheckedChanged
        SetupProgram.BitnessSystem = "32bit"
    End Sub

    Private Sub AeroGlassCheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles AeroGlassCheckBox.CheckedChanged
        If AeroGlassCheckBox.Checked = True Then
            SetupProgram.InstAeroGlass = True
        ElseIf AeroGlassCheckBox.Checked = False Then
            SetupProgram.InstAeroGlass = False
        End If
    End Sub

    Private Sub InfoAeroGlassCheckBox_Click(sender As Object, e As EventArgs) Handles InfoAeroGlassCheckBox.Click
        MessageBox.Show("Aero Glass is needed to enable transparency effects which were natively available in Windows 8 Release Preview.", "Info about Aero Glass", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Private Sub ONECheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles ONECheckBox.CheckedChanged
        If ONECheckBox.Checked = True Then
            SetupProgram.InstONE = True
        ElseIf ONECheckBox.Checked = False Then
            SetupProgram.InstONE = False
            MessageBox.Show("It is not recommended to discard the installation of OldNewExplorer, otherwise the theme cannot be applied fully.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub

    Private Sub InfoONECheckBox_Click(sender As Object, e As EventArgs) Handles InfoONECheckBox.Click
        SetupOptionsInstallONEHelpEnglish.ShowDialog()
    End Sub

    Private Sub UXPatchComboBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles UXPatchComboBox.SelectedIndexChanged
        If UXPatchComboBox.SelectedIndex = 0 Then
            SetupProgram.InstUX = "UltraUX"
        ElseIf UXPatchComboBox.SelectedIndex = 1 Then
            SetupProgram.InstUX = "UXTSB"
        ElseIf UXPatchComboBox.SelectedIndex = 2 Then
            MessageBox.Show("Please make sure that you have already installed an UXTheme patcher, otherwise the theme will not be applied!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            SetupProgram.InstUX = "none"
        End If
    End Sub

    Private Sub InfoMethodCheckBox_Click(sender As Object, e As EventArgs) Handles InfoMethodCheckBox.Click
        SetupOptionsInstallUXTHelpEnglish.ShowDialog()
    End Sub

    Private Sub QueroCheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles QueroCheckBox.CheckedChanged
        If QueroCheckBox.Checked = True Then
            SetupProgram.InstQuero = True
        ElseIf QueroCheckBox.Checked = False Then
            SetupProgram.InstQuero = False
            MessageBox.Show("When the theme with the blue address bar is being installed, Internet Explorer will be messed visually by showing an opaque instead of a transparent address bar. Quero Toolbar is a workaround which replaces the native address bar of Internet Explorer to match the Aero Glass, i.e. transparent visual style.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End If
    End Sub

    Private Sub InfoQueroCheckBox_Click(sender As Object, e As EventArgs) Handles InfoQueroCheckBox.Click
        MessageBox.Show("This toolbar replaces the native address bar of Internet Explorer, to make it look right with the Aero theme. It must be enabled and set up manually after installation, when Internet Explorer will autostart an XHTML file containing the manual to do it.", "Info about Quero Toolbar", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Private Sub SysFilesCheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles SysFilesCheckBox.CheckedChanged
        If SysFilesCheckBox.Checked = True Then
            SetupProgram.InstSysFiles = True
        ElseIf SysFilesCheckBox.Checked = False Then
            SetupProgram.InstSysFiles = False
        End If
    End Sub

    Private Sub SoundsCheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles SoundsCheckBox.CheckedChanged
        If SoundsCheckBox.Checked = True Then
            SetupProgram.InstSounds = True
        ElseIf SoundsCheckBox.Checked = False Then
            SetupProgram.InstSounds = False
        End If
    End Sub

    Private Sub InfoSoundSchemeCheckBox_Click(sender As Object, e As EventArgs) Handles InfoSoundSchemeCheckBox.Click
        MessageBox.Show("If you enable this option, the original sound scheme of Windows 8 Release Preview, which has been copied from an original installation of it, will be installed and set as default on your system.", "Info about sound scheme", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Private Sub InfoReplaceSysFilesCheckBox_Click(sender As Object, e As EventArgs) Handles InfoReplaceSysFilesCheckBox.Click
        If SetupProgram.IsSelectedLanguageDisplayLanguage = True Then
            SetupOptionsReplaceSysFilesHelpEnglish.ShowDialog()
        End If
    End Sub

    Private Sub GadgetsCheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles GadgetsCheckBox.CheckedChanged
        If GadgetsCheckBox.Checked = True Then
            SetupProgram.InstGadgets = True
        ElseIf GadgetsCheckBox.Checked = False Then
            SetupProgram.InstGadgets = False
        End If
    End Sub

    Private Sub InfoGadgetsCheckBox_Click(sender As Object, e As EventArgs) Handles InfoGadgetsCheckBox.Click
        SetupOptionsInstallGadgetsHelpEnglish.ShowDialog()
    End Sub

    Private Sub SelectAllButton_Click(sender As Object, e As EventArgs) Handles SelectAllButton.Click
        If AeroGlassCheckBox.Enabled = True Then
            AeroGlassCheckBox.Checked = True
        End If
        ONECheckBox.Checked = True
        If QueroCheckBox.Enabled = True Then
            QueroCheckBox.Checked = True
        End If
        If SysFilesCheckBox.Enabled = True Then
            SysFilesCheckBox.Checked = True
        End If
        SoundsCheckBox.Checked = True
        GadgetsCheckBox.Checked = True
    End Sub

    Private Sub ClearAllButton_Click(sender As Object, e As EventArgs) Handles ClearAllButton.Click
        If AeroGlassCheckBox.Enabled = True Then
            AeroGlassCheckBox.Checked = False
        End If
        ONECheckBox.Checked = False
        If QueroCheckBox.Enabled = True Then
            QueroCheckBox.Checked = False
        End If
        If SysFilesCheckBox.Enabled = True Then
            SysFilesCheckBox.Checked = False
        End If
        SoundsCheckBox.Checked = False
        GadgetsCheckBox.Checked = False
    End Sub

    Private Sub InfoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles InfoToolStripMenuItem.Click
        SetupToolStripInfo.ShowDialog()
    End Sub

    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()
    End Sub

    Private Sub BackButton_Click(sender As Object, e As EventArgs) Handles BackButton.Click
        Me.Dispose()
        SetupThemeSelectEnglish.Show()
    End Sub

    Private Sub NextButton_Click(sender As Object, e As EventArgs) Handles NextButton.Click
        If MessageBox.Show("Please make sure that your Windows 8.1 installation is fully updated and all programs are closed, otherwise your system may be bricked. I will create a system restore point prior setup so if something goes wrong, you can easily restore your system back to normal. Are you sure you want to proceed?", "Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) = System.Windows.Forms.DialogResult.Yes Then
            SetupProgram.ProgSettings.Lang = SetupProgram.Language
            SetupProgram.ProgSettings.AddressBarStyle = SetupProgram.AddressBarStyle
            SetupProgram.ProgSettings.BitnessSystem = SetupProgram.BitnessSystem
            SetupProgram.ProgSettings.InstAeroGlass = SetupProgram.InstAeroGlass
            SetupProgram.ProgSettings.InstONE = SetupProgram.InstONE
            SetupProgram.ProgSettings.InstUX = SetupProgram.InstUX
            SetupProgram.ProgSettings.InstQuero = SetupProgram.InstQuero
            SetupProgram.ProgSettings.InstSysFiles = SetupProgram.InstSysFiles
            SetupProgram.ProgSettings.InstSounds = SetupProgram.InstSounds
            SetupProgram.ProgSettings.InstGadgets = SetupProgram.InstGadgets
            SetupProgram.ProgSettings.Save()
            Me.Dispose()
            SetupInstallEnglish.Show()
        End If
    End Sub

    Private Sub CloseButton_Click(sender As Object, e As EventArgs) Handles CloseButton.Click
        Me.Close()
    End Sub

    Private Sub SetupOptionsEnglish_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        If e.CloseReason = CloseReason.UserClosing Then
            If MessageBox.Show("Are you sure you want to cancel? This program will be closed.", "Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) = System.Windows.Forms.DialogResult.Yes Then
                SetupProgram.Close()
            ElseIf System.Windows.Forms.DialogResult.No Then
                e.Cancel = True
            End If
        End If
    End Sub
End Class
