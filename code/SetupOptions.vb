Imports Microsoft.Win32
Imports System.Globalization

Public Class SetupOptions
    Private Sub SetupOptions_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        InitialiseOptions()
    End Sub

    Private Sub InitialiseOptions()
        ' Check bitness of system and choose bitness automatically
        If Environment.Is64BitOperatingSystem = True Then
            CPU64BitRadioButton.Checked = True
            CPU32BitRadioButton.Enabled = False
            SetupWizard.BitnessSystem = "64Bit"
        Else
            CPU32BitRadioButton.Checked = True
            CPU64BitRadioButton.Enabled = False
            SetupWizard.BitnessSystem = "32Bit"
        End If

        ' Show which address bar style has been chosen
        If SetupWizard.AddressBarStyle = "BlueAddressBar" Then
            SelectedStylePictureBox.Image = My.Resources.BlueAddressBar
        ElseIf SetupWizard.AddressBarStyle = "WhiteAddressBar" Then
            SelectedStylePictureBox.Image = My.Resources.WhiteAddressBar
        End If

        ' Check if Aero Glass is already installed via uninstall registry key
        Using UninstallRegKey As RegistryKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall", False)
            Dim AeroGlassRegKeys As New List(Of String)
            Dim AeroGlassPath As String = String.Empty
            For Each RegKey In UninstallRegKey.GetSubKeyNames
                If RegKey.StartsWith("Aero Glass for Win8.1") Then
                    AeroGlassRegKeys.Add(RegKey)
                End If
            Next
            If AeroGlassRegKeys.Count > 0 Then
                Using AeroGlassRegKey As RegistryKey = UninstallRegKey.OpenSubKey(AeroGlassRegKeys(0))
                    AeroGlassPath = AeroGlassRegKey.GetValue("InstallLocation")
                    If FileIO.FileSystem.DirectoryExists(AeroGlassPath) Then
                        SetupWizard.AeroGlassInstalled = True
                        AeroGlassCheckBox.Text = "Reinstall Aero Glass by Big Muscle"
                        AeroGlassCheckBox.Enabled = True
                        InfoAeroGlassCheckBox.Visible = False
                        If SetupWizard.InstAeroGlass = True Then
                            AeroGlassCheckBox.Checked = True
                        Else
                            AeroGlassCheckBox.Checked = False
                        End If
                    Else
                        SetupWizard.AeroGlassInstalled = False
                        SetupWizard.InstAeroGlass = True
                    End If
                End Using
            End If
        End Using

        If SetupWizard.AddressBarStyle = "BlueAddressBar" Then
            QueroCheckBox.Enabled = True
            InfoQueroCheckBox.Enabled = True
            If SetupWizard.InstQuero = True Then
                QueroCheckBox.Checked = True
            ElseIf SetupWizard.InstQuero = False Then
                QueroCheckBox.Checked = False
            End If
        ElseIf SetupWizard.AddressBarStyle = "WhiteAddressBar" Then
            SetupWizard.InstQuero = False
            QueroCheckBox.Enabled = False
            QueroCheckBox.Checked = False
        End If

        ' Load default setting for UXPatch combo box
        UXPatchComboBox.SelectedIndex = 0

        ' If program language is not UI language, disable system files replacement
        If SetupWizard.IsSelectedLanguageDisplayLanguage = False Then
            SysFilesCheckBox.Enabled = False
            SysFilesCheckBox.Checked = False
            InfoReplaceSysFilesCheckBox.Image = My.Resources.WarningSmall
            InfoReplaceSysFilesCheckBox.Cursor = Cursors.Help
            ToolTipSetup.SetToolTip(InfoReplaceSysFilesCheckBox, "The setup language does not match the system’s UI language, so replacing system files is disabled.")
        ElseIf SetupWizard.IsSelectedLanguageDisplayLanguage = True Then
            If SetupWizard.InstSysFiles = True Then
                SysFilesCheckBox.Checked = True
            ElseIf SetupWizard.InstSysFiles = False Then
                SysFilesCheckBox.Checked = False
            End If
        End If

        ' Set program directory label
        InstDirectoryLabel.Text = "Program path: " & SetupWizard.ProgDir
    End Sub

    Private Sub AeroGlassCheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles AeroGlassCheckBox.CheckedChanged
        If AeroGlassCheckBox.Checked = True Then
            SetupWizard.InstAeroGlass = True
        Else
            SetupWizard.InstAeroGlass = False
        End If
    End Sub

    Private Sub InfoAeroGlassCheckBox_Click(sender As Object, e As EventArgs) Handles InfoAeroGlassCheckBox.Click
        MessageBox.Show("Aero Glass is needed to enable transparency effects which were natively available in Windows 8 Release Preview.", "Info about Aero Glass", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Private Sub ONECheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles ONECheckBox.CheckedChanged
        If ONECheckBox.Checked = True Then
            SetupWizard.InstONE = True
        Else
            MessageBox.Show("It is not recommended to discard the installation of OldNewExplorer, otherwise the theme cannot be applied fully.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information)
            SetupWizard.InstONE = False
        End If
    End Sub

    Private Sub InfoONECheckBox_Click(sender As Object, e As EventArgs) Handles InfoONECheckBox.Click
        SetupOptionsInstallONEHelp.ShowDialog()
    End Sub

    Private Sub UXPatchComboBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles UXPatchComboBox.SelectedIndexChanged
        Select Case UXPatchComboBox.SelectedItem
            Case "UltraUXThemePatcher by Manuel Hoefs"
                SetupWizard.InstUX = "UltraUX"
            Case "UXThemeSignatureBypass by Big Muscle"
                SetupWizard.InstUX = "UXTSB"
            Case "I already have patched UXTheme"
                MessageBox.Show("Please make sure that you have already installed an UXTheme patcher, otherwise the theme cannot be applied to your system!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                SetupWizard.InstUX = "NoUXPatch"
        End Select
    End Sub

    Private Sub InfoMethodCheckBox_Click(sender As Object, e As EventArgs) Handles InfoMethodCheckBox.Click
        SetupOptionsInstallUXTHelp.ShowDialog()
    End Sub

    Private Sub QueroCheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles QueroCheckBox.CheckedChanged
        If QueroCheckBox.Checked = True Then
            SetupWizard.InstQuero = True
        Else
            SetupWizard.InstQuero = False
            If SetupWizard.AddressBarStyle = "BlueAddressBar" Then
                MessageBox.Show("When the theme with the blue address bar is being installed, Internet Explorer will be messed visually by showing an opaque instead of a transparent address bar. Quero Toolbar is a workaround which replaces the native address bar of Internet Explorer to match the Aero Glass, i.e. transparent visual style.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            End If
        End If
    End Sub

    Private Sub InfoQueroCheckBox_Click(sender As Object, e As EventArgs) Handles InfoQueroCheckBox.Click
        MessageBox.Show("This toolbar replaces the native address bar of Internet Explorer, to make it look right with the Aero theme. It must be enabled and set up manually after installation, when Internet Explorer will autostart an XHTML file containing the manual to do it.", "Info about Quero Toolbar", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Private Sub SysFilesCheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles SysFilesCheckBox.CheckedChanged
        If SysFilesCheckBox.Checked = True Then
            SetupWizard.InstSysFiles = True
        Else
            SetupWizard.InstSysFiles = False
        End If
    End Sub

    Private Sub SoundsCheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles SoundsCheckBox.CheckedChanged
        If SoundsCheckBox.Checked = True Then
            SetupWizard.InstSounds = True
        Else
            SetupWizard.InstSounds = False
        End If
    End Sub

    Private Sub InfoSoundSchemeCheckBox_Click(sender As Object, e As EventArgs) Handles InfoSoundSchemeCheckBox.Click
        MessageBox.Show("If you enable this option, the original sound scheme of Windows 8 Release Preview, which has been copied from an original installation of it, will be installed and set as default on your system.", "Info about sound scheme", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Private Sub InfoReplaceSysFilesCheckBox_Click(sender As Object, e As EventArgs) Handles InfoReplaceSysFilesCheckBox.Click
        If SetupWizard.IsSelectedLanguageDisplayLanguage = True Then
            SetupOptionsReplaceSysFilesHelp.ShowDialog()
        End If
    End Sub

    Private Sub GadgetsCheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles GadgetsCheckBox.CheckedChanged
        If GadgetsCheckBox.Checked = True Then
            SetupWizard.InstGadgets = True
        Else
            SetupWizard.InstGadgets = False
        End If
    End Sub

    Private Sub TaskbarTweakerCheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles TaskbarTweakerCheckBox.CheckedChanged
        If TaskbarTweakerCheckBox.Checked = True Then
            SetupWizard.Inst7TaskTw = True
        Else
            SetupWizard.Inst7TaskTw = False
        End If
    End Sub

    Private Sub InfoGadgetsCheckBox_Click(sender As Object, e As EventArgs) Handles InfoGadgetsCheckBox.Click
        SetupOptionsInstallGadgetsHelp.ShowDialog()
    End Sub

    Private Sub InfoTaskbarTweakerCheckBox_Click(sender As Object, e As EventArgs) Handles InfoTaskbarTweakerCheckBox.Click
        SetupOptionsInstall7TaskTwHelp.ShowDialog()
    End Sub

    Private Sub SaveSettingsToXmlFile()
        With SetupWizard.ProgSettings
            .ProgLanguage = SetupWizard.ProgLanguage
            .AddressBarStyle = SetupWizard.AddressBarStyle
            .BitnessSystem = SetupWizard.BitnessSystem
            .InstAeroGlass = SetupWizard.InstAeroGlass
            .InstONE = SetupWizard.InstONE
            .InstUX = SetupWizard.InstUX
            .InstQuero = SetupWizard.InstQuero
            .InstSysFiles = SetupWizard.InstSysFiles
            .InstSounds = SetupWizard.InstSounds
            .InstGadgets = SetupWizard.InstGadgets
            .Inst7TaskTw = SetupWizard.Inst7TaskTw
            .Save()
        End With
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
        TaskbarTweakerCheckBox.Checked = True
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
        TaskbarTweakerCheckBox.Checked = False
    End Sub

    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()
    End Sub

    Private Sub InfoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles InfoToolStripMenuItem.Click
        AboutProgram.ShowDialog()
    End Sub

    Private Sub ContactMeToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ContactMeToolStripMenuItem.Click
        ContactMe.ShowDialog()
    End Sub

    Private Sub BackButton_Click(sender As Object, e As EventArgs) Handles BackButton.Click
        Me.Dispose()
        SetupThemeSelect.Show()
    End Sub

    Private Sub NextButton_Click(sender As Object, e As EventArgs) Handles NextButton.Click
        If MessageBox.Show("Please make sure to save any unsaved documents, as well as close any programs to ensure a successful installation. Also, it is strongly recommended that your Windows 8.1 installation is fully updated, otherwise your system may be bricked. This setup will create a system restore point though, that you can easily restore your system back to normal. Are you sure you want to proceed?", "Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) = System.Windows.Forms.DialogResult.Yes Then
            SaveSettingsToXmlFile()
            ' Start installation now
            Me.Dispose()
            SetupInstall.Show()
        End If
    End Sub

    Private Sub CloseButton_Click(sender As Object, e As EventArgs) Handles CloseButton.Click
        Me.Close()
    End Sub

    Private Sub SetupOptions_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        If e.CloseReason = CloseReason.UserClosing Then
            If MessageBox.Show("Are you sure you want to cancel? This program will be closed.", "Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) = System.Windows.Forms.DialogResult.Yes Then
                SetupWizard.Close()
            ElseIf System.Windows.Forms.DialogResult.No Then
                e.Cancel = True
            End If
        End If
    End Sub
End Class