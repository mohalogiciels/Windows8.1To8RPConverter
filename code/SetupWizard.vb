Imports Microsoft.Win32
Imports System.Configuration
Imports System.IO
Imports System.Globalization

Public Class SetupWizard
    ' Initialise program settings
    Public Shared ProgSettings As New My.MySettings

    ' Load default program settings
    Public Shared ProgLanguage As String = ProgSettings.ProgLanguage
    Public AddressBarStyle As String = ProgSettings.AddressBarStyle
    Public Shared BitnessSystem As String = ProgSettings.BitnessSystem
    Public InstONE As Boolean = ProgSettings.InstONE
    Public InstUX As String = ProgSettings.InstUX
    Public InstQuero As Boolean = ProgSettings.InstQuero
    Public InstSysFiles As Boolean = ProgSettings.InstSysFiles
    Public InstSounds As Boolean = ProgSettings.InstSounds
    Public InstGadgets As Boolean = ProgSettings.InstGadgets
    Public Inst7TaskTw As Boolean = ProgSettings.Inst7TaskTw

    ' Variables for check if Aero Glass is already installed
    Public InstAeroGlass As Boolean = False
    Public AeroGlassInstalled As Boolean = False

    ' Variables for program
    Dim ProgramFiles = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles)
    Public ProgDir = ProgramFiles & "\Windows 8.1 to 8 RP Converter"
    Public LanguageForLocalisedStrings As String = CultureInfo.CurrentUICulture.Name.Replace("-", "_")
    Public IsSelectedLanguageDisplayLanguage As Boolean = True

    Private Sub SetupWizard_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Set language for English to universal i.e. no region
        If CultureInfo.CurrentUICulture.Name.StartsWith("en") Then
            LanguageForLocalisedStrings = "en"
        End If

        ' Check if OS is Windows 8.1
        Using WindowsVersion As RegistryKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\Microsoft\Windows NT\CurrentVersion", False)
            If WindowsVersion.GetValue("CurrentBuild") <> "9600" Then
                If My.Resources.ResourceManager.GetString("IsNotWin8_" & LanguageForLocalisedStrings) IsNot Nothing Then
                    MessageBox.Show(My.Resources.ResourceManager.GetString("IsNotWin8_" & LanguageForLocalisedStrings), My.Resources.ResourceManager.GetString("NotSupportedOS_" & LanguageForLocalisedStrings), MessageBoxButtons.OK, MessageBoxIcon.Error)
                Else
                    MessageBox.Show(My.Resources.ResourceManager.GetString("IsNotWin8_en"), My.Resources.ResourceManager.GetString("NotSupportedOS_en"), MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If
                Me.Close()
            End If
        End Using

        ' Check if 64-bit program started when system is 64-bit
        If Environment.Is64BitOperatingSystem <> Environment.Is64BitProcess Then
            If My.Resources.ResourceManager.GetString("No64BitProcess_" & LanguageForLocalisedStrings) IsNot Nothing Then
                MessageBox.Show(My.Resources.ResourceManager.GetString("No64BitProcess_" & LanguageForLocalisedStrings), My.Resources.ResourceManager.GetString("Error_" & LanguageForLocalisedStrings), MessageBoxButtons.OK, MessageBoxIcon.Error)
            Else
                MessageBox.Show(My.Resources.ResourceManager.GetString("No64BitProcess_en"), My.Resources.ResourceManager.GetString("Error_en"), MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
            Me.Close()
        End If

        ' Check if program has already been installed
        If FileIO.FileSystem.DirectoryExists(ProgDir) Then
            If My.Resources.ResourceManager.GetString("AlreadyInstalled_" & LanguageForLocalisedStrings) IsNot Nothing Then
                MessageBox.Show(My.Resources.ResourceManager.GetString("AlreadyInstalled_" & LanguageForLocalisedStrings), My.Resources.ResourceManager.GetString("Warning_" & LanguageForLocalisedStrings), MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Else
                MessageBox.Show(My.Resources.ResourceManager.GetString("AlreadyInstalled_en"), My.Resources.ResourceManager.GetString("Warning_en"), MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            End If
            Me.Close()
        End If
    End Sub

    Private Sub SetupWizard_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown
        ' Detect system language and choose language automatically
        If CultureInfo.CurrentUICulture.ThreeLetterWindowsLanguageName = "DEU" Then
            LangSelectCB.SelectedIndex = 0
        ElseIf CultureInfo.CurrentUICulture.ThreeLetterWindowsLanguageName = "ENG" Then
            LangSelectCB.SelectedIndex = 1
        ElseIf CultureInfo.CurrentUICulture.ThreeLetterWindowsLanguageName = "ENU" Then
            LangSelectCB.SelectedIndex = 2
        ElseIf CultureInfo.CurrentUICulture.ThreeLetterWindowsLanguageName = "ESN" Then
            LangSelectCB.SelectedIndex = 3
        ElseIf CultureInfo.CurrentUICulture.ThreeLetterWindowsLanguageName = "FRA" Then
            LangSelectCB.SelectedIndex = 4
        ElseIf CultureInfo.CurrentUICulture.ThreeLetterWindowsLanguageName = "ITA" Then
            LangSelectCB.SelectedIndex = 5
        ElseIf CultureInfo.CurrentUICulture.ThreeLetterWindowsLanguageName = "NLD" Then
            LangSelectCB.SelectedIndex = 6
        ElseIf CultureInfo.CurrentUICulture.ThreeLetterWindowsLanguageName = "PTB" Then
            LangSelectCB.SelectedIndex = 7
        ElseIf CultureInfo.CurrentUICulture.ThreeLetterWindowsLanguageName = "PTG" Then
            LangSelectCB.SelectedIndex = 8
        ElseIf CultureInfo.CurrentUICulture.ThreeLetterWindowsLanguageName = "KOR" Then
            LangSelectCB.SelectedIndex = 9
        ElseIf CultureInfo.CurrentUICulture.ThreeLetterWindowsLanguageName = "RUS" Then
            LangSelectCB.SelectedIndex = 10
        ElseIf CultureInfo.CurrentUICulture.ThreeLetterWindowsLanguageName = "UKR" Then
            LangSelectCB.SelectedIndex = 11
        Else
            LangSelectCB.SelectedIndex = 2
        End If
    End Sub

    Private Sub LangSelectCB_SelectedIndexChanged(sender As Object, e As EventArgs) Handles LangSelectCB.SelectedIndexChanged
        Select Case LangSelectCB.SelectedIndex
            Case 0
                ProgLanguage = "de-DE"
                LanguageForLocalisedStrings = ProgLanguage.Replace("-", "_")
            Case 1
                ProgLanguage = "en-GB"
                LanguageForLocalisedStrings = "en"
            Case 2
                ProgLanguage = "en-US"
                LanguageForLocalisedStrings = "en"
            Case 3
                ProgLanguage = "es-ES"
                LanguageForLocalisedStrings = ProgLanguage.Replace("-", "_")
            Case 4
                ProgLanguage = "fr-FR"
                LanguageForLocalisedStrings = ProgLanguage.Replace("-", "_")
            Case 5
                ProgLanguage = "it-IT"
                LanguageForLocalisedStrings = ProgLanguage.Replace("-", "_")
            Case 6
                ProgLanguage = "nl-NL"
                LanguageForLocalisedStrings = ProgLanguage.Replace("-", "_")
            Case 7
                ProgLanguage = "pt-BR"
                LanguageForLocalisedStrings = ProgLanguage.Replace("-", "_")
            Case 8
                ProgLanguage = "pt-PT"
                LanguageForLocalisedStrings = ProgLanguage.Replace("-", "_")
            Case 9
                ProgLanguage = "ko-KR"
                LanguageForLocalisedStrings = ProgLanguage.Replace("-", "_")
            Case 10
                ProgLanguage = "ru-RU"
                LanguageForLocalisedStrings = ProgLanguage.Replace("-", "_")
            Case 11
                ProgLanguage = "uk-UA"
                LanguageForLocalisedStrings = ProgLanguage.Replace("-", "_")
        End Select

        ' Set localised strings for UI on SetupWizard window
        Me.Text = My.Resources.ResourceManager.GetString("SetupWizardText_" & LanguageForLocalisedStrings)
        LangSelectLabel.Text = My.Resources.ResourceManager.GetString("LangSelectLabel_" & LanguageForLocalisedStrings)
        OKButton.Text = My.Resources.ResourceManager.GetString("OKButton_" & LanguageForLocalisedStrings)
        CloseButton.Text = My.Resources.ResourceManager.GetString("CloseButton_" & LanguageForLocalisedStrings)
    End Sub

    Private Sub OKButton_Click(sender As Object, e As EventArgs) Handles OKButton.Click
        ' If language choice does not match display language
        If ProgLanguage = CultureInfo.CurrentUICulture.Name Then
            IsSelectedLanguageDisplayLanguage = True
        Else
            IsSelectedLanguageDisplayLanguage = False
        End If

        ' Load program finally
        Me.Hide()
        SetupWelcome.Show()
    End Sub

    Private Sub CancelButton_Click(sender As Object, e As EventArgs) Handles CloseButton.Click
        Me.Close()
    End Sub

    Private Sub SetupWizard_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        ' Delete temporary files, if existing, when closing program -> delete program folder in LocalAppData 
        If FileIO.FileSystem.DirectoryExists(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) & "\SetupWindows8RPConv") Then
            FileIO.FileSystem.DeleteDirectory(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) & "\SetupWindows8RPConv", FileIO.DeleteDirectoryOption.DeleteAllContents)
        End If
    End Sub
End Class
