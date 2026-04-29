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
    Public LanguageForLocalisedStrings As String = String.Empty
    Public IsSelectedLanguageDisplayLanguage As Boolean = False

    Private Sub SetupWizard_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Detect system language and choose language automatically
        Select Case CultureInfo.CurrentUICulture.ThreeLetterWindowsLanguageName
            Case "DEU"
                ProgLanguage = "de-DE"
                LanguageForLocalisedStrings = ProgLanguage.Replace("-", "_")
            Case "ENG"
                ProgLanguage = "en-GB"
                LanguageForLocalisedStrings = "en"
            Case "ENU"
                ProgLanguage = "en-US"
                LanguageForLocalisedStrings = "en"
            Case "ESN"
                ProgLanguage = "es-ES"
                LanguageForLocalisedStrings = ProgLanguage.Replace("-", "_")
            Case "FRA"
                ProgLanguage = "fr-FR"
                LanguageForLocalisedStrings = ProgLanguage.Replace("-", "_")
            Case "ITA"
                ProgLanguage = "it-IT"
                LanguageForLocalisedStrings = ProgLanguage.Replace("-", "_")
            Case "NLD"
                ProgLanguage = "nl-NL"
                LanguageForLocalisedStrings = ProgLanguage.Replace("-", "_")
            Case "PTB"
                ProgLanguage = "pt-BR"
                LanguageForLocalisedStrings = ProgLanguage.Replace("-", "_")
            Case "PTG"
                ProgLanguage = "pt-PT"
                LanguageForLocalisedStrings = ProgLanguage.Replace("-", "_")
            Case "KOR"
                ProgLanguage = "ko-KR"
                LanguageForLocalisedStrings = ProgLanguage.Replace("-", "_")
            Case "RUS"
                ProgLanguage = "ru-RU"
                LanguageForLocalisedStrings = ProgLanguage.Replace("-", "_")
            Case "UKR"
                ProgLanguage = "uk-UA"
                LanguageForLocalisedStrings = ProgLanguage.Replace("-", "_")
            Case Else
                ProgLanguage = "en-US"
                LanguageForLocalisedStrings = "en"
        End Select

        ' Check if UI language is in program language list
        If CultureInfo.CurrentUICulture.Name = ProgLanguage Then
            IsSelectedLanguageDisplayLanguage = True
        Else
            IsSelectedLanguageDisplayLanguage = False
        End If

        ' Check if OS is Windows 8.1
        Using WindowsVersion As RegistryKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\Microsoft\Windows NT\CurrentVersion", False)
            If WindowsVersion.GetValue("CurrentBuild") <> "9600" Then
                If My.Resources.ResourceManager.GetString("IsNotWin8_" & LanguageForLocalisedStrings) IsNot Nothing Then
                    MessageBox.Show(My.Resources.ResourceManager.GetString("IsNotWin8_" & LanguageForLocalisedStrings), My.Resources.ResourceManager.GetString("NotSupportedOS_" & LanguageForLocalisedStrings), MessageBoxButtons.OK, MessageBoxIcon.Error)
                Else
                    MessageBox.Show(My.Resources.ResourceManager.GetString("IsNotWin8_en"), My.Resources.ResourceManager.GetString("NotSupportedOS_en"), MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If
                Application.Exit()
                Exit Sub
            End If
        End Using

        ' Check if 64-bit program started when system is 64-bit
        If Environment.Is64BitOperatingSystem <> Environment.Is64BitProcess Then
            If My.Resources.ResourceManager.GetString("No64BitProcess_" & LanguageForLocalisedStrings) IsNot Nothing Then
                MessageBox.Show(My.Resources.ResourceManager.GetString("No64BitProcess_" & LanguageForLocalisedStrings), My.Resources.ResourceManager.GetString("Error_" & LanguageForLocalisedStrings), MessageBoxButtons.OK, MessageBoxIcon.Error)
            Else
                MessageBox.Show(My.Resources.ResourceManager.GetString("No64BitProcess_en"), My.Resources.ResourceManager.GetString("Error_en"), MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
            Application.Exit()
            Exit Sub
        End If

        ' Check if program has already been installed
        If FileIO.FileSystem.DirectoryExists(ProgDir) Then
            If My.Resources.ResourceManager.GetString("AlreadyInstalled_" & LanguageForLocalisedStrings) IsNot Nothing Then
                MessageBox.Show(My.Resources.ResourceManager.GetString("AlreadyInstalled_" & LanguageForLocalisedStrings), My.Resources.ResourceManager.GetString("Warning_" & LanguageForLocalisedStrings), MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Else
                MessageBox.Show(My.Resources.ResourceManager.GetString("AlreadyInstalled_en"), My.Resources.ResourceManager.GetString("Warning_en"), MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            End If
            Application.Exit()
            Exit Sub
        End If

        ' Wait 3 seconds before showing Welcome screen
        System.Threading.Thread.Sleep(3000)

        ' Show SetupWelcome screen now
        Me.Hide()
        SetupWelcome.Show()
    End Sub

    Private Sub SetupWizard_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        ' Delete temporary files, if existing, when closing program -> delete program folder in LocalAppData 
        If FileIO.FileSystem.DirectoryExists(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) & "\Moha_Logiciels") Then
            FileIO.FileSystem.DeleteDirectory(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) & "\Moha_Logiciels", FileIO.DeleteDirectoryOption.DeleteAllContents)
        End If
    End Sub
End Class
