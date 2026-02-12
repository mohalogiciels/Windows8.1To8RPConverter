Imports System.Globalization
Imports System.IO
Imports System.Xml

Public Class ModifyWizard
    ' Variables for settings read from file config.xml
    Public ConfigFile As New XmlDocument
    Public Shared ProgLanguage As String
    Public AddressBarStyle As String
    Public Shared BitnessSystem As String
    Public InstONE As Boolean
    Public InstUX As String
    Public InstQuero As Boolean
    Public InstSysFiles As Boolean
    Public InstSounds As Boolean
    Public InstGadgets As Boolean
    Public Shared Inst7TaskTw As Boolean

    ' Selected modify options
    Public ChangeAddBarStyle As Boolean
    Public ChangeToSelectedStyle As String
    Public ChangeONEChecked As Boolean
    Public Shared ModifyInstallONE As Boolean
    Public ChangeUXMethodChecked As Boolean
    Public Shared ModifyChangeToUXMethod As String
    Public ChangeQueroChecked As Boolean
    Public Shared ModifyInstallQuero As Boolean
    Public ChangeSysFilesChecked As Boolean
    Public Shared ModifyInstallSysFiles As Boolean
    Public InstallSoundsChecked As Boolean
    Public ChangeGadgetsChecked As Boolean
    Public Shared ModifyInstallGadgets As Boolean
    Public Change7TaskTwChecked As Boolean
    Public Shared ModifyInstall7TaskTw As Boolean

    ' Variables for program
    Public Shared ProgDir As String = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles) & "\Windows 8.1 to 8 RP Converter"
    Public Shared LanguageForLocalisedStrings As String
    Dim LanguageBeforeLoadForLocalisedStrings As String
    Public IsSelectedLanguageDisplayLanguage As Boolean = True
    Public IsThemeChangeNeeded As Boolean = False
    Public IsRestartNeeded As Boolean = False

    Private Sub Wait(ByVal SecondsToWait As Integer)
        For Index As Integer = 0 To SecondsToWait * 100
            System.Threading.Thread.Sleep(10)
            Application.DoEvents()
        Next
    End Sub

    Private Sub ModifyWizard_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Set language of current UI before loading language from config.xml file
        If CultureInfo.CurrentUICulture.Name.StartsWith("en") Then
            LanguageBeforeLoadForLocalisedStrings = "en"
        Else
            LanguageBeforeLoadForLocalisedStrings = CultureInfo.CurrentUICulture.Name.Replace("-", "_")
        End If

        ' Change window title and ActionLabel depending on system language
        If Not CultureInfo.CurrentUICulture.Name.StartsWith("en") Then
            Me.Text = My.Resources.ResourceManager.GetString("LoadingScreenTitle_" & LanguageBeforeLoadForLocalisedStrings)
            ActionLabel.Text = My.Resources.ResourceManager.GetString("LoadingScreenLabel_" & LanguageBeforeLoadForLocalisedStrings)
        End If

    End Sub

    Private Sub ModifyWizard_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown
        ' Check if 64-bit program started when system is 64-bit
        If Environment.Is64BitOperatingSystem <> Environment.Is64BitProcess Then
            MessageBox.Show(My.Resources.ResourceManager.GetString("No64BitProcess_" & LanguageBeforeLoadForLocalisedStrings), My.Resources.ResourceManager.GetString("Error_" & LanguageBeforeLoadForLocalisedStrings), MessageBoxButtons.OK, MessageBoxIcon.Error)
            Me.Close()
        End If

        ' Check if exists program folder
        If Not FileIO.FileSystem.DirectoryExists(ProgDir) Then
            MessageBox.Show(My.Resources.ResourceManager.GetString("ProgramFolderNotFound_" & LanguageBeforeLoadForLocalisedStrings), My.Resources.ResourceManager.GetString("Error_" & LanguageBeforeLoadForLocalisedStrings), MessageBoxButtons.OK, MessageBoxIcon.Error)
            Me.Close()
        End If

        ' Check if ResFolder exists
        If Not FileIO.FileSystem.DirectoryExists(ProgDir & "\res") Then
            MessageBox.Show(My.Resources.ResourceManager.GetString("ResourceFolderNotFound_" & LanguageBeforeLoadForLocalisedStrings), My.Resources.ResourceManager.GetString("Error_" & LanguageBeforeLoadForLocalisedStrings), MessageBoxButtons.OK, MessageBoxIcon.Error)
            Me.Close()
        End If

        ' Check if file config.xml exists, then load settings
        If FileIO.FileSystem.FileExists(ProgDir & "\config.xml") Then
            LoadSettingsFromXmlFile(ProgDir & "\config.xml")
        Else
            MessageBox.Show(My.Resources.ResourceManager.GetString("ConfigFileNotFound_" & LanguageBeforeLoadForLocalisedStrings), My.Resources.ResourceManager.GetString("Error_" & LanguageBeforeLoadForLocalisedStrings), MessageBoxButtons.OK, MessageBoxIcon.Error)
            Me.Close()
        End If

        ' Check if chosen language matches display language
        If ProgLanguage <> CultureInfo.CurrentUICulture.Name Then
            IsSelectedLanguageDisplayLanguage = False
        End If

        ' Set the LanguageForLocalisedStrings
        If ProgLanguage.StartsWith("en") Then
            LanguageForLocalisedStrings = "en"
        Else
            LanguageForLocalisedStrings = ProgLanguage.Replace("-", "_")
        End If

        ' Show loading screen for 1 second
        Wait(1)

        ' Go to next step -> Show theme select window
        Me.Hide()
        ModifyThemeSelect.Show()
    End Sub

    Private Sub LoadSettingsFromXmlFile(ByVal XmlFile As String)
        ConfigFile.Load(XmlFile)
        ProgLanguage = ConfigFile.SelectSingleNode("//setting[@name='ProgLanguage']/value").InnerText
        AddressBarStyle = ConfigFile.SelectSingleNode("//setting[@name='AddressBarStyle']/value").InnerText
        BitnessSystem = ConfigFile.SelectSingleNode("//setting[@name='BitnessSystem']/value").InnerText
        InstONE = ConfigFile.SelectSingleNode("//setting[@name='InstONE']/value").InnerText
        InstUX = ConfigFile.SelectSingleNode("//setting[@name='InstUX']/value").InnerText
        InstQuero = ConfigFile.SelectSingleNode("//setting[@name='InstQuero']/value").InnerText
        InstSysFiles = ConfigFile.SelectSingleNode("//setting[@name='InstSysFiles']/value").InnerText
        InstSounds = ConfigFile.SelectSingleNode("//setting[@name='InstSounds']/value").InnerText
        InstGadgets = ConfigFile.SelectSingleNode("//setting[@name='InstGadgets']/value").InnerText
        Inst7TaskTw = ConfigFile.SelectSingleNode("//setting[@name='Inst7TaskTw']/value").InnerText
    End Sub
End Class