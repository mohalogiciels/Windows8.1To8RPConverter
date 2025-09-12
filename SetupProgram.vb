Imports System.IO
Imports System.Configuration
Imports System.Globalization
Imports Microsoft.Win32

Public Class SetupProgram
    ' Initialise program settings
    Public ProgSettings As New My.MySettings

    ' Load default program settings
    Public Language As String = ProgSettings.Lang
    Public AddressBarStyle As String = ProgSettings.AddressBarStyle
    Public BitnessSystem As String = ProgSettings.BitnessSystem
    Public InstONE As Boolean = ProgSettings.InstONE
    Public InstUX As String = ProgSettings.InstUX
    Public InstQuero As Boolean = ProgSettings.InstQuero
    Public InstSysFiles As Boolean = ProgSettings.InstSysFiles
    Public InstSounds As Boolean = ProgSettings.InstSounds
    Public InstGadgets As Boolean = ProgSettings.InstGadgets

    ' Check if Aero Glass is already installed
    Public InstAeroGlass As Boolean = False
    Public AeroGlassInstalled As Boolean = False

    ' Check if language equals system language
    Public IsSelectedLanguageDisplayLanguage As Boolean

    ' Variables for program
    Dim ProgramFiles = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles)
    Dim ProgDir = ProgramFiles & "\Windows8.1to8RPConv"

    Private Sub SetupProgram_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Check if OS is Windows 8.1
        Dim WindowsVersion As RegistryKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\Microsoft\Windows NT\CurrentVersion", False)
        If WindowsVersion.GetValue("CurrentBuild") <> "9600" Then
            If CultureInfo.CurrentUICulture.ThreeLetterWindowsLanguageName = "DEU" Then
                MessageBox.Show("Dieses Programm läuft nur unter Windows 8.1", "Nicht unterstütztes Betriebssystem", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ElseIf CultureInfo.CurrentUICulture.ThreeLetterWindowsLanguageName = "ENU" Or CultureInfo.CurrentUICulture.ThreeLetterWindowsLanguageName = "ENG" Then
                MessageBox.Show("This program can only run on Windows 8.1", "Not supported operating system", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ElseIf CultureInfo.CurrentUICulture.ThreeLetterWindowsLanguageName = "ESN" Then
                MessageBox.Show("Este programa solo se puede ejecutar en Windows 8.1", "Sistema operativo no compatible", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ElseIf CultureInfo.CurrentUICulture.ThreeLetterWindowsLanguageName = "FRA" Then
                MessageBox.Show("Ce logiciel ne peut être exécuté que sous Windows 8.1", "Système d’exploitation non compatible", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ElseIf CultureInfo.CurrentUICulture.ThreeLetterWindowsLanguageName = "ITA" Then
                MessageBox.Show("Questo programma può essere eseguito solo su Windows 8.1", "Sistema operativo non compatibile", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ElseIf CultureInfo.CurrentUICulture.ThreeLetterWindowsLanguageName = "NLD" Then
                MessageBox.Show("Dit programma kan alleen worden uitgevoerd onder Windows 8.1", "Niet compatibel besturingssysteem", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ElseIf CultureInfo.CurrentUICulture.ThreeLetterWindowsLanguageName = "PTB" Then
                MessageBox.Show("Esse programa só pode ser executado no Windows 8.1", "Sistema operacional não compatível", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ElseIf CultureInfo.CurrentUICulture.ThreeLetterWindowsLanguageName = "PTG" Then
                MessageBox.Show("Este programa só pode ser executado no Windows 8.1", "Sistema operativo não compatível", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ElseIf CultureInfo.CurrentUICulture.ThreeLetterWindowsLanguageName = "KOR" Then
                MessageBox.Show("이 프로그램은 Windows 8.1에서만 실행할 수 있습니다.", "호환되지 않는 운영 체제", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ElseIf CultureInfo.CurrentUICulture.ThreeLetterWindowsLanguageName = "RUS" Then
                MessageBox.Show("Эта программа может работать только под управлением Windows 8.1", "Не совместимая операционная система", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ElseIf CultureInfo.CurrentUICulture.ThreeLetterWindowsLanguageName = "UKR" Then
                MessageBox.Show("Ця програма може працювати лише на Windows 8.1", "Несумісна операційна система", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Else
                MessageBox.Show("This program can only run on Windows 8.1", "Not supported operating system", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
            WindowsVersion.Close()
            Me.Close()
        End If

        ' Check if 64-bit program started when system is 64-bit
        Dim Is64BitSystem As Boolean = Environment.Is64BitOperatingSystem
        Dim Is64BitProcess As Boolean = Environment.Is64BitProcess
        If Is64BitSystem <> Is64BitProcess Then
            If CultureInfo.CurrentUICulture.ThreeLetterWindowsLanguageName = "DEU" Then
                MessageBox.Show("Das 64-bit-Programm konnte nicht gestartet werden!", "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ElseIf CultureInfo.CurrentUICulture.ThreeLetterWindowsLanguageName = "ENU" Or CultureInfo.CurrentUICulture.ThreeLetterWindowsLanguageName = "ENG" Then
                MessageBox.Show("The 64-bit program could not be loaded!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ElseIf CultureInfo.CurrentUICulture.ThreeLetterWindowsLanguageName = "ESN" Then
                MessageBox.Show("¡No se pudo iniciar el programa de 64 bits!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ElseIf CultureInfo.CurrentUICulture.ThreeLetterWindowsLanguageName = "FRA" Then
                MessageBox.Show("Ce logiciel 64 bits ne peut être pas lancé sur cet ordinateur !", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ElseIf CultureInfo.CurrentUICulture.ThreeLetterWindowsLanguageName = "ITA" Then
                MessageBox.Show("Il programma a 64 bit non può essere avviato!", "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ElseIf CultureInfo.CurrentUICulture.ThreeLetterWindowsLanguageName = "NLD" Then
                MessageBox.Show("Het 64-bits programma kan niet worden gestart!", "Fout", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ElseIf CultureInfo.CurrentUICulture.ThreeLetterWindowsLanguageName = "PTB" Then
                MessageBox.Show("Não foi possível iniciar o programa de 64 bits!", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ElseIf CultureInfo.CurrentUICulture.ThreeLetterWindowsLanguageName = "PTG" Then
                MessageBox.Show("Não foi possível iniciar o programa de 64 bits!", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ElseIf CultureInfo.CurrentUICulture.ThreeLetterWindowsLanguageName = "KOR" Then
                MessageBox.Show("64비트 프로그램이 실행되지 않았습니다!", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ElseIf CultureInfo.CurrentUICulture.ThreeLetterWindowsLanguageName = "RUS" Then
                MessageBox.Show("Не удалось запустить 64-битную программу!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ElseIf CultureInfo.CurrentUICulture.ThreeLetterWindowsLanguageName = "UKR" Then
                MessageBox.Show("64-бітну програму не вдалося запустити!", "Помилка", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Else
                MessageBox.Show("The 64-bit program could not be loaded!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
            Me.Close()
        End If

        ' Check if program has already been installed
        If FileIO.FileSystem.DirectoryExists(ProgDir) Then
            If CultureInfo.CurrentUICulture.ThreeLetterWindowsLanguageName = "DEU" Then
                MessageBox.Show("Dieses Programm wurde bereits installiert! Um es zu ändern oder zu entfernen, öffne die Systemsteuerung, wähle dieses Programm aus und klicke dann auf Ändern oder Deinstallieren.", "Warnung", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            ElseIf CultureInfo.CurrentUICulture.ThreeLetterWindowsLanguageName = "ENG" Or CultureInfo.CurrentUICulture.ThreeLetterWindowsLanguageName = "ENU" Then
                MessageBox.Show("This program has already been installed! To modify or remove it, open the Control Panel, choose this program from the list of installed programs and click on Modify or Uninstall.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            ElseIf CultureInfo.CurrentUICulture.ThreeLetterWindowsLanguageName = "ESN" Then
                MessageBox.Show("¡Este programa ya está instalado! Para modificarlo o desinstalarlo, inicia al Panel de control, después seleccione este programa en la lista de programas instalados y haga clic en Modificar o Desinstalar.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            ElseIf CultureInfo.CurrentUICulture.ThreeLetterWindowsLanguageName = "FRA" Then
                MessageBox.Show("Ce logiciel a déjà été installé sur cet ordinateur ! Pour faire des modifications ou le désinstaller, veuillez lancer le Panneau de configuration pour après choisir l’entrée de ce logiciel dans la liste, et cliquez sur Modifier ou Désinstaller.", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            ElseIf CultureInfo.CurrentUICulture.ThreeLetterWindowsLanguageName = "ITA" Then
                MessageBox.Show("Questo programma è già stato installato! Per modificarlo o rimuoverlo, aprire il Pannello di controllo e scegliere questo programma dall'elenco dei programmi installati, quindi fare clic su Modifica o Disinstalla.", "Avvertenze", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            ElseIf CultureInfo.CurrentUICulture.ThreeLetterWindowsLanguageName = "NLD" Then
                MessageBox.Show("Dit programma is al geïnstalleerd! Om het te wijzigen of te verwijderen, open je het Configuratiescherm en kies je dit programma uit de lijst met geïnstalleerde programma's en klik je op Wijzigen of Verwijderen.", "Waarschuwing", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            ElseIf CultureInfo.CurrentUICulture.ThreeLetterWindowsLanguageName = "PTB" Then
                MessageBox.Show("Esse programa já foi instalado! Para modificá-lo ou removê-lo, abra o Painel de controle, escolha esse programa na lista de programas instalados e clique em Modificar ou Desinstalar.", "Advertência", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            ElseIf CultureInfo.CurrentUICulture.ThreeLetterWindowsLanguageName = "PTG" Then
                MessageBox.Show("Este programa já foi instalado! Para o modificar ou remover, abra o Painel de Controlo e escolha este programa na lista de programas instalados, e clique em Modificar ou Desinstalar.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            ElseIf CultureInfo.CurrentUICulture.ThreeLetterWindowsLanguageName = "KOR" Then
                MessageBox.Show("이 프로그램이 이미 설치되었습니다! 수정하거나 제거하려면 제어판을 열고 설치된 프로그램 목록에서 이 프로그램을 선택한 다음 수정 또는 제거를 클릭합니다.", "경고", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            ElseIf CultureInfo.CurrentUICulture.ThreeLetterWindowsLanguageName = "RUS" Then
                MessageBox.Show("Эта программа уже установлена! Чтобы изменить или удалить ее, откройте Панель управления, выберите эту программу в списке установленных программ и нажмите на кнопку Изменить или Удалить.", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            ElseIf CultureInfo.CurrentUICulture.ThreeLetterWindowsLanguageName = "UKR" Then
                MessageBox.Show("Ця програма вже встановлена! Щоб змінити або видалити її, відкрийте Панель керування, виберіть цю програму зі списку встановлених програм і натисніть кнопку Змінити або Видалити.", "Попередження", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Else
                MessageBox.Show("This program has already been installed! To modify or remove it, open the Control Panel, choose this program from the list of installed programs and click on Modify or Uninstall.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            End If
            Me.Close()
        End If
    End Sub

    Private Sub SetupProgram_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown
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
        If LangSelectCB.SelectedIndex = 0 Then
            Language = "de-DE"
            Me.Text = "Sprache auswählen"
            LangSelectLabel.Text = "Bitte wähle deine Sprache aus:"
            OKButton.Text = "OK"
            CloseButton.Text = "Abbrechen"
        ElseIf LangSelectCB.SelectedIndex = 1 Then
            Language = "en-GB"
            Me.Text = "Select your language"
            LangSelectLabel.Text = "Please select your language:"
            OKButton.Text = "OK"
            CloseButton.Text = "Cancel"
        ElseIf LangSelectCB.SelectedIndex = 2 Then
            Language = "en-US"
            Me.Text = "Select your language"
            LangSelectLabel.Text = "Please select your language:"
            OKButton.Text = "OK"
            CloseButton.Text = "Cancel"
        ElseIf LangSelectCB.SelectedIndex = 3 Then
            Language = "es-ES"
            Me.Text = "Seleccionar el idioma"
            LangSelectLabel.Text = "Por favor, seleccione su idioma:"
            OKButton.Text = "Aceptar"
            CloseButton.Text = "Cancelar"
        ElseIf LangSelectCB.SelectedIndex = 4 Then
            Language = "fr-FR"
            Me.Text = "Sélectionner la langue"
            LangSelectLabel.Text = "Veuillez sélectionner la langue :"
            OKButton.Text = "OK"
            CloseButton.Text = "Annuler"
        ElseIf LangSelectCB.SelectedIndex = 5 Then
            Language = "it-IT"
            Me.Text = "Selezionare la lingua"
            LangSelectLabel.Text = "Selezionare la lingua desiderata:"
            OKButton.Text = "OK"
            CloseButton.Text = "Annulla"
        ElseIf LangSelectCB.SelectedIndex = 6 Then
            Language = "nl-NL"
            Me.Text = "Selecteer uw taal"
            LangSelectLabel.Text = "Selecteer je taal:"
            OKButton.Text = "OK"
            CloseButton.Text = "Annuleren"
        ElseIf LangSelectCB.SelectedIndex = 7 Then
            Language = "pt-BR"
            Me.Text = "Selecione seu idioma"
            LangSelectLabel.Text = "Selecione seu idioma:"
            OKButton.Text = "OK"
            CloseButton.Text = "Cancelar"
        ElseIf LangSelectCB.SelectedIndex = 8 Then
            Language = "pt-PT"
            Me.Text = "Selecione a sua língua"
            LangSelectLabel.Text = "Por favor, selecione a sua língua:"
            OKButton.Text = "OK"
            CloseButton.Text = "Cancelar"
        ElseIf LangSelectCB.SelectedIndex = 9 Then
            Language = "ko-KR"
            Me.Text = "언어 선택"
            LangSelectLabel.Text = "언어를 선택해 주세요:"
            OKButton.Text = "확인"
            CloseButton.Text = "취소"
        ElseIf LangSelectCB.SelectedIndex = 10 Then
            Language = "ru-RU"
            Me.Text = "Выберите язык"
            LangSelectLabel.Text = "Пожалуйста, выберите язык:"
            OKButton.Text = "ОК"
            CloseButton.Text = "Отмена"
        ElseIf LangSelectCB.SelectedIndex = 11 Then
            Language = "uk-UA"
            Me.Text = "Виберіть свою мову"
            LangSelectLabel.Text = "Будь ласка, оберіть мову:"
            OKButton.Text = "ОК"
            CloseButton.Text = "Скасувати"
        End If
    End Sub

    Private Sub OKButton_Click(sender As Object, e As EventArgs) Handles OKButton.Click
        ' If language choice does not match display language
        If Language <> CultureInfo.CurrentUICulture.ToString Then
            IsSelectedLanguageDisplayLanguage = False
            CultureInfo.DefaultThreadCurrentUICulture = CultureInfo.GetCultureInfo("en-US")
        Else
            IsSelectedLanguageDisplayLanguage = True
            CultureInfo.DefaultThreadCurrentUICulture = CultureInfo.GetCultureInfo(Language)
        End If

        ' Load program in selected language
        ' If Language = "de-DE" Then
        '    Me.Hide()
        '    SetupWelcomeDeutsch.show()
        If Language = "en-GB" Or Language = "en-US" Then
            Me.Hide()
            SetupWelcomeEnglish.Show()
            'ElseIf Language = "es-ES" Then
            '    Me.Hide()
            '    SetupWelcomeEspanol.show()
            'ElseIf Language = "fr-FR" Then
            '    Me.Hide()
            '    SetupWelcomeFrancais.show()
            'ElseIf Language = "it-IT" Then
            '    Me.Hide()
            '    SetupWelcomeItaliano.show()
            'ElseIf Language = "nl-NL" Then
            '    Me.Hide()
            '    SetupWelcomeNederlands.show()
            'ElseIf Language = "pt-BR" Then
            '    Me.Hide()
            '    SetupWelcomePortuguesBrasil.show()
            'ElseIf Language = "pt-PT" Then
            '    Me.Hide()
            '    SetupWelcomePortuguesPortugal.show()
            'ElseIf Language = "ko-KR" Then
            '    Me.Hide()
            '    SetupWelcomeKorean.show()
            'ElseIf Language = "ru-RU" Then
            '    Me.Hide()
            '    SetupWelcomeRussian.show()
            'ElseIf Language = "uk-UA" Then
            '    Me.Hide()
            '    SetupWelcomeUkrainian.show()
        Else
            ' !!temporary code!!
            MessageBox.Show("Not available yet!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End If
    End Sub

    Private Sub CancelButton_Click(sender As Object, e As EventArgs) Handles CloseButton.Click
        Me.Close()
    End Sub

    Private Sub SetupProgram_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        ' Delete temporary files, if existing, when closing program -> delete program folder in LocalAppData 
        Dim LocalAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData)
        If FileIO.FileSystem.DirectoryExists(LocalAppData & "\SetupWin8RPConverter") Then
            FileIO.FileSystem.DeleteDirectory(LocalAppData & "\SetupWin8RPConverter", FileIO.DeleteDirectoryOption.DeleteAllContents)
        End If
    End Sub
End Class
