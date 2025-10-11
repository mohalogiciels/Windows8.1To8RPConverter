Imports System.IO
Imports System.Xml
Imports System.Globalization

Public Class UninstallWizard
    ' Variables for settings read from programs.xml
    Public Language As String
    Public BitnessSystem As String
    Public InstAeroGlass As Boolean
    Public InstONE As Boolean
    Public InstUX As String
    Public InstQuero As Boolean
    Public InstSysFiles As Boolean
    Public InstSounds As Boolean
    Public InstGadgets As Boolean

    ' Program settings
    Public UninstAero As Boolean
    Public UninstONE As Boolean
    Public UninstUX As Boolean
    Public UninstQuero As Boolean
    Public UninstSysFiles As Boolean
    Public UninstSounds As Boolean
    Public UninstGadgets As Boolean

    ' Variable for program
    Dim ProgramFiles As String = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles)

    Private Sub Wait(ByVal seconds As Integer)
        For Index As Integer = 0 To seconds * 100
            System.Threading.Thread.Sleep(10)
            Application.DoEvents()
        Next
    End Sub

    Private Sub UninstallWizard_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown
        Dim ProgDir = ProgramFiles & "\Windows8.1to8RPConv"
        Dim ConfigFile As New XmlDocument

        ' Set .NET Framework language to English first
        CultureInfo.DefaultThreadCurrentUICulture = CultureInfo.GetCultureInfo("en-US")

        ' Change title and ActionLabel depending on system language
        If CultureInfo.CurrentUICulture.ThreeLetterWindowsLanguageName = "DEU" Then
            Me.Text = "Programm wird geladen"
            ActionLabel.Text = "Wird geladen..."
        ElseIf CultureInfo.CurrentUICulture.ThreeLetterWindowsLanguageName = "ESN" Then
            Me.Text = "Cargando programa"
            ActionLabel.Text = "Cargando..."
        ElseIf CultureInfo.CurrentUICulture.ThreeLetterWindowsLanguageName = "FRA" Then
            Me.Text = "Chargement du logiciel"
            ActionLabel.Text = "Chargement en cours..."
        ElseIf CultureInfo.CurrentUICulture.ThreeLetterWindowsLanguageName = "ITA" Then
            Me.Text = "Caricamento del programma"
            ActionLabel.Text = "Caricamento in corso..."
        ElseIf CultureInfo.CurrentUICulture.ThreeLetterWindowsLanguageName = "NLD" Then
            Me.Text = "Programma wordt geladen"
            ActionLabel.Text = "Nu aan het laden..."
        ElseIf CultureInfo.CurrentUICulture.ThreeLetterWindowsLanguageName = "PTB" Then
            Me.Text = "Carregando programa"
            ActionLabel.Text = "Carregando agora..."
        ElseIf CultureInfo.CurrentUICulture.ThreeLetterWindowsLanguageName = "PTG" Then
            Me.Text = "A carregar programa"
            ActionLabel.Text = "A carregar agora..."
        ElseIf CultureInfo.CurrentUICulture.ThreeLetterWindowsLanguageName = "KOR" Then
            Me.Text = "프로그램 로드 중"
            ActionLabel.Text = "지금 로드 중..."
        ElseIf CultureInfo.CurrentUICulture.ThreeLetterWindowsLanguageName = "RUS" Then
            Me.Text = "Загрузка программы"
            ActionLabel.Text = "Загрузка сейчас..."
        ElseIf CultureInfo.CurrentUICulture.ThreeLetterWindowsLanguageName = "UKR" Then
            Me.Text = "Завантаження програми"
            ActionLabel.Text = "Завантажується..."
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

        ' Check if exists program folder
        If Not FileIO.FileSystem.DirectoryExists(ProgDir) Then
            If CultureInfo.CurrentUICulture.ThreeLetterWindowsLanguageName = "DEU" Then
                MessageBox.Show("Programmordner konnte nicht gefunden werden!", "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ElseIf CultureInfo.CurrentUICulture.ThreeLetterWindowsLanguageName = "ENG" Or CultureInfo.CurrentUICulture.ThreeLetterWindowsLanguageName = "ENU" Then
                MessageBox.Show("Program folder cannot be found on this system!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ElseIf CultureInfo.CurrentUICulture.ThreeLetterWindowsLanguageName = "ESN" Then
                MessageBox.Show("¡La carpeta del programa no se encuentra en este sistema!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ElseIf CultureInfo.CurrentUICulture.ThreeLetterWindowsLanguageName = "FRA" Then
                MessageBox.Show("Le dossier du logiciel n’est pas disponible sur cet ordinateur !", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ElseIf CultureInfo.CurrentUICulture.ThreeLetterWindowsLanguageName = "ITA" Then
                MessageBox.Show("La cartella del programma non può essere trovata su questo sistema!", "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ElseIf CultureInfo.CurrentUICulture.ThreeLetterWindowsLanguageName = "NLD" Then
                MessageBox.Show("Programmamap kan niet worden gevonden op dit systeem!", "Fout", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ElseIf CultureInfo.CurrentUICulture.ThreeLetterWindowsLanguageName = "PTB" Then
                MessageBox.Show("A pasta do programa não pode ser encontrada neste sistema!", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ElseIf CultureInfo.CurrentUICulture.ThreeLetterWindowsLanguageName = "PTG" Then
                MessageBox.Show("A pasta do programa não pode ser encontrada neste sistema!", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ElseIf CultureInfo.CurrentUICulture.ThreeLetterWindowsLanguageName = "KOR" Then
                MessageBox.Show("이 시스템에서 프로그램 폴더를 찾을 수 없습니다!", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ElseIf CultureInfo.CurrentUICulture.ThreeLetterWindowsLanguageName = "RUS" Then
                MessageBox.Show("Папка программы не может быть найдена в этой системе!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ElseIf CultureInfo.CurrentUICulture.ThreeLetterWindowsLanguageName = "UKR" Then
                MessageBox.Show("Теку з програмою не може бути знайдено в цій системі!", "Помилка", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
            Me.Close()
        End If

        ' Check if exists programs.xml, then load settings
        If FileIO.FileSystem.FileExists(ProgDir & "\programs.xml") Then
            ConfigFile.Load(ProgDir & "\programs.xml")
            Language = ConfigFile.SelectSingleNode("//setting[@name='Lang']/value").InnerText
            BitnessSystem = ConfigFile.SelectSingleNode("//setting[@name='BitnessSystem']/value").InnerText
            InstAeroGlass = ConfigFile.SelectSingleNode("//setting[@name='InstAeroGlass']/value").InnerText
            InstONE = ConfigFile.SelectSingleNode("//setting[@name='InstONE']/value").InnerText
            InstUX = ConfigFile.SelectSingleNode("//setting[@name='InstUX']/value").InnerText
            InstQuero = ConfigFile.SelectSingleNode("//setting[@name='InstQuero']/value").InnerText
            InstSysFiles = ConfigFile.SelectSingleNode("//setting[@name='InstSysFiles']/value").InnerText
            InstSounds = ConfigFile.SelectSingleNode("//setting[@name='InstSounds']/value").InnerText
            InstGadgets = ConfigFile.SelectSingleNode("//setting[@name='InstGadgets']/value").InnerText
        ElseIf FileIO.FileSystem.DirectoryExists(ProgDir) And Not FileIO.FileSystem.FileExists(ProgDir & "\programs.xml") Then
            If CultureInfo.CurrentUICulture.ThreeLetterWindowsLanguageName = "DEU" Then
                MessageBox.Show("Konfigurationsdatei konnte nicht gefunden werden!", "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ElseIf CultureInfo.CurrentUICulture.ThreeLetterWindowsLanguageName = "ENG" Or CultureInfo.CurrentUICulture.ThreeLetterWindowsLanguageName = "ENU" Then
                MessageBox.Show("Config file cannot be found on this system!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ElseIf CultureInfo.CurrentUICulture.ThreeLetterWindowsLanguageName = "ESN" Then
                MessageBox.Show("¡El fichero de configuración no se encuentra en este sistema!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ElseIf CultureInfo.CurrentUICulture.ThreeLetterWindowsLanguageName = "FRA" Then
                MessageBox.Show("Le fichier de configuration n’est pas disponible sur cet ordinateur !", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ElseIf CultureInfo.CurrentUICulture.ThreeLetterWindowsLanguageName = "ITA" Then
                MessageBox.Show("Il file di configurazione non può essere trovato su questo sistema!", "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ElseIf CultureInfo.CurrentUICulture.ThreeLetterWindowsLanguageName = "NLD" Then
                MessageBox.Show("Configuratiebestand kan niet worden gevonden op dit systeem!", "Fout", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ElseIf CultureInfo.CurrentUICulture.ThreeLetterWindowsLanguageName = "PTB" Then
                MessageBox.Show("O arquivo de configuração não pode ser encontrado neste sistema!", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ElseIf CultureInfo.CurrentUICulture.ThreeLetterWindowsLanguageName = "PTG" Then
                MessageBox.Show("O ficheiro de configuração não pode ser encontrado neste sistema!", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ElseIf CultureInfo.CurrentUICulture.ThreeLetterWindowsLanguageName = "KOR" Then
                MessageBox.Show("이 시스템에서 구성 파일을 찾을 수 없습니다!", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ElseIf CultureInfo.CurrentUICulture.ThreeLetterWindowsLanguageName = "RUS" Then
                MessageBox.Show("Файл конфигурации не может быть найден в этой системе!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ElseIf CultureInfo.CurrentUICulture.ThreeLetterWindowsLanguageName = "UKR" Then
                MessageBox.Show("Конфігураційний файл не може бути знайдено в цій системі!", "Помилка", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
            Me.Close()
        End If

        ' Check if exists ResFolder
        If FileIO.FileSystem.DirectoryExists(ProgDir) And Not FileIO.FileSystem.DirectoryExists(ProgDir & "\res") Then
            If CultureInfo.CurrentUICulture.ThreeLetterWindowsLanguageName = "DEU" Then
                MessageBox.Show("Ressourcenordner konnte nicht gefunden werden!", "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ElseIf CultureInfo.CurrentUICulture.ThreeLetterWindowsLanguageName = "ENG" Or CultureInfo.CurrentUICulture.ThreeLetterWindowsLanguageName = "ENU" Then
                MessageBox.Show("Resources folder cannot be found on this system!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ElseIf CultureInfo.CurrentUICulture.ThreeLetterWindowsLanguageName = "ESN" Then
                MessageBox.Show("¡La carpeta de recursos no se encuentra en este sistema!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ElseIf CultureInfo.CurrentUICulture.ThreeLetterWindowsLanguageName = "FRA" Then
                MessageBox.Show("Le dossier de ressources n’est pas disponible sur cet ordinateur !", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ElseIf CultureInfo.CurrentUICulture.ThreeLetterWindowsLanguageName = "ITA" Then
                MessageBox.Show("La cartella delle risorse non può essere trovata su questo sistema!", "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ElseIf CultureInfo.CurrentUICulture.ThreeLetterWindowsLanguageName = "NLD" Then
                MessageBox.Show("De map resources kan niet worden gevonden op dit systeem!", "Fout", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ElseIf CultureInfo.CurrentUICulture.ThreeLetterWindowsLanguageName = "PTB" Then
                MessageBox.Show("A pasta de recursos não pode ser encontrada neste sistema!", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ElseIf CultureInfo.CurrentUICulture.ThreeLetterWindowsLanguageName = "PTG" Then
                MessageBox.Show("A pasta de recursos não pode ser encontrada neste sistema!", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ElseIf CultureInfo.CurrentUICulture.ThreeLetterWindowsLanguageName = "KOR" Then
                MessageBox.Show("이 시스템에서 리소스 폴더를 찾을 수 없습니다!", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ElseIf CultureInfo.CurrentUICulture.ThreeLetterWindowsLanguageName = "RUS" Then
                MessageBox.Show("Папка ресурсов не может быть найдена в этой системе!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ElseIf CultureInfo.CurrentUICulture.ThreeLetterWindowsLanguageName = "UKR" Then
                MessageBox.Show("Теку з ресурсами не вдається знайти в цій системі!", "Помилка", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
            Me.Close()
        End If

        ' Set .NET Framework language to match selected language, if matches display language
        If Language = CultureInfo.CurrentUICulture.ToString Then
            CultureInfo.DefaultThreadCurrentUICulture = CultureInfo.GetCultureInfo(Language)
        End If

        ' Show loading screen for 1 second
        Wait(1)

        'Check language setting
        'If Language = "de-DE" Then
        '    Me.Hide()
        '    UninstallWelcomeDeutsch.show()
        If Language = "en-US" Or Language = "en-GB" Then
            Me.Hide()
            UninstallWelcomeEnglish.Show()
            'ElseIf Language = "es-ES" Then
            '    Me.Hide()
            '    UninstallWelcomeEspanol.show()
            'ElseIf Language = "fr-FR" Then
            '    Me.Hide()
            '    UninstallWelcomeFrancais.show()
            'ElseIf Language = "it-IT" Then
            '    Me.Hide()
            '    UninstallWelcomeItaliano.show()
            'ElseIf Language = "nl-NL" Then
            '    Me.Hide()
            '    UninstallWelcomeNederlands.show()
            'ElseIf Language = "pt-BR" Then
            '    Me.Hide()
            '    UninstallWelcomePortuguesBrasil.show()
            'ElseIf Language = "pt-PT" Then
            '    Me.Hide()
            '    UninstallWelcomePortuguesPortugal.show()
            'ElseIf Language = "ko-KR" Then
            '    Me.Hide()
            '    UninstallWelcomeKorean.show()
            'ElseIf Language = "ru-RU" Then
            '    Me.Hide()
            '    UninstallWelcomeRussian.show()
            'ElseIf Language = "uk-UA" Then
            '    Me.Hide()
            '    UninstallWelcomeUkrainian.show()
        Else
            '!!temporary code!!
            Me.Close()
        End If
    End Sub
End Class
