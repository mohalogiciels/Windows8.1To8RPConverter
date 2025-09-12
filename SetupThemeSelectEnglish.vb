Imports System.Globalization

Public Class SetupThemeSelectEnglish

    Private Sub SetupThemeSelectEnglish_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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

        ' If language choice doesn’t match display language
        If Not SetupProgram.Language = CultureInfo.CurrentUICulture.ToString Then
            SetupProgram.IsSelectedLanguageDisplayLanguage = False
        Else
            SetupProgram.IsSelectedLanguageDisplayLanguage = True
        End If

        ' If already chosen style, load setting
        If SetupProgram.AddressBarStyle = "blue" Then
            BlueAddBarRadioButton.Checked = True
        ElseIf SetupProgram.AddressBarStyle = "white" Then
            WhiteAddBarRadioButton.Checked = True
        End If
    End Sub

    Private Sub BlueAddBarRadioButton_CheckedChanged(sender As Object, e As EventArgs) Handles BlueAddBarRadioButton.CheckedChanged
        NextButton.Enabled = True
        SetupProgram.AddressBarStyle = "blue"
    End Sub

    Private Sub WhiteAddBarRadioButton_CheckedChanged(sender As Object, e As EventArgs) Handles WhiteAddBarRadioButton.CheckedChanged
        NextButton.Enabled = True
        SetupProgram.AddressBarStyle = "white"
    End Sub

    Private Sub NextButton_Click(sender As Object, e As EventArgs) Handles NextButton.Click
        Me.Dispose()
        SetupOptionsEnglish.Show()
    End Sub

    Private Sub CloseButton_Click(sender As Object, e As EventArgs) Handles CloseButton.Click
        Me.Close()
    End Sub

    Private Sub InfoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles InfoToolStripMenuItem.Click
        SetupToolStripInfo.ShowDialog()
    End Sub

    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()
    End Sub

    Private Sub LanguageToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DeutschDeutschlandToolStripMenuItem.Click, EnglishUnitedKingdomToolStripMenuItem.Click, EnglishUnitedStatesToolStripMenuItem.Click, EspanolEspanaToolStripMenuItem.Click, FrancaisFranceToolStripMenuItem.Click, ItalianoItaliaToolStripMenuItem.Click, NederlandsNederlandToolStripMenuItem.Click, _
        PortuguesBrasilToolStripMenuItem.Click, PortuguesPortugalToolStripMenuItem.Click, KoreanToolStripMenuItem.Click, RussianToolStripMenuItem.Click, UkrainianToolStripMenuItem.Click
        Dim SelectedLanguage = CType(sender, ToolStripMenuItem).Text
        'If SelectedLanguage = "Deutsch (Deutschland)" Then
        '    Me.Dispose()
        '    SetupProgram.Language = "de-DE"
        '    SetupWelcomeDeutsch.show()
        If SelectedLanguage = "English (United Kingdom)" Then
            Dim ShowMe = New SetupWelcomeEnglish
            SetupProgram.Language = "en-GB"
            If SetupProgram.IsSelectedLanguageDisplayLanguage = True Then
                CultureInfo.DefaultThreadCurrentUICulture = CultureInfo.GetCultureInfo(SetupProgram.Language)
            Else
                CultureInfo.DefaultThreadCurrentUICulture = CultureInfo.GetCultureInfo("en-US")
            End If
            ShowMe.Show()
            Me.Dispose()
        ElseIf SelectedLanguage = "English (United States)" Then
            Dim ShowMe = New SetupWelcomeEnglish
            SetupProgram.Language = "en-US"
            If SetupProgram.IsSelectedLanguageDisplayLanguage = True Then
                CultureInfo.DefaultThreadCurrentUICulture = CultureInfo.GetCultureInfo(SetupProgram.Language)
            Else
                CultureInfo.DefaultThreadCurrentUICulture = CultureInfo.GetCultureInfo("en-US")
            End If
            ShowMe.Show()
            Me.Dispose()
            'ElseIf SelectedLanguage = "Español (España)" Then
            '    Me.Dispose()
            '    SetupProgram.Language = "es-ES"
            '    SetupWelcomeEspanol.show()
            'ElseIf SelectedLanguage = "Français (France)" Then
            '    Me.Dispose()
            '    SetupProgram.Language = "fr-FR"
            '    SetupWelcomeFrancais.show()
            'ElseIf SelectedLanguage = "Italiano (Italia)" Then
            '    Me.Dispose()
            '    SetupProgram.Language = "it-IT"
            '    SetupWelcomeItaliano.show()
            'ElseIf SelectedLanguage = "Nederlands (Nederland)" Then
            '    Me.Dispose()
            '    SetupProgram.Language = "nl-NL"
            '    SetupWelcomeNederlands.show()
            'ElseIf SelectedLanguage = "Português (Brasil)" Then
            '    Me.Dispose()
            '    SetupProgram.Language = "pt-BR"
            '    SetupWelcomePortuguesBrasil.show()
            'ElseIf SelectedLanguage = "Português (Portugal)" Then
            '    Me.Dispose()
            '    SetupProgram.Language = "pt-PT"
            '    SetupWelcomePortuguesPortugal.show()
            'ElseIf SelectedLanguage = "한국어" Then
            '    Me.Dispose()
            '    SetupProgram.Language = "ko-KR"
            '    SetupWelcomeKorean.show()
            'ElseIf SelectedLanguage = "русский" Then
            '    Me.Dispose()
            '    SetupProgram.Language = "ru-RU"
            '    SetupWelcomeRussian.show()
            'ElseIf SelectedLanguage = "українська" Then
            '    Me.Dispose()
            '    SetupProgram.Language = "uk-UA"
            '    SetupWelcomeUkrainian.show()
        Else
            MessageBox.Show("Not available yet!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End If
    End Sub

    Private Sub SetupThemeSelectEnglish_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        If e.CloseReason = CloseReason.UserClosing Then
            If MessageBox.Show("Are you sure you want to cancel? This program will be closed.", "Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) = System.Windows.Forms.DialogResult.Yes Then
                SetupProgram.Close()
            ElseIf System.Windows.Forms.DialogResult.No Then
                e.Cancel = True
            End If
        End If
    End Sub
End Class