Public Class UninstallOptionsInfoRemoveOnlyThemeEnglish
    Dim ProgramList As New List(Of String)

    Private Sub UninstallOptionsInfoRemoveOnlyThemeEnglish_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If UninstallWizard.InstAeroGlass = True Then
            ProgramList.Add("Aero Glass for Windows 8")
        End If
        If UninstallWizard.InstUX <> "none" Then
            If UninstallWizard.InstUX = "UltraUX" Then
                ProgramList.Add("UltraUXThemePatcher")
            ElseIf UninstallWizard.InstUX = "UXTSB" Then
                ProgramList.Add("UXThemeSignatureBypass")
            End If
        End If
        If UninstallWizard.InstONE = True Then
            ProgramList.Add("OldNewExplorer")
        End If
        If UninstallWizard.InstQuero = True Then
            ProgramList.Add("Quero Toolbar")
        End If
        If UninstallWizard.InstSysFiles = True Then
            ProgramList.Add("UIRibbon.dll && UIRibbonRes.dll")
        End If
        If UninstallWizard.InstSounds = True Then
            ProgramList.Add("Sound scheme")
        End If
        If UninstallWizard.InstGadgets = True Then
            ProgramList.Add("8GadgetPack")
        End If
    End Sub

    Private Sub UninstallOptionsInfoRemoveOnlyThemeEnglish_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown
        Dim ProgListString As String = String.Empty
        For Each prog In ProgramList
            ProgListString &= prog & vbCrLf
        Next
        ProgramsLabel.Text = ProgListString
    End Sub

    Private Sub OKButton_Click(sender As Object, e As EventArgs) Handles OKButton.Click
        Me.Dispose()
    End Sub

    Private Sub UninstallOptionsInfoRemoveOnlyThemeEnglish_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        Me.Dispose()
    End Sub
End Class
