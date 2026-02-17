Imports Microsoft.Win32
Imports System.Xml

Public Class MainProgram
    Dim Arguments As String() = Command.ToLower.Split(" ")
    Dim ProgramFiles As String = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles)
    Dim ProgDir As String = ProgramFiles & "\Windows 8.1 to 8 RP Converter"
    Dim WinDir As String = Environment.GetFolderPath(Environment.SpecialFolder.Windows)

    Private Sub MainProgram_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If Arguments.Length = 0 Or Not (Arguments.Contains("/setup") Or Arguments.Contains("/modify")) Then
            MessageBox.Show("This program should be opened with at least one argument!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Application.Exit()
        End If

        If Arguments.Count = 1 And Arguments.Contains("/modify") Then
            MessageBox.Show("Not sufficient commands to run this program!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Application.Exit()
        End If
    End Sub

    Private Sub MainProgram_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown
        InitialiseProgressBar()
        ' /setup: Run after setup converter
        ' /modify: Run after modify converter
        ' /quero: Run modify.xht in IE when Quero installed with modify
        ' /applytheme: Apply the theme after reboot if required
        ' /sounds: Apply theme with sound scheme
        RunToolsAndFiles(Arguments)
    End Sub

    Private Sub Wait(ByVal Seconds As Integer)
        For Index As Integer = 0 To Seconds * 100
            System.Threading.Thread.Sleep(10)
            Application.DoEvents()
        Next
    End Sub

    Private Sub InitialiseProgressBar()
        Dim ProgBarCounter As Integer = 0
        ' If theme should be applied
        If Arguments.Contains("/setup") Or (Arguments.Contains("/modify") And Arguments.Contains("/applytheme")) Then
            ProgBarCounter += 10
        End If
        ' If firstrun document should be opened
        If Arguments.Contains("/setup") Or (Arguments.Contains("/modify") And Arguments.Contains("/quero")) Then
            ProgBarCounter += 10
        End If

        ' Set progress bar values
        With StatusProgressBar
            .Style = ProgressBarStyle.Continuous
            .Value = 0
            .Minimum = 0
            .Maximum = ProgBarCounter
        End With
    End Sub

    Private Sub RunToolsAndFiles(ByVal ArgumentsArray As String())
        ' Apply theme to system
        If ArgumentsArray.Contains("/setup") Or (ArgumentsArray.Contains("/modify") And ArgumentsArray.Contains("/applytheme")) Then
            StatusLabel.Text = "Applying theme to system..."
            StatusLabel.Refresh()
            Using ApplyTheme As New Process
                If ArgumentsArray.Contains("/sounds") Then
                    With ApplyTheme
                        .StartInfo.FileName = WinDir & "\Resources\Themes\aerorp_sounds.theme"
                        .Start()
                    End With
                Else
                    With ApplyTheme
                        .StartInfo.FileName = WinDir & "\Resources\Themes\aerorp.theme"
                        .Start()
                    End With
                End If
            End Using
            StatusProgressBar.PerformStep()
            StatusProgressBar.Refresh()
            Wait(5)
        End If

        ' Run firstrun.xht/modify.xht document in Internet Explorer
        If ArgumentsArray.Contains("/setup") Or (ArgumentsArray.Contains("/modify") And ArgumentsArray.Contains("/quero")) Then
            StatusLabel.Text = "Open firstrun file in Internet Explorer..."
            StatusLabel.Refresh()
            Using RunFirstrunInIeProcess As New Process
                With RunFirstrunInIeProcess
                    .StartInfo.FileName = ProgramFiles & "\Internet Explorer\iexplore.exe"
                    If ArgumentsArray.Contains("/setup") Then
                        .StartInfo.Arguments = """" & ProgDir & "\firstrun.xht"""
                    ElseIf ArgumentsArray.Contains("/modify") Then
                        .StartInfo.Arguments = """" & ProgDir & "\modify.xht"""
                    End If
                    .Start()
                End With
            End Using
            StatusProgressBar.PerformStep()
            StatusProgressBar.Refresh()
            Wait(2)
        End If

        ' Finished
        StatusLabel.Text = "Setup has been finished!"
        StatusLabel.Refresh()
        Wait(2)
        Me.Close()
    End Sub
End Class
