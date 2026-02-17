Imports Microsoft.Win32
Imports System.Xml

Public Class MainProgram
    Dim Commands As String() = Command.Split(" ")
    Dim ProgramFiles As String = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles)
    Dim ProgDir As String = ProgramFiles & "\Windows 8.1 to 8 RP Converter"
    Dim WinDir As String = Environment.GetFolderPath(Environment.SpecialFolder.Windows)

    Private Sub MainProgram_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If Command.Length = 0 Then
            MessageBox.Show("This program should be opened with at least one argument!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Me.Close()
        End If

        If Commands.Count = 1 And Commands.Contains("/modify") Then
            MessageBox.Show("Not sufficient commands to run this program!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Me.Close()
        End If
    End Sub

    Private Sub MainProgram_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown
        InitialiseProgressBar()
        ' /setup: Run after setup converter
        ' /modify: Run after modify converter
        ' /quero: Run modify.xht in IE when Quero installed with modify
        ' /applytheme: Apply the theme after reboot if required
        ' /sounds: Apply theme with sound scheme
        RunToolsAndFiles(Commands)
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
        If Commands.Contains("/setup") Or (Commands.Contains("/modify") And Commands.Contains("/applytheme")) Then
            ProgBarCounter += 10
        End If
        ' If firstrun document should be opened
        If Commands.Contains("/setup") Or (Commands.Contains("/modify") And Commands.Contains("/quero")) Then
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

    Private Sub RunToolsAndFiles(ByVal CommandArray As String())
        ' Apply theme to system
        If CommandArray.Contains("/setup") Or (CommandArray.Contains("/modify") And CommandArray.Contains("/applytheme")) Then
            StatusLabel.Text = "Applying theme to system..."
            StatusLabel.Refresh()
            Using ApplyTheme As New Process
                If CommandArray.Contains("/sounds") Then
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
        If CommandArray.Contains("/setup") Or (CommandArray.Contains("/modify") And CommandArray.Contains("/quero")) Then
            StatusLabel.Text = "Open firstrun file in Internet Explorer..."
            StatusLabel.Refresh()
            Using RunFirstrunInIeProcess As New Process
                With RunFirstrunInIeProcess
                    .StartInfo.FileName = ProgramFiles & "\Internet Explorer\iexplore.exe"
                    If CommandArray.Contains("/setup") Then
                        .StartInfo.Arguments = """" & ProgDir & "\firstrun.xht"""
                    ElseIf CommandArray.Contains("/modify") Then
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
