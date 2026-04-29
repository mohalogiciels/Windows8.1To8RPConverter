Imports Microsoft.Win32
Imports System.Runtime.InteropServices
Imports System.Xml

Public Class MainProgram
    Dim Arguments As IReadOnlyCollection(Of String) = My.Application.CommandLineArgs
    Dim ArgumentsUpperCase As New List(Of String)
    Dim ProgramFiles As String = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles)
    Dim ProgDir As String = ProgramFiles & "\Windows 8.1 to 8 RP Converter"
    Dim WinDir As String = Environment.GetFolderPath(Environment.SpecialFolder.Windows)
    ' Import User32.dll to project -> disable close button
    Private Const MF_BYPOSITION As Integer = &H400
    <DllImport("User32")>
    Private Shared Function RemoveMenu(HMenu As IntPtr, NPosition As Integer, WFlags As Integer) As Integer
    End Function
    <DllImport("User32")>
    Private Shared Function GetSystemMenu(HWnd As IntPtr, BRevert As Boolean) As IntPtr
    End Function
    <DllImport("User32")>
    Private Shared Function GetMenuItemCount(HWnd As IntPtr) As Integer
    End Function

    Private Sub MainProgram_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Disable close button on the top right
        Dim HMenu As IntPtr = GetSystemMenu(Me.Handle, False)
        RemoveMenu(HMenu, GetMenuItemCount(HMenu) - 1, MF_BYPOSITION)
        HMenu = Nothing

        ' Add commands as upper-case in List ArgumentsUpperCase
        For Each Argument As String In Arguments
            ArgumentsUpperCase.Add(Argument.ToUpper)
        Next

        If ArgumentsUpperCase.Contains("/HELP") Or ArgumentsUpperCase.Contains("/?") Then
            Me.Opacity = 0
            MessageBox.Show("This program runs with following arguments:" & vbCrLf & _
                            "/HELP or /?: Show an info window with all valid arguments." & vbCrLf & _
                            "/SETUP: Run this file after converter has been installed. It will always apply theme (without sound scheme) and open the ""firstrun.xht"" file in Internet Explorer." & vbCrLf & _
                            "/MODIFY: Run this file after converter has been modified. More than one argument needed (i.e. those under this argument)!" & vbCrLf & _
                            "/QUERO: Run ""modify.xht"" in IE when Quero installed. For both ""/SETUP"" and ""/MODIFY"" switches." & vbCrLf & _
                            "/APPLYTHEME: Apply the theme after reboot if required. For ""/MODIFY"" switch only." & vbCrLf & _
                            "/SOUNDS: Apply theme with sound scheme. For both ""/SETUP"" and ""/MODIFY"" switches.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Application.Exit()
            Exit Sub
        End If

        If Arguments.Count = 0 Then
            Me.Opacity = 0
            MessageBox.Show("This program should be opened with at least one argument!" & vbCrLf & "Run this program with ""/HELP"" or ""/?"" to view list of arguments.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Application.Exit()
            Exit Sub
        End If

        If Not (ArgumentsUpperCase.Contains("/SETUP") Or ArgumentsUpperCase.Contains("/MODIFY")) Then
            Me.Opacity = 0
            MessageBox.Show("This program should be opened with at least one valid argument!" & vbCrLf & "Run this program with ""/HELP"" or ""/?"" to view list of arguments.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Application.Exit()
            Exit Sub
        End If

        If ArgumentsUpperCase.Contains("/MODIFY") And Not (ArgumentsUpperCase.Contains("/APPLYTHEME") Or ArgumentsUpperCase.Contains("/QUERO")) Then
            Me.Opacity = 0
            MessageBox.Show("Not sufficient commands to run this program!" & vbCrLf & "Run this program with ""/HELP"" or ""/?"" to view list of arguments.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Application.Exit()
            Exit Sub
        End If
    End Sub

    Private Sub MainProgram_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown
        InitialiseProgressBar()
        Run()
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
        If ArgumentsUpperCase.Contains("/SETUP") Or (ArgumentsUpperCase.Contains("/MODIFY") And ArgumentsUpperCase.Contains("/APPLYTHEME")) Then
            ProgBarCounter += 10
        End If
        ' If firstrun document should be opened
        If ArgumentsUpperCase.Contains("/SETUP") Or (ArgumentsUpperCase.Contains("/MODIFY") And ArgumentsUpperCase.Contains("/QUERO")) Then
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

    Private Sub Run()
        ' Apply theme to system
        If ArgumentsUpperCase.Contains("/SETUP") Or (ArgumentsUpperCase.Contains("/MODIFY") And ArgumentsUpperCase.Contains("/APPLYTHEME")) Then
            StatusLabel.Text = "Applying theme to system..."
            StatusLabel.Refresh()
            Using ApplyTheme As New Process
                If ArgumentsUpperCase.Contains("/SOUNDS") Then
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
        If ArgumentsUpperCase.Contains("/SETUP") Or (ArgumentsUpperCase.Contains("/MODIFY") And ArgumentsUpperCase.Contains("/QUERO")) Then
            StatusLabel.Text = "Open firstrun file in Internet Explorer..."
            StatusLabel.Refresh()
            Using RunFirstrunInIeProcess As New Process
                With RunFirstrunInIeProcess
                    .StartInfo.FileName = ProgramFiles & "\Internet Explorer\iexplore.exe"
                    If ArgumentsUpperCase.Contains("/SETUP") Then
                        .StartInfo.Arguments = """" & ProgDir & "\firstrun.xht"""
                    ElseIf ArgumentsUpperCase.Contains("/MODIFY") Then
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
        Application.Exit()
    End Sub
End Class
