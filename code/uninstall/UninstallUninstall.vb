Imports Microsoft.Win32
Imports System.ComponentModel
Imports System.IO
Imports System.Security.AccessControl
Imports System.Security.Principal

Public Class UninstallUninstall
    Private WithEvents BgWorker As BackgroundWorker
    Dim ProgramFiles = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles)
    Dim ProgramFilesX86 = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86)
    Dim WinDir = Environment.GetFolderPath(Environment.SpecialFolder.Windows)
    Dim ProgBarCounter As Integer = 0
    Dim IsFinished As Boolean = False
    Dim RestartNow As Boolean = False

    Private Sub UninstallUninstall_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown
        ' Set progress bar to marquee
        With ProgressBarUninst
            .Style = ProgressBarStyle.Marquee
            .MarqueeAnimationSpeed = 100
            .Refresh()
        End With

        ' Initialise progress bar
        InitialiseProgressBar()

        ' Initialise BackgroundWorker
        InitialiseBackgroundWorker()

        ' Change theme back to standard
        Me.TopMost = True
        BgWorker.RunWorkerAsync("ChangeToDefaultTheme")
        Wait(5)
        Me.TopMost = False
        UninstallLauncher()
    End Sub

    Private Sub Wait(ByVal SecondsToWait As Integer)
        For Index As Integer = 0 To SecondsToWait * 100
            System.Threading.Thread.Sleep(10)
            Application.DoEvents()
        Next
    End Sub

    Private Sub InitialiseBackgroundWorker()
        BgWorker = New BackgroundWorker
        BgWorker.WorkerSupportsCancellation = True
    End Sub

    Private Sub InitialiseProgressBar()
        ' Initialise progress bar
        If UninstallWizard.UninstAero = True Then
            ProgBarCounter += 10
        End If

        If UninstallWizard.UninstONE = True Then
            ProgBarCounter += 10
        End If

        If UninstallWizard.UninstUX = True Then
            ProgBarCounter += 10
        End If

        If UninstallWizard.UninstQuero = True Then
            ProgBarCounter += 10
        End If

        If UninstallWizard.UninstSysFiles = True Then
            ProgBarCounter += 10
        End If

        If UninstallWizard.UninstSounds = True Then
            ProgBarCounter += 10
        End If

        If UninstallWizard.UninstGadgets = True Then
            ProgBarCounter += 10
        End If

        If UninstallWizard.Uninst7TaskTw = True Then
            ProgBarCounter += 10
        End If
    End Sub

    Private Sub WaitUntilTaskFinishes()
        Do
            Application.DoEvents()
        Loop While BgWorker.IsBusy
    End Sub

    Private Sub BgWorker_DoWork(sender As Object, e As DoWorkEventArgs) Handles BgWorker.DoWork
        Dim WhatToDo As String = Convert.ToString(e.Argument)
        Select Case WhatToDo
            Case "ChangeToDefaultTheme"
                ' Change to standard theme
                Using ThemeChangeProcess As New Process
                    With ThemeChangeProcess
                        .StartInfo.FileName = WinDir & "\Resources\Themes\aero.theme"
                        .Start()
                    End With
                End Using
            Case "CloseExplorer"
                ' Close explorer.exe
                Using EndExplorerTask As New Process
                    With EndExplorerTask
                        .StartInfo.FileName = WinDir & "\System32\taskkill.exe"
                        .StartInfo.Arguments = "/f /im explorer.exe"
                        .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                        .Start()
                        .WaitForExit()
                    End With
                End Using
            Case "UninstAeroGlass"
                ' Get Aero Glass path from uninstall list in registry
                Dim AeroGlassPath As String = String.Empty
                Using UninstallReg As RegistryKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall", False)
                    Dim UninstallRegKeys As String() = UninstallReg.GetSubKeyNames
                    Dim UninstallAero As New List(Of String)
                    For Each RegKey In UninstallRegKeys
                        If RegKey.StartsWith("Aero Glass for Win8.1") Then
                            UninstallAero.Add(RegKey)
                        End If
                    Next

                    ' Check if Aero Glass folder exists
                    If UninstallAero.Count > 0 Then
                        Using UninstallAeroKey As RegistryKey = UninstallReg.OpenSubKey(UninstallAero(0))
                            AeroGlassPath = UninstallAeroKey.GetValue("InstallLocation")
                        End Using
                    End If
                    UninstallReg.Close()
                    UninstallAero = Nothing
                End Using

                If FileIO.FileSystem.DirectoryExists(AeroGlassPath) Then
                    Using AeroGlassUninst As New Process
                        With AeroGlassUninst
                            .StartInfo.FileName = AeroGlassPath & "unins000.exe"
                            .StartInfo.Arguments = "/SILENT /NORESTART /SUPPRESSMSGBOXES"
                            .Start()
                            .WaitForExit()
                        End With
                    End Using
                    If ProgressBarUninst.InvokeRequired Then
                        ProgressBarUninst.Invoke(Sub()
                                                     ProgressBarUninst.PerformStep()
                                                     ProgressBarUninst.Refresh()
                                                 End Sub)
                    Else
                        ProgressBarUninst.PerformStep()
                        ProgressBarUninst.Refresh()
                    End If

                    If ActionLabel.InvokeRequired Then
                        ActionLabel.Invoke(Sub()
                                               ActionLabel.Text = "Aero Glass has been removed!"
                                               ActionLabel.Refresh()
                                           End Sub)
                    Else
                        ActionLabel.Text = "Aero Glass has been removed!"
                        ActionLabel.Refresh()
                    End If
                Else
                    MessageBox.Show("Aero Glass could not be found on this system!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    If ActionLabel.InvokeRequired Then
                        ActionLabel.Text = "Aero Glass could not be uninstalled!"
                        ActionLabel.Refresh()
                    Else
                        ActionLabel.Text = "Aero Glass could not be uninstalled!"
                        ActionLabel.Refresh()
                    End If
                    Wait(3)
                End If
            Case "UninstONE"
                If FileIO.FileSystem.DirectoryExists(ProgramFiles & "\OldNewExplorer") Then
                    Using ONESetup As New Process
                        ' Unreg OldNewExplorer DLLs
                        If UninstallWizard.BitnessSystem = "64Bit" Then
                            With ONESetup
                                .StartInfo.FileName = WinDir & "\System32\regsvr32.exe"
                                .StartInfo.Arguments = "/u /s """ & ProgramFiles & "\OldNewExplorer\OldNewExplorer64.dll"""
                                .Start()
                                .WaitForExit()
                            End With
                        End If
                        With ONESetup
                            .StartInfo.FileName = WinDir & "\System32\regsvr32.exe"
                            .StartInfo.Arguments = "/u /s """ & ProgramFiles & "\OldNewExplorer\OldNewExplorer32.dll"""
                            .Start()
                            .WaitForExit()
                        End With
                    End Using

                    ' Remove desktop shortcut
                    If FileIO.FileSystem.FileExists(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) & "\OldNewExplorer Configuration.lnk") Then
                        FileIO.FileSystem.DeleteFile(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) & "\OldNewExplorer Configuration.lnk")
                    End If

                    If ProgressBarUninst.InvokeRequired Then
                        ProgressBarUninst.Invoke(Sub()
                                                     ProgressBarUninst.PerformStep()
                                                     ProgressBarUninst.Refresh()
                                                 End Sub)
                    Else
                        ProgressBarUninst.PerformStep()
                        ProgressBarUninst.Refresh()
                    End If

                    If ActionLabel.InvokeRequired Then
                        ActionLabel.Invoke(Sub()
                                               ActionLabel.Text = "OldNewExplorer has been removed!"
                                               ActionLabel.Refresh()
                                           End Sub)
                    Else
                        ActionLabel.Text = "OldNewExplorer has been removed!"
                        ActionLabel.Refresh()
                    End If
                Else
                    MessageBox.Show("OldNewExplorer could not be found on this system!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    If ActionLabel.InvokeRequired Then
                        ActionLabel.Invoke(Sub()
                                               ActionLabel.Text = "OldNewExplorer could not be removed!"
                                               ActionLabel.Refresh()
                                           End Sub)
                    Else
                        ActionLabel.Text = "OldNewExplorer could not be removed!"
                        ActionLabel.Refresh()
                    End If

                    If ProgressBarUninst.InvokeRequired Then
                        ProgressBarUninst.Invoke(Sub()
                                                     ProgressBarUninst.PerformStep()
                                                     ProgressBarUninst.Refresh()
                                                 End Sub)
                    Else
                        ProgressBarUninst.PerformStep()
                        ProgressBarUninst.Refresh()
                    End If
                End If
            Case "UninstUX"
                If UninstallWizard.InstUX = "UltraUX" Then
                    ' Uninstall UltraUX
                    If ActionLabel.InvokeRequired Then
                        ActionLabel.Invoke(Sub()
                                               ActionLabel.Text = "Uninstalling UltraUXThemePatcher..."
                                               ActionLabel.Refresh()
                                           End Sub)
                    Else
                        ActionLabel.Text = "Uninstalling UltraUXThemePatcher..."
                        ActionLabel.Refresh()
                    End If

                    If FileIO.FileSystem.FileExists(ProgramFilesX86 & "\UltraUXThemePatcher\Uninstall.exe") Or FileIO.FileSystem.FileExists(ProgramFiles & "\UltraUXThemePatcher\Uninstall.exe") Then
                        If UninstallWizard.BitnessSystem = "64Bit" Then
                            Using UninstUltraUX As New Process
                                With UninstUltraUX
                                    .StartInfo.FileName = ProgramFilesX86 & "\UltraUXThemePatcher\Uninstall.exe"
                                    .Start()
                                End With
                            End Using
                            ' WaitForExit for process called Un.exe because Uninstall.exe closes
                            Do
                                System.Threading.Thread.Sleep(1000)
                            Loop Until Process.GetProcessesByName("Un").Count = 0
                            Application.DoEvents()
                        ElseIf UninstallWizard.BitnessSystem = "32Bit" Then
                            Using UninstUltraUX As New Process
                                With UninstUltraUX
                                    .StartInfo.FileName = ProgramFiles & "\UltraUXThemePatcher\Uninstall.exe"
                                    .Start()
                                End With
                            End Using
                            ' WaitForExit for process called Un.exe because Uninstall.exe closes
                            Do
                                Application.DoEvents()
                            Loop Until Process.GetProcessesByName("Un").Count = 0
                        End If

                        If ProgressBarUninst.InvokeRequired Then
                            ProgressBarUninst.Invoke(Sub()
                                                         ProgressBarUninst.PerformStep()
                                                         ProgressBarUninst.Refresh()
                                                     End Sub)
                        Else
                            ProgressBarUninst.PerformStep()
                            ProgressBarUninst.Refresh()
                        End If

                        If ActionLabel.InvokeRequired Then
                            ActionLabel.Invoke(Sub()
                                                   ActionLabel.Text = "UltraUXThemePatcher has been uninstalled!"
                                                   ActionLabel.Refresh()
                                               End Sub)
                        Else
                            ActionLabel.Text = "UltraUXThemePatcher has been uninstalled!"
                            ActionLabel.Refresh()
                        End If
                    Else
                        MessageBox.Show("UltraUXThemePatcher could not be found on this system!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        If ActionLabel.InvokeRequired Then
                            ActionLabel.Invoke(Sub()
                                                   ActionLabel.Text = "UltraUXThemePatcher could not be removed!"
                                                   ActionLabel.Refresh()
                                               End Sub)
                        Else
                            ActionLabel.Text = "UltraUXThemePatcher could not be removed!"
                            ActionLabel.Refresh()
                        End If

                        If ProgressBarUninst.InvokeRequired Then
                            ProgressBarUninst.Invoke(Sub()
                                                         ProgressBarUninst.PerformStep()
                                                         ProgressBarUninst.Refresh()
                                                     End Sub)
                        Else
                            ProgressBarUninst.PerformStep()
                            ProgressBarUninst.Refresh()
                        End If
                    End If
                ElseIf UninstallWizard.InstUX = "UXTSB" Then
                    If ActionLabel.InvokeRequired Then
                        ActionLabel.Invoke(Sub()
                                               ActionLabel.Text = "Removing UXThemeSignatureBypass..."
                                               ActionLabel.Refresh()
                                           End Sub)
                    Else
                        ActionLabel.Text = "Removing UXThemeSignatureBypass..."
                        ActionLabel.Refresh()
                    End If

                    Try
                        ' Removing UXTSB after reboot
                        Using RegKeyDelete As RegistryKey = Registry.LocalMachine.OpenSubKey("SYSTEM\CurrentControlSet\Control\Session Manager", True)
                            If RegKeyDelete.GetValue("PendingFileRenameOperations") IsNot Nothing Or RegKeyDelete.GetValue("PendingFileRenameOperations") <> String.Empty Then
                                Dim RegKeyValue As String() = RegKeyDelete.GetValue("PendingFileRenameOperations")
                                Dim RegKeyValueString As String = String.Join(Chr(0), RegKeyValue)
                                Dim DeleteFiles As String() = {RegKeyValueString, "\??\" & WinDir & "\System32\UxThemeSignatureBypass.dll" & Chr(0), "\??\" & WinDir & "\SysWOW64\UxThemeSignatureBypass.dll" & Chr(0)}
                                Dim DeleteFiles32 As String() = {RegKeyValueString, "\??\" & WinDir & "\System32\UxThemeSignatureBypass.dll" & Chr(0)}
                                If UninstallWizard.BitnessSystem = "64Bit" Then
                                    RegKeyDelete.SetValue("PendingFileRenameOperations", DeleteFiles, RegistryValueKind.MultiString)
                                ElseIf UninstallWizard.BitnessSystem = "32Bit" Then
                                    RegKeyDelete.SetValue("PendingFileRenameOperations", DeleteFiles32, RegistryValueKind.MultiString)
                                End If
                            Else
                                Dim DeleteFiles As String() = {"\??\" & WinDir & "\System32\UxThemeSignatureBypass.dll" & Chr(0), "\??\" & WinDir & "\SysWOW64\UxThemeSignatureBypass.dll" & Chr(0)}
                                Dim DeleteFiles32 As String() = {"\??\" & WinDir & "\System32\UxThemeSignatureBypass.dll" & Chr(0)}
                                If UninstallWizard.BitnessSystem = "64Bit" Then
                                    RegKeyDelete.SetValue("PendingFileRenameOperations", DeleteFiles, RegistryValueKind.MultiString)
                                ElseIf UninstallWizard.BitnessSystem = "32Bit" Then
                                    RegKeyDelete.SetValue("PendingFileRenameOperations", DeleteFiles32, RegistryValueKind.MultiString)
                                End If
                            End If
                            RegKeyDelete.Close()
                        End Using
                        ' UXTSB -> remove registry key to load file with Windows
                        Using RegKey As RegistryKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\Microsoft\Windows NT\CurrentVersion\Windows", True)
                            Dim AppInitLoad As String = RegKey.GetValue("AppInit_DLLs")
                            Dim AppInitKeyList As List(Of String) = AppInitLoad.Split(",").ToList
                            If AppInitLoad <> WinDir & "\System32\UxThemeSignatureBypass.dll" Then
                                If AppInitKeyList.Contains(" " & WinDir & "\System32\UxThemeSignatureBypass.dll") Then
                                    AppInitKeyList.RemoveAt(AppInitKeyList.IndexOf(" " & WinDir & "\System32\UxThemeSignatureBypass.dll"))
                                    RegKey.SetValue("AppInit_DLLs", String.Join(",", AppInitKeyList))
                                End If
                                RegKey.Close()
                            Else
                                RegKey.SetValue("AppInit_DLLs", String.Empty)
                                If RegKey.GetValue("LoadAppInit_DLLs") = 1 Then
                                    RegKey.SetValue("LoadAppInit_DLLs", 0)
                                End If
                                RegKey.Close()
                            End If
                        End Using
                        ' UXTSB\Registry\SysWOW64
                        If UninstallWizard.BitnessSystem = "64Bit" Then
                            Using RegKey32 As RegistryKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\Wow6432Node\Microsoft\Windows NT\CurrentVersion\Windows", True)
                                Dim AppInitLoad32 As String = RegKey32.GetValue("AppInit_DLLs")
                                Dim AppInitKeyList32 As List(Of String) = AppInitLoad32.Split(",").ToList
                                If AppInitLoad32 <> WinDir & "\SysWOW64\UxThemeSignatureBypass.dll" Then
                                    If AppInitKeyList32.Contains(" " & WinDir & "\SysWOW64\UxThemeSignatureBypass.dll") Then
                                        AppInitKeyList32.RemoveAt(AppInitKeyList32.IndexOf(" " & WinDir & "\SysWOW64\UxThemeSignatureBypass.dll"))
                                        RegKey32.SetValue("AppInit_DLLs", String.Join(",", AppInitKeyList32))
                                    End If
                                    RegKey32.Close()
                                Else
                                    RegKey32.SetValue("AppInit_DLLs", String.Empty)
                                    If RegKey32.GetValue("LoadAppInit_DLLs") = 1 Then
                                        RegKey32.SetValue("LoadAppInit_DLLs", 0)
                                    End If
                                    RegKey32.Close()
                                End If
                            End Using
                        End If

                        If ProgressBarUninst.InvokeRequired Then
                            ProgressBarUninst.Invoke(Sub()
                                                         ProgressBarUninst.PerformStep()
                                                         ProgressBarUninst.Refresh()
                                                     End Sub)
                        Else
                            ProgressBarUninst.PerformStep()
                            ProgressBarUninst.Refresh()
                        End If

                        If ActionLabel.InvokeRequired Then
                            ActionLabel.Invoke(Sub()
                                                   ActionLabel.Text = "UXThemeSignatureBypass has been removed!"
                                                   ActionLabel.Refresh()
                                               End Sub)
                        Else
                            ActionLabel.Text = "UXThemeSignatureBypass has been removed!"
                            ActionLabel.Refresh()
                        End If
                    Catch ex As Exception
                        MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        If ActionLabel.InvokeRequired Then
                            ActionLabel.Invoke(Sub()
                                                   ActionLabel.Text = "UXThemeSignatureBypass could not be removed!"
                                                   ActionLabel.Refresh()
                                               End Sub)
                        Else
                            ActionLabel.Text = "UXThemeSignatureBypass could not be removed!"
                            ActionLabel.Refresh()
                        End If

                        If ProgressBarUninst.InvokeRequired Then
                            ProgressBarUninst.Invoke(Sub()
                                                         ProgressBarUninst.PerformStep()
                                                         ProgressBarUninst.Refresh()
                                                     End Sub)
                        Else
                            ProgressBarUninst.PerformStep()
                            ProgressBarUninst.Refresh()
                        End If
                    End Try
                End If
            Case "UninstQuero"
                If FileIO.FileSystem.DirectoryExists(ProgramFiles & "\Quero Toolbar") Then
                    ' Restore address bar of Internet Explorer
                    Using HklmIENoNavBar As RegistryKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\Policies\Microsoft\Internet Explorer\Toolbars\Restrictions", True)
                        If HklmIENoNavBar IsNot Nothing Then
                            If HklmIENoNavBar.GetValue("NoNavBar") IsNot Nothing Then
                                HklmIENoNavBar.DeleteValue("NoNavBar")
                            End If
                            HklmIENoNavBar.Close()
                        End If
                    End Using
                    Using HkcuIENoNavBar As RegistryKey = Registry.CurrentUser.OpenSubKey("Software\Policies\Microsoft\Internet Explorer\Toolbars\Restrictions", True)
                        If HkcuIENoNavBar IsNot Nothing Then
                            If HkcuIENoNavBar.GetValue("NoNavBar") IsNot Nothing Then
                                HkcuIENoNavBar.DeleteValue("NoNavBar")
                            End If
                            HkcuIENoNavBar.Close()
                        End If
                    End Using
                    ' Run uninstaller
                    Using QueroUninst As New Process
                        With QueroUninst
                            .StartInfo.FileName = ProgramFiles & "\Quero Toolbar\unins000.exe"
                            .StartInfo.Arguments = "/VERYSILENT /SUPPRESSMSGBOXES"
                            .Start()
                            .WaitForExit()
                        End With
                    End Using
                Else
                    MessageBox.Show("Quero Toolbar could not be found on this system!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    If ActionLabel.InvokeRequired Then
                        ActionLabel.Invoke(Sub()
                                               ActionLabel.Text = "Quero Toolbar could not be removed!"
                                               ActionLabel.Refresh()
                                           End Sub)
                    Else
                        ActionLabel.Text = "Quero Toolbar could not be removed!"
                        ActionLabel.Refresh()
                    End If

                    If ProgressBarUninst.InvokeRequired Then
                        ProgressBarUninst.Invoke(Sub()
                                                     ProgressBarUninst.PerformStep()
                                                     ProgressBarUninst.Refresh()
                                                 End Sub)
                    Else
                        ProgressBarUninst.PerformStep()
                        ProgressBarUninst.Refresh()
                    End If
                End If
            Case "RestoreSysFiles"
                Dim EveryoneSidAsUserName As String = New SecurityIdentifier(WellKnownSidType.WorldSid, Nothing).Translate(GetType(NTAccount)).Value
                Dim EveryoneAuditRule As FileSystemAuditRule = New FileSystemAuditRule(EveryoneSidAsUserName, FileSystemRights.CreateFiles + FileSystemRights.CreateDirectories + FileSystemRights.WriteAttributes + FileSystemRights.WriteExtendedAttributes + FileSystemRights.Delete + FileSystemRights.ChangePermissions + FileSystemRights.TakeOwnership, AuditFlags.Success + AuditFlags.Failure)
                Dim UiRibbonAuditRules As FileSecurity = Nothing
                ' Restore system files
                ' Check if backup files exist
                If UninstallWizard.BitnessSystem = "64Bit" Then
                    If Not FileIO.FileSystem.FileExists(WinDir & "\System32\UIRibbon.dll.bak") Or Not FileIO.FileSystem.FileExists(WinDir & "\System32\UIRibbonRes.dll.bak") Or Not FileIO.FileSystem.FileExists(WinDir & "\System32\" & UninstallWizard.ProgLanguage & "\UIRibbon.dll.mui.bak") Or Not FileIO.FileSystem.FileExists(WinDir & "\SysWOW64\UIRibbon.dll.bak") Or Not FileIO.FileSystem.FileExists(WinDir & "\SysWOW64\UIRibbonRes.dll.bak") Or Not FileIO.FileSystem.FileExists(WinDir & "\SysWOW64\" & UninstallWizard.ProgLanguage & "\UIRibbon.dll.mui.bak") Then
                        MessageBox.Show("One or more backup files could not be found on this system! Skipping restore of system files...", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        If ProgressBarUninst.InvokeRequired Then
                            ProgressBarUninst.Invoke(Sub()
                                                         ProgressBarUninst.PerformStep()
                                                         ProgressBarUninst.Refresh()
                                                     End Sub)
                        Else
                            ProgressBarUninst.PerformStep()
                            ProgressBarUninst.Refresh()
                        End If
                        Exit Select
                    End If
                ElseIf UninstallWizard.BitnessSystem = "32Bit" Then
                    If Not FileIO.FileSystem.FileExists(WinDir & "\System32\UIRibbon.dll.bak") Or Not FileIO.FileSystem.FileExists(WinDir & "\System32\UIRibbonRes.dll.bak") Or Not FileIO.FileSystem.FileExists(WinDir & "\System32\" & UninstallWizard.ProgLanguage & "\UIRibbon.dll.mui.bak") Then
                        MessageBox.Show("One or more backup files could not be found on this system! Skipping restore of system files...", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        If ProgressBarUninst.InvokeRequired Then
                            ProgressBarUninst.Invoke(Sub()
                                                         ProgressBarUninst.PerformStep()
                                                         ProgressBarUninst.Refresh()
                                                     End Sub)
                        Else
                            ProgressBarUninst.PerformStep()
                            ProgressBarUninst.Refresh()
                        End If
                        Exit Select
                    End If
                End If

                ' Restore files now
                Dim SysFilesSetup As New Process
                With SysFilesSetup
                    .StartInfo.FileName = WinDir & "\System32\takeown.exe"
                    .StartInfo.Arguments = "/f " & WinDir & "\System32\UIRibbon.dll /a"
                    .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                    .Start()
                    .WaitForExit()
                End With
                With SysFilesSetup
                    .StartInfo.FileName = WinDir & "\System32\icacls.exe"
                    .StartInfo.Arguments = WinDir & "\System32\UIRibbon.dll /grant *S-1-5-32-545:F *S-1-5-32-544:F"
                    .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                    .Start()
                    .WaitForExit()
                End With
                With SysFilesSetup
                    .StartInfo.FileName = WinDir & "\System32\takeown.exe"
                    .StartInfo.Arguments = "/f " & WinDir & "\System32\UIRibbonRes.dll /a"
                    .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                    .Start()
                    .WaitForExit()
                End With
                With SysFilesSetup
                    .StartInfo.FileName = WinDir & "\System32\icacls.exe"
                    .StartInfo.Arguments = WinDir & "\System32\UIRibbonRes.dll /grant *S-1-5-32-545:F *S-1-5-32-544:F"
                    .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                    .Start()
                    .WaitForExit()
                End With
                With SysFilesSetup
                    .StartInfo.FileName = WinDir & "\System32\takeown.exe"
                    .StartInfo.Arguments = "/f " & WinDir & "\System32\UIRibbon.dll.bak /a"
                    .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                    .Start()
                    .WaitForExit()
                End With
                With SysFilesSetup
                    .StartInfo.FileName = WinDir & "\System32\icacls.exe"
                    .StartInfo.Arguments = WinDir & "\System32\UIRibbon.dll.bak /grant *S-1-5-32-545:F *S-1-5-32-544:F"
                    .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                    .Start()
                    .WaitForExit()
                End With
                With SysFilesSetup
                    .StartInfo.FileName = WinDir & "\System32\takeown.exe"
                    .StartInfo.Arguments = "/f " & WinDir & "\System32\UIRibbonRes.dll.bak /a"
                    .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                    .Start()
                    .WaitForExit()
                End With
                With SysFilesSetup
                    .StartInfo.FileName = WinDir & "\System32\icacls.exe"
                    .StartInfo.Arguments = WinDir & "\System32\UIRibbonRes.dll.bak /grant *S-1-5-32-545:F *S-1-5-32-544:F"
                    .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                    .Start()
                    .WaitForExit()
                End With
                FileIO.FileSystem.DeleteFile(WinDir & "\System32\UIRibbon.dll")
                FileIO.FileSystem.DeleteFile(WinDir & "\System32\UIRibbonRes.dll")
                FileIO.FileSystem.RenameFile(WinDir & "\System32\UIRibbon.dll.bak", "UIRibbon.dll")
                FileIO.FileSystem.RenameFile(WinDir & "\System32\UIRibbonRes.dll.bak", "UIRibbonRes.dll")
                With SysFilesSetup
                    .StartInfo.FileName = WinDir & "\System32\takeown.exe"
                    .StartInfo.Arguments = "/f " & WinDir & "\System32\" & UninstallWizard.ProgLanguage & "\UIRibbon.dll.mui /a"
                    .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                    .Start()
                    .WaitForExit()
                End With
                With SysFilesSetup
                    .StartInfo.FileName = WinDir & "\System32\icacls.exe"
                    .StartInfo.Arguments = WinDir & "\System32\" & UninstallWizard.ProgLanguage & "\UIRibbon.dll.mui /grant *S-1-5-32-545:F *S-1-5-32-544:F"
                    .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                    .Start()
                    .WaitForExit()
                End With
                With SysFilesSetup
                    .StartInfo.FileName = WinDir & "\System32\takeown.exe"
                    .StartInfo.Arguments = "/f " & WinDir & "\System32\" & UninstallWizard.ProgLanguage & "\UIRibbon.dll.mui.bak /a"
                    .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                    .Start()
                    .WaitForExit()
                End With
                With SysFilesSetup
                    .StartInfo.FileName = WinDir & "\System32\icacls.exe"
                    .StartInfo.Arguments = WinDir & "\System32\" & UninstallWizard.ProgLanguage & "\UIRibbon.dll.mui.bak /grant *S-1-5-32-545:F *S-1-5-32-544:F"
                    .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                    .Start()
                    .WaitForExit()
                End With
                FileIO.FileSystem.DeleteFile(WinDir & "\System32\" & UninstallWizard.ProgLanguage & "\UIRibbon.dll.mui")
                FileIO.FileSystem.RenameFile(WinDir & "\System32\" & UninstallWizard.ProgLanguage & "\UIRibbon.dll.mui.bak", "UIRibbon.dll.mui")
                ' Original files -> change to original FileSystemSecurity
                With SysFilesSetup
                    .StartInfo.FileName = WinDir & "\System32\icacls.exe"
                    .StartInfo.Arguments = WinDir & "\System32\UIRibbon.dll /setowner ""NT SERVICE\TrustedInstaller"""
                    .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                    .Start()
                    .WaitForExit()
                End With
                With SysFilesSetup
                    .StartInfo.FileName = WinDir & "\System32\icacls.exe"
                    .StartInfo.Arguments = WinDir & "\System32\UIRibbon.dll /reset"
                    .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                    .Start()
                    .WaitForExit()
                End With
                With SysFilesSetup
                    .StartInfo.FileName = WinDir & "\System32\icacls.exe"
                    .StartInfo.Arguments = WinDir & "\System32\UIRibbon.dll /grant *S-1-5-18:RX *S-1-5-32-544:RX *S-1-5-32-545:RX *S-1-15-2-1:RX ""NT SERVICE\TrustedInstaller"":F /inheritance:r"
                    .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                    .Start()
                    .WaitForExit()
                End With
                UiRibbonAuditRules = File.GetAccessControl(WinDir & "\System32\UIRibbon.dll", AccessControlSections.Audit)
                UiRibbonAuditRules.SetAuditRule(EveryoneAuditRule)
                File.SetAccessControl(WinDir & "\System32\UIRibbon.dll", UiRibbonAuditRules)
                UiRibbonAuditRules = Nothing
                With SysFilesSetup
                    .StartInfo.FileName = WinDir & "\System32\icacls.exe"
                    .StartInfo.Arguments = WinDir & "\System32\UIRibbonRes.dll /setowner ""NT SERVICE\TrustedInstaller"""
                    .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                    .Start()
                    .WaitForExit()
                End With
                With SysFilesSetup
                    .StartInfo.FileName = WinDir & "\System32\icacls.exe"
                    .StartInfo.Arguments = WinDir & "\System32\UIRibbonRes.dll /reset"
                    .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                    .Start()
                    .WaitForExit()
                End With
                With SysFilesSetup
                    .StartInfo.FileName = WinDir & "\System32\icacls.exe"
                    .StartInfo.Arguments = WinDir & "\System32\UIRibbonRes.dll /grant *S-1-5-18:RX *S-1-5-32-544:RX *S-1-5-32-545:RX *S-1-15-2-1:RX ""NT SERVICE\TrustedInstaller"":F /inheritance:r"
                    .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                    .Start()
                    .WaitForExit()
                End With
                UiRibbonAuditRules = File.GetAccessControl(WinDir & "\System32\UIRibbonRes.dll", AccessControlSections.Audit)
                UiRibbonAuditRules.SetAuditRule(EveryoneAuditRule)
                File.SetAccessControl(WinDir & "\System32\UIRibbonRes.dll", UiRibbonAuditRules)
                UiRibbonAuditRules = Nothing
                With SysFilesSetup
                    .StartInfo.FileName = WinDir & "\System32\icacls.exe"
                    .StartInfo.Arguments = WinDir & "\System32\" & UninstallWizard.ProgLanguage & "\UIRibbon.dll.mui /setowner ""NT SERVICE\TrustedInstaller"""
                    .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                    .Start()
                    .WaitForExit()
                End With
                With SysFilesSetup
                    .StartInfo.FileName = WinDir & "\System32\icacls.exe"
                    .StartInfo.Arguments = WinDir & "\System32\" & UninstallWizard.ProgLanguage & "\UIRibbon.dll.mui /reset"
                    .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                    .Start()
                    .WaitForExit()
                End With
                With SysFilesSetup
                    .StartInfo.FileName = WinDir & "\System32\icacls.exe"
                    .StartInfo.Arguments = WinDir & "\System32\" & UninstallWizard.ProgLanguage & "\UIRibbon.dll.mui /grant *S-1-5-18:RX *S-1-5-32-544:RX *S-1-5-32-545:RX *S-1-15-2-1:RX ""NT SERVICE\TrustedInstaller"":F /inheritance:r"
                    .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                    .Start()
                    .WaitForExit()
                End With
                UiRibbonAuditRules = File.GetAccessControl(WinDir & "\System32\" & UninstallWizard.ProgLanguage & "\UIRibbon.dll.mui", AccessControlSections.Audit)
                UiRibbonAuditRules.SetAuditRule(EveryoneAuditRule)
                File.SetAccessControl(WinDir & "\System32\" & UninstallWizard.ProgLanguage & "\UIRibbon.dll.mui", UiRibbonAuditRules)
                UiRibbonAuditRules = Nothing
                ' If System is 64-bit, repeat for files in SysWOW64
                If UninstallWizard.BitnessSystem = "64Bit" Then
                    With SysFilesSetup
                        .StartInfo.FileName = WinDir & "\System32\takeown.exe"
                        .StartInfo.Arguments = "/f " & WinDir & "\SysWOW64\UIRibbon.dll /a"
                        .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                        .Start()
                        .WaitForExit()
                    End With
                    With SysFilesSetup
                        .StartInfo.FileName = WinDir & "\System32\icacls.exe"
                        .StartInfo.Arguments = WinDir & "\SysWOW64\UIRibbon.dll /grant *S-1-5-32-545:F *S-1-5-32-544:F"
                        .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                        .Start()
                        .WaitForExit()
                    End With
                    With SysFilesSetup
                        .StartInfo.FileName = WinDir & "\System32\takeown.exe"
                        .StartInfo.Arguments = "/f " & WinDir & "\SysWOW64\UIRibbonRes.dll /a"
                        .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                        .Start()
                        .WaitForExit()
                    End With
                    With SysFilesSetup
                        .StartInfo.FileName = WinDir & "\System32\icacls.exe"
                        .StartInfo.Arguments = WinDir & "\SysWOW64\UIRibbonRes.dll /grant *S-1-5-32-545:F *S-1-5-32-544:F"
                        .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                        .Start()
                        .WaitForExit()
                    End With
                    With SysFilesSetup
                        .StartInfo.FileName = WinDir & "\System32\takeown.exe"
                        .StartInfo.Arguments = "/f " & WinDir & "\SysWOW64\UIRibbon.dll.bak /a"
                        .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                        .Start()
                        .WaitForExit()
                    End With
                    With SysFilesSetup
                        .StartInfo.FileName = WinDir & "\System32\icacls.exe"
                        .StartInfo.Arguments = WinDir & "\SysWOW64\UIRibbon.dll.bak /grant *S-1-5-32-545:F *S-1-5-32-544:F"
                        .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                        .Start()
                        .WaitForExit()
                    End With
                    With SysFilesSetup
                        .StartInfo.FileName = WinDir & "\System32\takeown.exe"
                        .StartInfo.Arguments = "/f " & WinDir & "\SysWOW64\UIRibbonRes.dll.bak /a"
                        .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                        .Start()
                        .WaitForExit()
                    End With
                    With SysFilesSetup
                        .StartInfo.FileName = WinDir & "\System32\icacls.exe"
                        .StartInfo.Arguments = WinDir & "\SysWOW64\UIRibbonRes.dll.bak /grant *S-1-5-32-545:F *S-1-5-32-544:F"
                        .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                        .Start()
                        .WaitForExit()
                    End With
                    FileIO.FileSystem.DeleteFile(WinDir & "\SysWOW64\UIRibbon.dll")
                    FileIO.FileSystem.DeleteFile(WinDir & "\SysWOW64\UIRibbonRes.dll")
                    FileIO.FileSystem.RenameFile(WinDir & "\SysWOW64\UIRibbon.dll.bak", "UIRibbon.dll")
                    FileIO.FileSystem.RenameFile(WinDir & "\SysWOW64\UIRibbonRes.dll.bak", "UIRibbonRes.dll")
                    With SysFilesSetup
                        .StartInfo.FileName = WinDir & "\System32\takeown.exe"
                        .StartInfo.Arguments = "/f " & WinDir & "\SysWOW64\" & UninstallWizard.ProgLanguage & "\UIRibbon.dll.mui /a"
                        .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                        .Start()
                        .WaitForExit()
                    End With
                    With SysFilesSetup
                        .StartInfo.FileName = WinDir & "\System32\icacls.exe"
                        .StartInfo.Arguments = WinDir & "\SysWOW64\" & UninstallWizard.ProgLanguage & "\UIRibbon.dll.mui /grant *S-1-5-32-545:F *S-1-5-32-544:F"
                        .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                        .Start()
                        .WaitForExit()
                    End With
                    With SysFilesSetup
                        .StartInfo.FileName = WinDir & "\System32\takeown.exe"
                        .StartInfo.Arguments = "/f " & WinDir & "\SysWOW64\" & UninstallWizard.ProgLanguage & "\UIRibbon.dll.mui.bak /a"
                        .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                        .Start()
                        .WaitForExit()
                    End With
                    With SysFilesSetup
                        .StartInfo.FileName = WinDir & "\System32\icacls.exe"
                        .StartInfo.Arguments = WinDir & "\SysWOW64\" & UninstallWizard.ProgLanguage & "\UIRibbon.dll.mui.bak /grant *S-1-5-32-545:F *S-1-5-32-544:F"
                        .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                        .Start()
                        .WaitForExit()
                    End With
                    FileIO.FileSystem.DeleteFile(WinDir & "\SysWOW64\" & UninstallWizard.ProgLanguage & "\UIRibbon.dll.mui")
                    FileIO.FileSystem.RenameFile(WinDir & "\SysWOW64\" & UninstallWizard.ProgLanguage & "\UIRibbon.dll.mui.bak", "UIRibbon.dll.mui")
                    ' Original files in SysWOW64 -> change to original FileSystemSecurity
                    With SysFilesSetup
                        .StartInfo.FileName = WinDir & "\System32\icacls.exe"
                        .StartInfo.Arguments = WinDir & "\SysWOW64\UIRibbon.dll /setowner ""NT SERVICE\TrustedInstaller"""
                        .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                        .Start()
                        .WaitForExit()
                    End With
                    With SysFilesSetup
                        .StartInfo.FileName = WinDir & "\System32\icacls.exe"
                        .StartInfo.Arguments = WinDir & "\SysWOW64\UIRibbon.dll /reset"
                        .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                        .Start()
                        .WaitForExit()
                    End With
                    With SysFilesSetup
                        .StartInfo.FileName = WinDir & "\System32\icacls.exe"
                        .StartInfo.Arguments = WinDir & "\SysWOW64\UIRibbon.dll /grant *S-1-5-18:RX *S-1-5-32-544:RX *S-1-5-32-545:RX *S-1-15-2-1:RX ""NT SERVICE\TrustedInstaller"":F /inheritance:r"
                        .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                        .Start()
                        .WaitForExit()
                    End With
                    UiRibbonAuditRules = File.GetAccessControl(WinDir & "\SysWOW64\UIRibbon.dll", AccessControlSections.Audit)
                    UiRibbonAuditRules.SetAuditRule(EveryoneAuditRule)
                    File.SetAccessControl(WinDir & "\SysWOW64\UIRibbon.dll", UiRibbonAuditRules)
                    UiRibbonAuditRules = Nothing
                    With SysFilesSetup
                        .StartInfo.FileName = WinDir & "\System32\icacls.exe"
                        .StartInfo.Arguments = WinDir & "\SysWOW64\UIRibbonRes.dll /setowner ""NT SERVICE\TrustedInstaller"""
                        .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                        .Start()
                        .WaitForExit()
                    End With
                    With SysFilesSetup
                        .StartInfo.FileName = WinDir & "\System32\icacls.exe"
                        .StartInfo.Arguments = WinDir & "\SysWOW64\UIRibbonRes.dll /reset"
                        .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                        .Start()
                        .WaitForExit()
                    End With
                    With SysFilesSetup
                        .StartInfo.FileName = WinDir & "\System32\icacls.exe"
                        .StartInfo.Arguments = WinDir & "\SysWOW64\UIRibbonRes.dll /grant *S-1-5-18:RX *S-1-5-32-544:RX *S-1-5-32-545:RX *S-1-15-2-1:RX ""NT SERVICE\TrustedInstaller"":F /inheritance:r"
                        .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                        .Start()
                        .WaitForExit()
                    End With
                    UiRibbonAuditRules = File.GetAccessControl(WinDir & "\SysWOW64\UIRibbonRes.dll", AccessControlSections.Audit)
                    UiRibbonAuditRules.SetAuditRule(EveryoneAuditRule)
                    File.SetAccessControl(WinDir & "\SysWOW64\UIRibbonRes.dll", UiRibbonAuditRules)
                    UiRibbonAuditRules = Nothing
                    With SysFilesSetup
                        .StartInfo.FileName = WinDir & "\System32\icacls.exe"
                        .StartInfo.Arguments = WinDir & "\SysWOW64\" & UninstallWizard.ProgLanguage & "\UIRibbon.dll.mui /setowner ""NT SERVICE\TrustedInstaller"""
                        .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                        .Start()
                        .WaitForExit()
                    End With
                    With SysFilesSetup
                        .StartInfo.FileName = WinDir & "\System32\icacls.exe"
                        .StartInfo.Arguments = WinDir & "\SysWOW64\" & UninstallWizard.ProgLanguage & "\UIRibbon.dll.mui /reset"
                        .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                        .Start()
                        .WaitForExit()
                    End With
                    With SysFilesSetup
                        .StartInfo.FileName = WinDir & "\System32\icacls.exe"
                        .StartInfo.Arguments = WinDir & "\SysWOW64\" & UninstallWizard.ProgLanguage & "\UIRibbon.dll.mui /grant *S-1-5-18:RX *S-1-5-32-544:RX *S-1-5-32-545:RX *S-1-15-2-1:RX ""NT SERVICE\TrustedInstaller"":F /inheritance:r"
                        .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                        .Start()
                        .WaitForExit()
                    End With
                    UiRibbonAuditRules = File.GetAccessControl(WinDir & "\SysWOW64\" & UninstallWizard.ProgLanguage & "\UIRibbon.dll.mui", AccessControlSections.Audit)
                    UiRibbonAuditRules.SetAuditRule(EveryoneAuditRule)
                    File.SetAccessControl(WinDir & "\SysWOW64\" & UninstallWizard.ProgLanguage & "\UIRibbon.dll.mui", UiRibbonAuditRules)
                    UiRibbonAuditRules = Nothing
                End If
                SysFilesSetup.Dispose()
                If ProgressBarUninst.InvokeRequired Then
                    ProgressBarUninst.Invoke(Sub()
                                                 ProgressBarUninst.PerformStep()
                                                 ProgressBarUninst.Refresh()
                                             End Sub)
                Else
                    ProgressBarUninst.PerformStep()
                    ProgressBarUninst.Refresh()
                End If
                If ActionLabel.InvokeRequired Then
                    ActionLabel.Invoke(Sub()
                                           ActionLabel.Text = "UIRibbon.dll && UIRibbonRes.dll have been restored!"
                                           ActionLabel.Refresh()
                                       End Sub)
                Else
                    ActionLabel.Text = "UIRibbon.dll && UIRibbonRes.dll have been restored!"
                    ActionLabel.Refresh()
                End If
            Case "UninstSounds"
                If FileIO.FileSystem.DirectoryExists(WinDir & "\Media\Windows 8 Release Preview") Then
                    FileIO.FileSystem.DeleteDirectory(WinDir & "\Media\Windows 8 Release Preview", FileIO.DeleteDirectoryOption.DeleteAllContents)
                    Using SoundSchemeNamesRegKey As RegistryKey = Registry.CurrentUser.OpenSubKey("AppEvents\Schemes\Names", True)
                        Using SoundSchemeEventsRegKey As RegistryKey = Registry.CurrentUser.OpenSubKey("AppEvents\Schemes\Apps", True)
                            ' Remove sound scheme from registry if available
                            Dim ReleasePreviewSoundSchemeKeyName As String = String.Empty
                            For Each SoundScheme In SoundSchemeNamesRegKey.GetSubKeyNames
                                If SoundSchemeNamesRegKey.OpenSubKey(SoundScheme).GetValue(String.Empty) = "Windows 8 Release Preview" Then
                                    ReleasePreviewSoundSchemeKeyName = SoundScheme
                                End If
                            Next
                            ' Remove sound scheme from events if available
                            If ReleasePreviewSoundSchemeKeyName <> String.Empty Then
                                For Each SoundSchemeEventCategory In SoundSchemeEventsRegKey.GetSubKeyNames
                                    For Each SoundSchemeEvent In SoundSchemeEventsRegKey.OpenSubKey(SoundSchemeEventCategory).GetSubKeyNames
                                        For Each SoundSchemeEventScheme In SoundSchemeEventsRegKey.OpenSubKey(SoundSchemeEventCategory & "\" & SoundSchemeEvent).GetSubKeyNames
                                            If SoundSchemeEventScheme = ReleasePreviewSoundSchemeKeyName Then
                                                SoundSchemeEventsRegKey.DeleteSubKey(SoundSchemeEventCategory & "\" & SoundSchemeEvent & "\" & SoundSchemeEventScheme)
                                            End If
                                        Next
                                    Next
                                Next
                                SoundSchemeNamesRegKey.DeleteSubKey(ReleasePreviewSoundSchemeKeyName)
                                SoundSchemeNamesRegKey.Close()
                                SoundSchemeEventsRegKey.Close()
                            End If
                        End Using
                    End Using

                    If ProgressBarUninst.InvokeRequired Then
                        ProgressBarUninst.Invoke(Sub()
                                                     ProgressBarUninst.PerformStep()
                                                     ProgressBarUninst.Refresh()
                                                 End Sub)
                    Else
                        ProgressBarUninst.PerformStep()
                        ProgressBarUninst.Refresh()
                    End If

                    If ActionLabel.InvokeRequired Then
                        ActionLabel.Invoke(Sub()
                                               ActionLabel.Text = "Sounds have been removed!"
                                               ActionLabel.Refresh()
                                           End Sub)
                    Else
                        ActionLabel.Text = "Sounds have been removed!"
                        ActionLabel.Refresh()
                    End If
                Else
                    MessageBox.Show("Sound files could not be found on this system!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    If ActionLabel.InvokeRequired Then
                        ActionLabel.Invoke(Sub()
                                               ActionLabel.Text = "Sounds could not be removed!"
                                               ActionLabel.Refresh()
                                           End Sub)
                    Else
                        ActionLabel.Text = "Sounds could not be removed!"
                        ActionLabel.Refresh()
                    End If

                    If ProgressBarUninst.InvokeRequired Then
                        ProgressBarUninst.Invoke(Sub()
                                                     ProgressBarUninst.PerformStep()
                                                     ProgressBarUninst.Refresh()
                                                 End Sub)
                    Else
                        ProgressBarUninst.PerformStep()
                        ProgressBarUninst.Refresh()
                    End If
                End If
            Case "UninstGadgets"
                If FileIO.FileSystem.FileExists(ProgramFiles & "\Windows Sidebar\8GadgetPack.exe") Then
                    Using GadgetsUninstall As New Process
                        With GadgetsUninstall
                            .StartInfo.FileName = WinDir & "\System32\msiexec.exe"
                            .StartInfo.Arguments = "/X{B6AF19AD-2D5B-44DC-9272-EC91965123E8} /quiet"
                            .Start()
                            .WaitForExit()
                        End With
                    End Using

                    If ProgressBarUninst.InvokeRequired Then
                        ProgressBarUninst.Invoke(Sub()
                                                     ProgressBarUninst.PerformStep()
                                                     ProgressBarUninst.Refresh()
                                                 End Sub)
                    Else
                        ProgressBarUninst.PerformStep()
                        ProgressBarUninst.Refresh()
                    End If

                    If ActionLabel.InvokeRequired Then
                        ActionLabel.Invoke(Sub()
                                               ActionLabel.Text = "8GadgetPack has been uninstalled!"
                                               ActionLabel.Refresh()
                                           End Sub)
                    Else
                        ActionLabel.Text = "8GadgetPack has been uninstalled!"
                        ActionLabel.Refresh()
                    End If
                Else
                    MessageBox.Show("8GadgetPack could not be found on this system!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    If ActionLabel.InvokeRequired Then
                        ActionLabel.Invoke(Sub()
                                               ActionLabel.Text = "Gadgets could not be removed!"
                                               ActionLabel.Refresh()
                                           End Sub)
                    Else
                        ActionLabel.Text = "Gadgets could not be removed!"
                        ActionLabel.Refresh()
                    End If

                    If ProgressBarUninst.InvokeRequired Then
                        ProgressBarUninst.Invoke(Sub()
                                                     ProgressBarUninst.PerformStep()
                                                     ProgressBarUninst.Refresh()
                                                 End Sub)
                    Else
                        ProgressBarUninst.PerformStep()
                        ProgressBarUninst.Refresh()
                    End If
                End If
            Case "Uninst7TaskTw"
                If FileIO.FileSystem.DirectoryExists(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) & "\Programs\7+ Taskbar Tweaker") Then
                    Using TaskTwSetup As New Process
                        With TaskTwSetup
                            .StartInfo.FileName = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) & "\Programs\7+ Taskbar Tweaker\uninstall.exe"
                            .StartInfo.Arguments = "/S"
                            .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                            .Start()
                            .WaitForExit()
                        End With
                    End Using
                    Do
                        Application.DoEvents()
                    Loop Until Process.GetProcessesByName("Un").Count = 0

                    If ActionLabel.InvokeRequired Then
                        ActionLabel.Invoke(Sub()
                                               ActionLabel.Text = "7+ Taskbar Tweaker has been uninstalled!"
                                               ActionLabel.Refresh()
                                           End Sub)
                    Else
                        ActionLabel.Text = "7+ Taskbar Tweaker has been uninstalled!"
                        ActionLabel.Refresh()
                    End If

                    If ProgressBarUninst.InvokeRequired Then
                        ProgressBarUninst.Invoke(Sub()
                                                     ProgressBarUninst.PerformStep()
                                                     ProgressBarUninst.Refresh()
                                                 End Sub)
                    Else
                        ProgressBarUninst.PerformStep()
                        ProgressBarUninst.Refresh()
                    End If
                Else
                    MessageBox.Show("7+ Taskbar Tweaker could not be found on this system!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    If ActionLabel.InvokeRequired Then
                        ActionLabel.Invoke(Sub()
                                               ActionLabel.Text = "7+ Taskbar Tweaker could not be uninstalled!"
                                               ActionLabel.Refresh()
                                           End Sub)
                    Else
                        ActionLabel.Text = "7+ Taskbar Tweaker could not be uninstalled!"
                        ActionLabel.Refresh()
                    End If

                    If ProgressBarUninst.InvokeRequired Then
                        ProgressBarUninst.Invoke(Sub()
                                                     ProgressBarUninst.PerformStep()
                                                     ProgressBarUninst.Refresh()
                                                 End Sub)
                    Else
                        ProgressBarUninst.PerformStep()
                        ProgressBarUninst.Refresh()
                    End If
                End If
        End Select

        ' CancelAsync of BackgroundWorker after task finishes
        BgWorker.CancelAsync()
    End Sub

    Private Sub UninstallLauncher()
        ' Initialise progress bar
        With ProgressBarUninst
            .Style = ProgressBarStyle.Continuous
            .Value = 0
            .Maximum = ProgBarCounter
            .Refresh()
        End With

        ' Close explorer.exe
        BgWorker.RunWorkerAsync("CloseExplorer")
        WaitUntilTaskFinishes()

        ' Run uninstaller now
        RunUninstall()
    End Sub

    Private Sub RunUninstall()
        ' Uninstall Aero Glass
        If UninstallWizard.UninstAero = True Then
            ActionLabel.Text = "Removing Aero Glass..."
            ActionLabel.Refresh()
            BgWorker.RunWorkerAsync("UninstAeroGlass")
            WaitUntilTaskFinishes()
        End If

        ' Uninstall OldNewExplorer
        If UninstallWizard.UninstONE = True Then
            ActionLabel.Text = "Removing OldNewExplorer..."
            ActionLabel.Refresh()
            BgWorker.RunWorkerAsync("UninstONE")
            WaitUntilTaskFinishes()
        End If

        ' Remove UXTheme patcher
        If UninstallWizard.UninstUX = True Then
            BgWorker.RunWorkerAsync("UninstUX")
            WaitUntilTaskFinishes()
        End If

        ' Remove Quero Toolbar
        If UninstallWizard.UninstQuero = True Then
            ActionLabel.Text = "Removing Quero Toolbar..."
            ActionLabel.Refresh()
            BgWorker.RunWorkerAsync("UninstQuero")
            WaitUntilTaskFinishes()
        End If

        ' Restore UIRibbon.dll & UIRibbonRes.dll
        If UninstallWizard.UninstSysFiles = True Then
            ActionLabel.Text = "Restoring UIRibbon.dll && UIRibbonRes.dll..."
            ActionLabel.Refresh()
            BgWorker.RunWorkerAsync("RestoreSysFiles")
            WaitUntilTaskFinishes()
        End If

        ' Remove sounds and sound scheme
        If UninstallWizard.UninstSounds = True Then
            ActionLabel.Text = "Removing sound scheme..."
            ActionLabel.Refresh()
            BgWorker.RunWorkerAsync("UninstSounds")
            WaitUntilTaskFinishes()
        End If

        ' Uninstall 8GadgetPack
        If UninstallWizard.UninstGadgets = True Then
            ActionLabel.Text = "Removing 8GadgetPack..."
            ActionLabel.Refresh()
            BgWorker.RunWorkerAsync("UninstGadgets")
            WaitUntilTaskFinishes()
        End If

        ' Uninstall 7+ Taskbar Tweaker
        If UninstallWizard.Uninst7TaskTw = True Then
            ActionLabel.Text = "Removing 7+ Taskbar Tweaker..."
            ActionLabel.Refresh()
            BgWorker.RunWorkerAsync("Uninst7TaskTw")
            WaitUntilTaskFinishes()
        End If

        ' Remove theme and wallpapers
        ActionLabel.Text = "Removing theme..."
        ActionLabel.Refresh()
        Try
            FileIO.FileSystem.DeleteFile(WinDir & "\Resources\Themes\aerorp.theme")
            If UninstallWizard.InstSounds = True Then
                FileIO.FileSystem.DeleteFile(WinDir & "\Resources\Themes\aerorp_sounds.theme")
            End If
            FileIO.FileSystem.DeleteFile(WinDir & "\Resources\Themes\theme1rp.theme")
            FileIO.FileSystem.DeleteDirectory(WinDir & "\Resources\Themes\aerorp", FileIO.DeleteDirectoryOption.DeleteAllContents)
            FileIO.FileSystem.DeleteDirectory(WinDir & "\Web\Wallpaper\Windows 8 Release Preview", FileIO.DeleteDirectoryOption.DeleteAllContents)
            ProgressBarUninst.PerformStep()
            ProgressBarUninst.Refresh()
            ActionLabel.Text = "Theme has been removed!"
            ActionLabel.Refresh()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ActionLabel.Text = "Theme could not be removed!"
            ActionLabel.Refresh()
            ProgressBarUninst.PerformStep()
            ProgressBarUninst.Refresh()
        End Try

        ' Removing program completely after reboot
        ActionLabel.Text = "Removing program..."
        ActionLabel.Refresh()
        ' Delete OldNewExplorer directory after reboot
        If UninstallWizard.UninstONE = True Then
            Using DeleteONERegKey As RegistryKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce", True)
                DeleteONERegKey.SetValue("Remove OldNewExplorer", """" & WinDir & "\System32\cmd.exe"" /c ""rmdir /s /q """ & ProgramFiles & "\OldNewExplorer""""")
                DeleteONERegKey.Close()
            End Using
        End If
        ' Delete this program’s folder after reboot
        Using DeleteProgramRegKey As RegistryKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce", True)
            DeleteProgramRegKey.SetValue("Remove Windows 8.1 to 8 RP Converter", """" & WinDir & "\System32\cmd.exe"" /c ""rmdir /s /q """ & UninstallWizard.ProgDir & """""")
            DeleteProgramRegKey.Close()
        End Using
        ' Remove entry from uninstall list in registry
        Using RegKeyControlPanel As RegistryKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall", True)
            If RegKeyControlPanel.OpenSubKey("Windows 8.1 to 8 RP Converter") IsNot Nothing Then
                RegKeyControlPanel.DeleteSubKey("Windows 8.1 to 8 RP Converter")
            End If
            RegKeyControlPanel.Close()
        End Using
        ProgressBarUninst.PerformStep()
        ProgressBarUninst.Refresh()
        ActionLabel.Text = "Theme and program(s) have been successfully removed!"
        ActionLabel.Refresh()
        ' Change Next button and mark setup as finished
        NextButton.Text = "&Finish"
        NextButton.Enabled = True
        IsFinished = True
    End Sub

    Private Sub NextButton_Click(sender As Object, e As EventArgs) Handles NextButton.Click
        If IsFinished = True Then
            Application.Exit()
        Else
            MessageBox.Show("The setup cannot be finished right now.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End If
    End Sub

    Private Sub UninstallUninstallEnglish_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        If IsFinished = True Then
            ' Restart prompt
            Dim ShutDownProcess As New Process
            If MessageBox.Show("You need to restart your system to finish the uninstallation. Do you want to restart now?", "Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) = System.Windows.Forms.DialogResult.Yes Then
                With ShutDownProcess
                    .StartInfo.FileName = WinDir & "\System32\shutdown.exe"
                    .StartInfo.Arguments = "/r /t 0"
                    .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                    .Start()
                End With
                Application.Exit()
            ElseIf System.Windows.Forms.DialogResult.No Then
                Application.Exit()
            End If
        Else
            MessageBox.Show("The setup cannot be finished right now.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End If
    End Sub
End Class