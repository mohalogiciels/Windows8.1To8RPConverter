Imports System.ComponentModel
Imports Microsoft.Win32

Public Class UninstallUninstallEnglish
    Dim ProgramFiles = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles)
    Dim ProgramFilesX86 = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86)
    Dim WinDir = Environment.GetFolderPath(Environment.SpecialFolder.Windows)
    Dim WithEvents BgWorker As New BackgroundWorker
    Dim ProgBarCounter As Integer = 20
    Dim IsFinished As Boolean = False
    Dim RestartNow As Boolean = False

    Private Sub Wait(ByVal seconds As Integer)
        For Index As Integer = 0 To seconds * 100
            System.Threading.Thread.Sleep(10)
            Application.DoEvents()
        Next
    End Sub

    Private Sub UninstallUninstallEnglish_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown
        ' Set progress bar to marquee
        With ProgressBarUninst
            .Style = ProgressBarStyle.Marquee
            .MarqueeAnimationSpeed = 100
            .Refresh()
        End With

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

        ' Change theme back to standard
        Me.TopMost = True
        BgWorker.RunWorkerAsync()
    End Sub

    Private Sub BgWorker_DoWork(sender As Object, e As DoWorkEventArgs) Handles BgWorker.DoWork
        Dim ThemeChangeProcess As New Process
        With ThemeChangeProcess
            .StartInfo.FileName = WinDir & "\Resources\Themes\aero.theme"
            .Start()
        End With
        Wait(5)
        BgWorker.CancelAsync()
    End Sub

    Private Sub BgWorker_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles BgWorker.RunWorkerCompleted
        ' Initialise progress bar
        With ProgressBarUninst
            .Style = ProgressBarStyle.Continuous
            .Value = 0
            .Maximum = ProgBarCounter
            .Refresh()
        End With

        ' Run uninstaller now
        Me.TopMost = False
        RunUninstall()
    End Sub

    Private Sub RunUninstall()
        Dim ProgDir As String = ProgramFiles & "\Windows8.1to8RPConv"
        Dim ResFolder As String = ProgDir & "\res"

        ' Uninstall Aero Glass
        If UninstallWizard.UninstAero = True Then
            '' Get Aero Glass path from uninstall list in registry
            ActionLabel.Text = "Removing Aero Glass..."
            ActionLabel.Refresh()
            Dim UninstallReg As RegistryKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall", False)
            Dim UninstallRegKeys As String() = UninstallReg.GetSubKeyNames
            Dim UninstallAero As New List(Of String)
            Dim AeroGlassPath As String = String.Empty
            For Each RegKey In UninstallRegKeys
                If RegKey.StartsWith("Aero Glass for Win8.1") Then
                    UninstallAero.Add(RegKey)
                End If
            Next
            '' Check if Aero Glass folder exists
            If UninstallAero.Count > 0 Then
                Dim UninstallAeroKey As RegistryKey = UninstallReg.OpenSubKey(UninstallAero(0))
                AeroGlassPath = UninstallAeroKey.GetValue("InstallLocation")
                UninstallAeroKey.Close()
            End If
            UninstallReg.Close()
            If FileIO.FileSystem.DirectoryExists(AeroGlassPath) Then
                Dim AeroGlassUninst As New Process
                With AeroGlassUninst
                    .StartInfo.FileName = AeroGlassPath & "unins000.exe"
                    .StartInfo.Arguments = "/SILENT /NORESTART /SUPPRESSMSGBOXES"
                    .Start()
                    .WaitForExit()
                End With
                ProgressBarUninst.PerformStep()
                ProgressBarUninst.Refresh()
                ActionLabel.Text = "Aero Glass has been removed!"
                ActionLabel.Refresh()
            Else
                MessageBox.Show("Aero Glass could not be found on this system!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                ActionLabel.Text = "Aero Glass could not be uninstalled!"
                ActionLabel.Refresh()
                ProgressBarUninst.PerformStep()
                ProgressBarUninst.Refresh()
            End If
        End If

        ' Uninstall OldNewExplorer
        If UninstallWizard.UninstONE = True Then
            ActionLabel.Text = "Removing OldNewExplorer..."
            ActionLabel.Refresh()
            If FileIO.FileSystem.DirectoryExists(ProgramFiles & "\OldNewExplorer") Then
                Dim ONESetup As New Process
                '' Unreg OldNewExplorer DLLs
                If UninstallWizard.BitnessSystem = "64bit" Then
                    With ONESetup
                        .StartInfo.FileName = WinDir & "\System32\regsvr32.exe"
                        .StartInfo.Arguments = "/u /s """ & ProgramFiles & "\OldNewExplorer\OldNewExplorer64.dll"""
                        .Start()
                        .WaitForExit()
                    End With
                    With ONESetup
                        .StartInfo.FileName = WinDir & "\System32\regsvr32.exe"
                        .StartInfo.Arguments = "/u /s """ & ProgramFiles & "\OldNewExplorer\OldNewExplorer32.dll"""
                        .Start()
                        .WaitForExit()
                    End With
                ElseIf UninstallWizard.BitnessSystem = "32bit" Then
                    With ONESetup
                        .StartInfo.FileName = WinDir & "\System32\regsvr32.exe"
                        .StartInfo.Arguments = "/u /s """ & ProgramFiles & "\OldNewExplorer\OldNewExplorer32.dll"""
                        .Start()
                        .WaitForExit()
                    End With
                End If
                '' Remove desktop shortcut
                If FileIO.FileSystem.FileExists(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) & "\OldNewExplorer Configuration.lnk") Then
                    FileIO.FileSystem.DeleteFile(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) & "\OldNewExplorer Configuration.lnk")
                End If
                ProgressBarUninst.PerformStep()
                ProgressBarUninst.Refresh()
                ActionLabel.Text = "OldNewExplorer has been removed!"
                ActionLabel.Refresh()
            Else
                MessageBox.Show("OldNewExplorer could not be found on this system!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                ActionLabel.Text = "OldNewExplorer could not be removed!"
                ActionLabel.Refresh()
                ProgressBarUninst.PerformStep()
                ProgressBarUninst.Refresh()
            End If
        End If

        ' Remove UXTheme patcher
        If UninstallWizard.UninstUX = True Then
            If UninstallWizard.InstUX = "UltraUX" Then
                '' Uninstall UltraUX
                ActionLabel.Text = "Uninstalling UltraUXThemePatcher..."
                ActionLabel.Refresh()
                If FileIO.FileSystem.FileExists(ProgramFilesX86 & "\UltraUXThemePatcher\Uninstall.exe") Or FileIO.FileSystem.FileExists(ProgramFiles & "\UltraUXThemePatcher\Uninstall.exe") Then
                    If UninstallWizard.BitnessSystem = "64bit" Then
                        Dim UninstUltraUX As New Process
                        With UninstUltraUX
                            .StartInfo.FileName = ProgramFilesX86 & "\UltraUXThemePatcher\Uninstall.exe"
                            .Start()
                        End With
                        '' WaitForExit for process called Un.exe because Uninstall.exe closes
                        Do
                            System.Threading.Thread.Sleep(1000)
                        Loop Until Process.GetProcessesByName("Un").Count = 0
                        Application.DoEvents()
                    ElseIf UninstallWizard.BitnessSystem = "32bit" Then
                        Dim UninstUltraUX As New Process
                        With UninstUltraUX
                            .StartInfo.FileName = ProgramFiles & "\UltraUXThemePatcher\Uninstall.exe"
                            .Start()
                        End With
                        '' WaitForExit for process called Un.exe because Uninstall.exe closes
                        Do
                            System.Threading.Thread.Sleep(1000)
                        Loop Until Process.GetProcessesByName("Un").Count = 0
                        Application.DoEvents()
                    End If
                    ProgressBarUninst.PerformStep()
                    ProgressBarUninst.Refresh()
                    ActionLabel.Text = "UltraUXThemePatcher has been uninstalled!"
                    ActionLabel.Refresh()
                Else
                    MessageBox.Show("UltraUXThemePatcher could not be found on this system!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    ActionLabel.Text = "UltraUXThemePatcher could not be removed!"
                    ActionLabel.Refresh()
                    ProgressBarUninst.PerformStep()
                    ProgressBarUninst.Refresh()
                End If
            ElseIf UninstallWizard.InstUX = "UXTSB" Then
                ActionLabel.Text = "Removing UXThemeSignatureBypass..."
                ActionLabel.Refresh()
                '' Removing UXTSB after reboot
                Try
                    Dim SessionManagerRegKey As RegistryKey = Registry.LocalMachine.OpenSubKey("SYSTEM\CurrentControlSet\Control\Session Manager", True)
                    If UninstallWizard.BitnessSystem = "64bit" Then
                        If SessionManagerRegKey.GetValue("PendingFileRenameOperations") IsNot Nothing Then
                            Dim RegKeyValue As String() = SessionManagerRegKey.GetValue("PendingFileRenameOperations")
                            Dim RegKeyValueString As String = String.Join(Chr(0), RegKeyValue)
                            Dim DeleteFiles As String() = {RegKeyValueString, "\??\" & WinDir & "\UxThemeSignatureBypass64.dll" & Chr(0), "\??\" & WinDir & "\UxThemeSignatureBypass32.dll" & Chr(0)}
                            SessionManagerRegKey.SetValue("PendingFileRenameOperations", DeleteFiles, RegistryValueKind.MultiString)
                        Else
                            Dim DeleteFiles As String() = {"\??\" & WinDir & "\UxThemeSignatureBypass64.dll" & Chr(0), "\??\" & WinDir & "\UxThemeSignatureBypass32.dll" & Chr(0)}
                            SessionManagerRegKey.SetValue("PendingFileRenameOperations", DeleteFiles, RegistryValueKind.MultiString)
                        End If
                    ElseIf UninstallWizard.BitnessSystem = "32bit" Then
                        If SessionManagerRegKey.GetValue("PendingFileRenameOperations") IsNot Nothing Then
                            Dim RegKeyValue As String() = SessionManagerRegKey.GetValue("PendingFileRenameOperations")
                            Dim RegKeyValueString As String = String.Join(Chr(0), RegKeyValue)
                            Dim DeleteFiles As String() = {RegKeyValueString, "\??\" & WinDir & "\UxThemeSignatureBypass32.dll" & Chr(0)}
                            SessionManagerRegKey.SetValue("PendingFileRenameOperations", DeleteFiles, RegistryValueKind.MultiString)
                        Else
                            Dim DeleteFiles As String() = {"\??\" & WinDir & "\UxThemeSignatureBypass32.dll" & Chr(0)}
                            SessionManagerRegKey.SetValue("PendingFileRenameOperations", DeleteFiles, RegistryValueKind.MultiString)
                        End If
                    End If
                    SessionManagerRegKey.Close()
                    '' UXTSB remove registry values to load file(s) with Windows
                    Dim RegKey As RegistryKey
                    Dim RegKey32 As RegistryKey
                    If UninstallWizard.BitnessSystem = "64bit" Then
                        RegKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\Microsoft\Windows NT\CurrentVersion\Windows", True)
                        Dim AppInitLoad As String = RegKey.GetValue("AppInit_DLLs")
                        Dim AppInitKeyList As List(Of String) = AppInitLoad.Split(",").ToList
                        If AppInitLoad <> WinDir & "\UxThemeSignatureBypass64.dll" Then
                            If AppInitKeyList.Contains(" " & WinDir & "\UxThemeSignatureBypass64.dll") Then
                                AppInitKeyList.RemoveAt(AppInitKeyList.IndexOf(" " & WinDir & "\UxThemeSignatureBypass64.dll"))
                                Dim StringWithoutUXTSB As String = String.Join(",", AppInitKeyList)
                                RegKey.SetValue("AppInit_DLLs", StringWithoutUXTSB)
                            End If
                            RegKey.Close()
                        Else
                            RegKey.SetValue("AppInit_DLLs", "")
                            If RegKey.GetValue("LoadAppInit_DLLs") = 1 Then
                                RegKey.SetValue("LoadAppInit_DLLs", 0)
                            End If
                            RegKey.Close()
                        End If
                        '' UXTSB\Registry\SysWOW64
                        RegKey32 = Registry.LocalMachine.OpenSubKey("SOFTWARE\Wow6432Node\Microsoft\Windows NT\CurrentVersion\Windows", True)
                        Dim AppInitLoad32 As String = RegKey32.GetValue("AppInit_DLLs")
                        Dim AppInitKeyList32 As List(Of String) = AppInitLoad32.Split(",").ToList
                        If AppInitLoad32 <> WinDir & "\UxThemeSignatureBypass32.dll" Then
                            If AppInitKeyList32.Contains(" " & WinDir & "\UxThemeSignatureBypass32.dll") Then
                                AppInitKeyList32.RemoveAt(AppInitKeyList32.IndexOf(" " & WinDir & "\UxThemeSignatureBypass32.dll"))
                                Dim StringWithoutUXTSB32 As String = String.Join(",", AppInitKeyList32)
                                RegKey32.SetValue("AppInit_DLLs", StringWithoutUXTSB32)
                            End If
                            RegKey32.Close()
                        Else
                            RegKey32.SetValue("AppInit_DLLs", "")
                            If RegKey32.GetValue("LoadAppInit_DLLs") = 1 Then
                                RegKey32.SetValue("LoadAppInit_DLLs", 0)
                            End If
                            RegKey32.Close()
                        End If
                    ElseIf UninstallWizard.BitnessSystem = "32bit" Then
                        RegKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\Microsoft\Windows NT\CurrentVersion\Windows", True)
                        Dim AppInitLoad As String = RegKey.GetValue("AppInit_DLLs")
                        Dim AppInitKeyList As List(Of String) = AppInitLoad.Split(",").ToList
                        If AppInitLoad <> WinDir & "\UxThemeSignatureBypass32.dll" Then
                            If AppInitKeyList.Contains(" " & WinDir & "\UxThemeSignatureBypass32.dll") Then
                                AppInitKeyList.RemoveAt(AppInitKeyList.IndexOf(" " & WinDir & "\UxThemeSignatureBypass32.dll"))
                                Dim StringWithoutUXTSB As String = String.Join(",", AppInitKeyList)
                                RegKey.SetValue("AppInit_DLLs", StringWithoutUXTSB)
                            End If
                            RegKey.Close()
                        End If
                    End If
                    ProgressBarUninst.PerformStep()
                    ProgressBarUninst.Refresh()
                    ActionLabel.Text = "UXThemeSignatureBypass has been removed!"
                    ActionLabel.Refresh()
                Catch ex As Exception
                    MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    ActionLabel.Text = "UXThemeSignatureBypass could not be removed!"
                    ActionLabel.Refresh()
                    ProgressBarUninst.PerformStep()
                    ProgressBarUninst.Refresh()
                End Try
            End If
        End If

        ' Remove Quero Toolbar
        If UninstallWizard.UninstQuero = True Then
            If FileIO.FileSystem.DirectoryExists(ProgramFiles & "\Quero Toolbar") Then
                '' Restore address bar of Internet Explorer
                Dim HklmIENoNavBar As RegistryKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\Policies\Microsoft\Internet Explorer\ToolBars\Restrictions", True)
                Dim HkcuIENoNavBar As RegistryKey = Registry.CurrentUser.OpenSubKey("Software\Policies\Microsoft\Internet Explorer\Toolbars\Restrictions", True)
                If HklmIENoNavBar IsNot Nothing Then
                    If HklmIENoNavBar.GetValue("NoNavBar") IsNot Nothing Then
                        HklmIENoNavBar.DeleteValue("NoNavBar")
                    End If
                    HklmIENoNavBar.Close()
                End If
                If HkcuIENoNavBar IsNot Nothing Then
                    If HkcuIENoNavBar.GetValue("NoNavBar") IsNot Nothing Then
                        HkcuIENoNavBar.DeleteValue("NoNavBar")
                    End If
                    HkcuIENoNavBar.Close()
                End If
                '' Run uninstaller
                Dim QueroUninstall As New Process
                ActionLabel.Text = "Uninstalling Quero Toolbar..."
                ActionLabel.Refresh()
                With QueroUninstall
                    .StartInfo.FileName = ProgramFiles & "\Quero Toolbar\unins000.exe"
                    .StartInfo.Arguments = "/SILENT /SUPPRESSMSGBOXES"
                    .Start()
                    .WaitForExit()
                End With
                ProgressBarUninst.PerformStep()
                ProgressBarUninst.Refresh()
                ActionLabel.Text = "Quero Toolbar has been uninstalled!"
                ActionLabel.Refresh()
            Else
                MessageBox.Show("Quero Toolbar could not be found on this system!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                ActionLabel.Text = "Quero Toolbar could not be removed!"
                ActionLabel.Refresh()
                ProgressBarUninst.PerformStep()
                ProgressBarUninst.Refresh()
            End If
        End If

        ' Restore UIRibbon.dll & UIRibbonRes.dll
        If UninstallWizard.UninstSysFiles = True Then
            Dim SysFilesSetup As New Process
            ActionLabel.Text = "Restoring UIRibbon.dll && UIRibbonRes.dll..."
            ActionLabel.Refresh()
            If UninstallWizard.BitnessSystem = "64bit" Then
                '' Replace files for 64-bit
                If FileIO.FileSystem.FileExists(WinDir & "\System32\UIRibbon.dll.bak") And FileIO.FileSystem.FileExists(WinDir & "\System32\UIRibbonRes.dll.bak") And FileIO.FileSystem.FileExists(WinDir & "\System32\" & UninstallWizard.Language & "\UIRibbon.dll.mui.bak") And FileIO.FileSystem.FileExists(WinDir & "\SysWOW64\UIRibbon.dll.bak") And FileIO.FileSystem.FileExists(WinDir & "\SysWOW64\UIRibbonRes.dll.bak") And FileIO.FileSystem.FileExists(WinDir & "\SysWOW64\" & UninstallWizard.Language & "\UIRibbon.dll.mui.bak") Then
                    '' End explorer.exe task
                    Dim EndExplorerTask As New Process
                    With EndExplorerTask
                        .StartInfo.FileName = WinDir & "\System32\taskkill.exe"
                        .StartInfo.Arguments = "/f /im explorer.exe"
                        .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                        .Start()
                        .WaitForExit()
                    End With
                    '' Replace files
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
                        SysFilesSetup.WaitForExit()
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
                        .StartInfo.Arguments = "/f " & WinDir & "\System32\" & UninstallWizard.Language & "\UIRibbon.dll.mui /a"
                        .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                        .Start()
                        .WaitForExit()
                    End With
                    With SysFilesSetup
                        .StartInfo.FileName = WinDir & "\System32\icacls.exe"
                        .StartInfo.Arguments = WinDir & "\System32\" & UninstallWizard.Language & "\UIRibbon.dll.mui /grant *S-1-5-32-545:F *S-1-5-32-544:F"
                        .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                        .Start()
                        .WaitForExit()
                    End With
                    With SysFilesSetup
                        .StartInfo.FileName = WinDir & "\System32\takeown.exe"
                        .StartInfo.Arguments = "/f " & WinDir & "\System32\" & UninstallWizard.Language & "\UIRibbon.dll.mui.bak /a"
                        .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                        .Start()
                        .WaitForExit()
                    End With
                    With SysFilesSetup
                        .StartInfo.FileName = WinDir & "\System32\icacls.exe"
                        .StartInfo.Arguments = WinDir & "\System32\" & UninstallWizard.Language & "\UIRibbon.dll.mui.bak /grant *S-1-5-32-545:F *S-1-5-32-544:F"
                        .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                        .Start()
                        .WaitForExit()
                    End With
                    FileIO.FileSystem.DeleteFile(WinDir & "\System32\" & UninstallWizard.Language & "\UIRibbon.dll.mui")
                    FileIO.FileSystem.RenameFile(WinDir & "\System32\" & UninstallWizard.Language & "\UIRibbon.dll.mui.bak", "UIRibbon.dll.mui")
                    '' Original files change to TrustedInstaller
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
                    With SysFilesSetup
                        .StartInfo.FileName = WinDir & "\System32\icacls.exe"
                        .StartInfo.Arguments = WinDir & "\System32\" & UninstallWizard.Language & "\UIRibbon.dll.mui /setowner ""NT SERVICE\TrustedInstaller"""
                        .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                        .Start()
                        .WaitForExit()
                    End With
                    With SysFilesSetup
                        .StartInfo.FileName = WinDir & "\System32\icacls.exe"
                        .StartInfo.Arguments = WinDir & "\System32\" & UninstallWizard.Language & "\UIRibbon.dll.mui /reset"
                        .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                        .Start()
                        .WaitForExit()
                    End With
                    With SysFilesSetup
                        .StartInfo.FileName = WinDir & "\System32\icacls.exe"
                        .StartInfo.Arguments = WinDir & "\System32\" & UninstallWizard.Language & "\UIRibbon.dll.mui /grant *S-1-5-18:RX *S-1-5-32-544:RX *S-1-5-32-545:RX *S-1-15-2-1:RX ""NT SERVICE\TrustedInstaller"":F /inheritance:r"
                        .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                        .Start()
                        .WaitForExit()
                    End With
                    '' UIRibbon.dll && UIRibbonRes.dll\SysWOW64
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
                        .StartInfo.Arguments = "/f " & WinDir & "\SysWOW64\" & UninstallWizard.Language & "\UIRibbon.dll.mui /a"
                        .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                        .Start()
                        .WaitForExit()
                    End With
                    With SysFilesSetup
                        .StartInfo.FileName = WinDir & "\System32\icacls.exe"
                        .StartInfo.Arguments = WinDir & "\SysWOW64\" & UninstallWizard.Language & "\UIRibbon.dll.mui /grant *S-1-5-32-545:F *S-1-5-32-544:F"
                        .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                        .Start()
                        .WaitForExit()
                    End With
                    With SysFilesSetup
                        .StartInfo.FileName = WinDir & "\System32\takeown.exe"
                        .StartInfo.Arguments = "/f " & WinDir & "\SysWOW64\" & UninstallWizard.Language & "\UIRibbon.dll.mui.bak /a"
                        .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                        .Start()
                        .WaitForExit()
                    End With
                    With SysFilesSetup
                        .StartInfo.FileName = WinDir & "\System32\icacls.exe"
                        .StartInfo.Arguments = WinDir & "\SysWOW64\" & UninstallWizard.Language & "\UIRibbon.dll.mui.bak /grant *S-1-5-32-545:F *S-1-5-32-544:F"
                        .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                        .Start()
                        .WaitForExit()
                    End With
                    FileIO.FileSystem.DeleteFile(WinDir & "\SysWOW64\" & UninstallWizard.Language & "\UIRibbon.dll.mui")
                    FileIO.FileSystem.RenameFile(WinDir & "\SysWOW64\" & UninstallWizard.Language & "\UIRibbon.dll.mui.bak", "UIRibbon.dll.mui")
                    '' Original files change to TrustedInstaller
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
                    With SysFilesSetup
                        .StartInfo.FileName = WinDir & "\System32\icacls.exe"
                        .StartInfo.Arguments = WinDir & "\SysWOW64\" & UninstallWizard.Language & "\UIRibbon.dll.mui /setowner ""NT SERVICE\TrustedInstaller"""
                        .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                        .Start()
                        .WaitForExit()
                    End With
                    With SysFilesSetup
                        .StartInfo.FileName = WinDir & "\System32\icacls.exe"
                        .StartInfo.Arguments = WinDir & "\SysWOW64\" & UninstallWizard.Language & "\UIRibbon.dll.mui /reset"
                        .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                        .Start()
                        .WaitForExit()
                    End With
                    With SysFilesSetup
                        .StartInfo.FileName = WinDir & "\System32\icacls.exe"
                        .StartInfo.Arguments = WinDir & "\SysWOW64\" & UninstallWizard.Language & "\UIRibbon.dll.mui /grant *S-1-5-18:RX *S-1-5-32-544:RX *S-1-5-32-545:RX *S-1-15-2-1:RX ""NT SERVICE\TrustedInstaller"":F /inheritance:r"
                        .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                        .Start()
                        .WaitForExit()
                    End With
                    ProgressBarUninst.PerformStep()
                    ProgressBarUninst.Refresh()
                    ActionLabel.Text = "UIRibbon.dll && UIRibbonRes.dll have been restored!"
                    ActionLabel.Refresh()
                    '' Start userinit.exe
                    UninstallUninstallExplorerStartInfoWindowEnglish.Show()
                    Dim StartExplorerTask As New Process
                    With StartExplorerTask
                        .StartInfo.FileName = WinDir & "\System32\userinit.exe"
                        .Start()
                        .WaitForExit()
                    End With
                    UninstallUninstallExplorerStartInfoWindowEnglish.Dispose()
                Else
                    MessageBox.Show("Backup files for UIRibbon.dll && UIRibbonRes.dll missing!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    ProgressBarUninst.PerformStep()
                    ProgressBarUninst.Refresh()
                    ActionLabel.Text = "UIRibbon.dll && UIRibbonRes.dll could not be restored!"
                    ActionLabel.Refresh()
                End If
            ElseIf UninstallWizard.BitnessSystem = "32bit" Then
                If FileIO.FileSystem.FileExists(WinDir & "\System32\UIRibbon.dll.bak") And FileIO.FileSystem.FileExists(WinDir & "\System32\UIRibbonRes.dll.bak") And FileIO.FileSystem.FileExists(WinDir & "\System32\" & UninstallWizard.Language & "\UIRibbon.dll.mui.bak") Then
                    '' End explorer.exe task
                    Dim EndExplorerTask As New Process
                    With EndExplorerTask
                        .StartInfo.FileName = WinDir & "\System32\taskkill.exe"
                        .StartInfo.Arguments = "/f /im explorer.exe"
                        .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                        .Start()
                        .WaitForExit()
                    End With
                    '' Replace files
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
                        .StartInfo.Arguments = "/f " & WinDir & "\System32\" & UninstallWizard.Language & "\UIRibbon.dll.mui /a"
                        .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                        .Start()
                        .WaitForExit()
                    End With
                    With SysFilesSetup
                        .StartInfo.FileName = WinDir & "\System32\icacls.exe"
                        .StartInfo.Arguments = WinDir & "\System32\" & UninstallWizard.Language & "\UIRibbon.dll.mui /grant *S-1-5-32-545:F *S-1-5-32-544:F"
                        .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                        .Start()
                        .WaitForExit()
                    End With
                    With SysFilesSetup
                        .StartInfo.FileName = WinDir & "\System32\takeown.exe"
                        .StartInfo.Arguments = "/f " & WinDir & "\System32\" & UninstallWizard.Language & "\UIRibbon.dll.mui.bak /a"
                        .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                        .Start()
                        .WaitForExit()
                    End With
                    With SysFilesSetup
                        .StartInfo.FileName = WinDir & "\System32\icacls.exe"
                        .StartInfo.Arguments = WinDir & "\System32\" & UninstallWizard.Language & "\UIRibbon.dll.mui.bak /grant *S-1-5-32-545:F *S-1-5-32-544:F"
                        .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                        .Start()
                        .WaitForExit()
                    End With
                    FileIO.FileSystem.DeleteFile(WinDir & "\System32\" & UninstallWizard.Language & "\UIRibbon.dll.mui")
                    FileIO.FileSystem.RenameFile(WinDir & "\System32\" & UninstallWizard.Language & "\UIRibbon.dll.mui.bak", "UIRibbon.dll.mui")
                    '' Original files change to TrustedInstaller
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
                    With SysFilesSetup
                        .StartInfo.FileName = WinDir & "\System32\icacls.exe"
                        .StartInfo.Arguments = WinDir & "\System32\" & UninstallWizard.Language & "\UIRibbon.dll.mui /setowner ""NT SERVICE\TrustedInstaller"""
                        .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                        .Start()
                        .WaitForExit()
                    End With
                    With SysFilesSetup
                        .StartInfo.FileName = WinDir & "\System32\icacls.exe"
                        .StartInfo.Arguments = WinDir & "\System32\" & UninstallWizard.Language & "\UIRibbon.dll.mui /reset"
                        .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                        .Start()
                        .WaitForExit()
                    End With
                    With SysFilesSetup
                        .StartInfo.FileName = WinDir & "\System32\icacls.exe"
                        .StartInfo.Arguments = WinDir & "\System32\" & UninstallWizard.Language & "\UIRibbon.dll.mui /grant *S-1-5-18:RX *S-1-5-32-544:RX *S-1-5-32-545:RX *S-1-15-2-1:RX ""NT SERVICE\TrustedInstaller"":F /inheritance:r"
                        .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                        .Start()
                        .WaitForExit()
                    End With
                    ProgressBarUninst.PerformStep()
                    ProgressBarUninst.Refresh()
                    ActionLabel.Text = "UIRibbon.dll && UIRibbonRes.dll have been restored!"
                    ActionLabel.Refresh()
                    '' Start userinit.exe
                    UninstallUninstallExplorerStartInfoWindowEnglish.Show()
                    Dim StartExplorerTask As New Process
                    With StartExplorerTask
                        .StartInfo.FileName = WinDir & "\System32\userinit.exe"
                        .Start()
                        .WaitForExit()
                    End With
                    UninstallUninstallExplorerStartInfoWindowEnglish.Dispose()
                Else
                    MessageBox.Show("Backup files for UIRibbon.dll && UIRibbonRes.dll missing!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    ProgressBarUninst.PerformStep()
                    ProgressBarUninst.Refresh()
                    ActionLabel.Text = "UIRibbon.dll && UIRibbonRes.dll could not be restored!"
                    ActionLabel.Refresh()
                End If
            End If
        End If

        ' Remove sounds and sound scheme
        If UninstallWizard.UninstSysFiles = True Then
            ActionLabel.Text = "Removing sounds..."
            ActionLabel.Refresh()
            Dim SoundSchemeNamesRegKey As RegistryKey = Registry.CurrentUser.OpenSubKey("AppEvents\Schemes\Names", True)
            Dim SoundSchemeEventsRegKey As RegistryKey = Registry.CurrentUser.OpenSubKey("AppEvents\Schemes\Apps", True)
            If FileIO.FileSystem.DirectoryExists(WinDir & "\Media\Windows 8 Release Preview") Then
                FileIO.FileSystem.DeleteDirectory(WinDir & "\Media\Windows 8 Release Preview", FileIO.DeleteDirectoryOption.DeleteAllContents)
                '' Remove sound scheme from registry if available
                Dim ReleasePreviewSoundSchemeKeyName As String = String.Empty
                For Each SoundScheme In SoundSchemeNamesRegKey.GetSubKeyNames
                    If SoundSchemeNamesRegKey.OpenSubKey(SoundScheme).GetValue(String.Empty) = "Windows 8 Release Preview" Then
                        ReleasePreviewSoundSchemeKeyName = SoundScheme
                    End If
                Next
                '' Remove sound scheme from events if available
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
                ProgressBarUninst.PerformStep()
                ProgressBarUninst.Refresh()
                ActionLabel.Text = "Sounds have been removed!"
                ActionLabel.Refresh()
            Else
                MessageBox.Show("Sound files could not be found on this system!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                ActionLabel.Text = "Sounds could not be removed!"
                ActionLabel.Refresh()
                ProgressBarUninst.PerformStep()
                ProgressBarUninst.Refresh()
            End If
        End If

        ' Uninstall 8GadgetPack
        If UninstallWizard.UninstGadgets = True Then
            ActionLabel.Text = "Uninstalling 8GadgetPack..."
            ActionLabel.Refresh()
            If FileIO.FileSystem.FileExists(ProgramFiles & "\Windows Sidebar\8GadgetPack.exe") Then
                Dim GadgetsUninstall As New Process
                With GadgetsUninstall
                    .StartInfo.FileName = WinDir & "\System32\msiexec.exe"
                    .StartInfo.Arguments = "/X{B6AF19AD-2D5B-44DC-9272-EC91965123E8} /passive"
                    .Start()
                    .WaitForExit()
                End With
                ProgressBarUninst.PerformStep()
                ProgressBarUninst.Refresh()
                ActionLabel.Text = "8GadgetPack has been uninstalled!"
                ActionLabel.Refresh()
            Else
                MessageBox.Show("8GadgetPack could not be found on this system!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                ActionLabel.Text = "Gadgets could not be removed!"
                ActionLabel.Refresh()
                ProgressBarUninst.PerformStep()
                ProgressBarUninst.Refresh()
            End If
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
        '' Delete OldNewExplorer directory after reboot
        If UninstallWizard.UninstONE = True Then
            Dim DeleteONERegKey As RegistryKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce", True)
            DeleteONERegKey.SetValue("Remove OldNewExplorer", """" & WinDir & "\System32\cmd.exe"" /c ""rmdir /s /q """ & ProgramFiles & "\OldNewExplorer""""")
            DeleteONERegKey.Close()
        End If
        '' Delete program folder after reboot
        Dim DeleteProgramRegKey As RegistryKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce", True)
        DeleteProgramRegKey.SetValue("Remove Windows 8.1 to 8 RP Converter", """" & WinDir & "\System32\cmd.exe"" /c ""rmdir /s /q """ & ProgDir & """""")
        DeleteProgramRegKey.Close()
        '' Remove entry from uninstall list in registry
        Dim RegKeyControlPanel As RegistryKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall", True)
        If RegKeyControlPanel.OpenSubKey("Windows 8.1 to 8 RP Converter") IsNot Nothing Then
            RegKeyControlPanel.DeleteSubKey("Windows 8.1 to 8 RP Converter")
        End If
        RegKeyControlPanel.Close()
        ProgressBarUninst.PerformStep()
        ProgressBarUninst.Refresh()
        ActionLabel.Text = "Theme and program(s) have been successfully removed!"
        ActionLabel.Refresh()
        '' Change Next button and mark setup as finished
        NextButton.Text = "&Finish"
        NextButton.Enabled = True
        IsFinished = True
    End Sub

    Private Sub NextButton_Click(sender As Object, e As EventArgs) Handles NextButton.Click
        If IsFinished = True Then
            Me.Close()
        Else
            MessageBox.Show("The setup cannot be finished right now.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End If
    End Sub

    Private Sub UninstallUninstallEnglish_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        If IsFinished = True Then
            If e.CloseReason = CloseReason.UserClosing Then
                ' Restart prompt
                Dim ShutDownProcess As New Process
                If MessageBox.Show("You need to restart your system to finish the uninstallation. Do you want to restart now?", "Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) = System.Windows.Forms.DialogResult.Yes Then
                    With ShutDownProcess
                        .StartInfo.FileName = WinDir & "\System32\shutdown.exe"
                        .StartInfo.Arguments = "/r /t 0"
                        .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                        .Start()
                    End With
                    UninstallWizard.Close()
                ElseIf System.Windows.Forms.DialogResult.No Then
                    UninstallWizard.Close()
                End If
            End If
        Else
            MessageBox.Show("The setup cannot be finished right now.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End If
    End Sub
End Class