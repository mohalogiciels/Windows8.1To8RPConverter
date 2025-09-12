Imports System.IO
Imports System.IO.Compression
Imports System.ComponentModel
Imports Microsoft.Win32

Public Class ModifyModifyEnglish
    Dim ProgramFiles = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles)
    Dim ProgramFilesX86 = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86)
    Dim WinDir = Environment.GetFolderPath(Environment.SpecialFolder.Windows)
    Dim WithEvents BgWorker As New BackgroundWorker
    Dim IsFinished As Boolean = False
    Dim ProgBarCounter As Integer = 10

    Private Sub Wait(ByVal seconds As Integer)
        For Index As Integer = 0 To seconds * 100
            System.Threading.Thread.Sleep(10)
            Application.DoEvents()
        Next
    End Sub

    Private Sub ModifyModifyEnglish_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown
        ' Set progress bar to marquee
        With ProgressBarMod
            .Style = ProgressBarStyle.Marquee
            .MarqueeAnimationSpeed = 100
            .Refresh()
        End With

        ' Initialise progress bar
        If ModifyWizard.ChangeONEChecked = True Then
            ProgBarCounter += 10
        End If
        If ModifyWizard.ChangeUXMethodChecked = True Then
            ProgBarCounter += 20
        End If
        If ModifyWizard.ChangeQueroChecked = True Then
            ProgBarCounter += 10
        End If
        If ModifyWizard.ChangeSysFilesChecked = True Then
            ProgBarCounter += 10
        End If
        If ModifyWizard.InstallSoundsChecked = True Then
            ProgBarCounter += 10
        End If
        If ModifyWizard.ChangeGadgetsChecked = True Then
            ProgBarCounter += 10
        End If
        If ModifyWizard.ChangeAddBarStyle = True Then
            ProgBarCounter += 10
        End If

        ' Check if theme change is needed
        If ModifyWizard.IsThemeChangeNeeded = True Then
            Me.TopMost = True
            BgWorker.RunWorkerAsync()
        ElseIf ModifyWizard.IsThemeChangeNeeded = False Then
            '' Initialise progress bar
            With ProgressBarMod
                .Style = ProgressBarStyle.Continuous
                .Value = 0
                .Maximum = ProgBarCounter
                .Refresh()
            End With
            '' Start modifying
            Try
                RunModify()
            Catch ex As Exception
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End If
    End Sub

    Private Sub BgWorker_DoWork(sender As Object, e As DoWorkEventArgs) Handles BgWorker.DoWork
        ' Change to standard theme
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
        With ProgressBarMod
            .Style = ProgressBarStyle.Continuous
            .Value = 0
            .Maximum = ProgBarCounter
            .Refresh()
        End With

        ' Start modifying now
        Try
            RunModify()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub RunModify()
        Dim ProgDir As String = ProgramFiles & "\Windows8.1to8RPConv"
        Dim ResFolder As String = ProgDir & "\res"

        ' If theme change needed put TopMost back to standard
        If ModifyWizard.IsThemeChangeNeeded = True Then
            Me.TopMost = False
        End If

        ' Change OldNewExplorer
        Dim SourcePathONE As String = ResFolder & "\OldNewExplorer"
        Dim DestPathONE As String = ProgramFiles & "\OldNewExplorer"
        If ModifyWizard.ChangeONEChecked = True Then
            If ModifyWizard.InstONE = True Then
                '' Uninstall OldNewExplorer
                Dim ONESetup As New Process
                ActionLabel.Text = "Removing OldNewExplorer..."
                ActionLabel.Refresh()
                '' Unreg OldNewExplorer DLLs
                If ModifyWizard.BitnessSystem = "64bit" Then
                    With ONESetup
                        .StartInfo.FileName = WinDir & "\System32\regsvr32.exe"
                        .StartInfo.Arguments = "/u /s """ & DestPathONE & "\OldNewExplorer64.dll"""
                        .Start()
                        .WaitForExit()
                    End With
                    With ONESetup
                        .StartInfo.FileName = WinDir & "\System32\regsvr32.exe"
                        .StartInfo.Arguments = "/u /s """ & DestPathONE & "\OldNewExplorer32.dll"""
                        .Start()
                        .WaitForExit()
                    End With
                ElseIf ModifyWizard.BitnessSystem = "32bit" Then
                    With ONESetup
                        .StartInfo.FileName = WinDir & "\System32\regsvr32.exe"
                        .StartInfo.Arguments = "/u /s """ & DestPathONE & "\OldNewExplorer32.dll"""
                        .Start()
                        .WaitForExit()
                    End With
                End If
                ProgressBarMod.PerformStep()
                ProgressBarMod.Refresh()
                ActionLabel.Text = "OldNewExplorer has been removed!"
                ActionLabel.Refresh()
            ElseIf ModifyWizard.InstONE = False Then
                '' Install OldNewExplorer
                Dim ONESetup As New Process
                ActionLabel.Text = "Installing OldNewExplorer..."
                ActionLabel.Refresh()
                FileIO.FileSystem.CopyDirectory(SourcePathONE, DestPathONE)
                If ModifyWizard.BitnessSystem = "64bit" Then
                    With ONESetup
                        .StartInfo.FileName = WinDir & "\System32\regsvr32.exe"
                        .StartInfo.Arguments = "/s """ & DestPathONE & "\OldNewExplorer64.dll"""
                        .Start()
                        .WaitForExit()
                    End With
                    With ONESetup
                        .StartInfo.FileName = WinDir & "\System32\regsvr32.exe"
                        .StartInfo.Arguments = "/s """ & DestPathONE & "\OldNewExplorer32.dll"""
                        .Start()
                        .WaitForExit()
                    End With
                ElseIf ModifyWizard.BitnessSystem = "32bit" Then
                    With ONESetup
                        .StartInfo.FileName = WinDir & "\System32\regsvr32.exe"
                        .StartInfo.Arguments = "/s """ & DestPathONE & "\OldNewExplorer32.dll"""
                        .Start()
                        .WaitForExit()
                    End With
                End If
                '' Create shortcut on desktop
                Dim ObjectShell As Object
                Dim ObjectLink As Object
                Try
                    ObjectShell = CreateObject("WScript.Shell")
                    ObjectLink = ObjectShell.CreateShortcut(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) & "\OldNewExplorer Configuration.lnk")
                    With ObjectLink
                        .TargetPath = ProgramFiles & "\OldNewExplorer\OldNewExplorerCfg.exe"
                        .WorkingDirectory = ProgramFiles & "\OldNewExplorer"
                        .WindowStyle = 1
                        .Save()
                    End With
                Catch ex As Exception
                    MessageBox.Show("Shortcut to OldNewExplorer configuration program could not be created!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Finally
                    Application.DoEvents()
                End Try
                ProgressBarMod.PerformStep()
                ProgressBarMod.Refresh()
                ActionLabel.Text = "OldNewExplorer has been installed!"
                ActionLabel.Refresh()
            End If
        End If

        ' Change UXTheme patcher
        If ModifyWizard.ChangeUXMethodChecked = True Then
            If ModifyWizard.InstUX = "UltraUX" Then
                '' Uninstall UltraUX
                ActionLabel.Text = "Uninstalling UltraUXThemePatcher..."
                ActionLabel.Refresh()
                If ModifyWizard.BitnessSystem = "64bit" Then
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
                ElseIf ModifyWizard.BitnessSystem = "32bit" Then
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
                ProgressBarMod.PerformStep()
                ProgressBarMod.Refresh()
                ActionLabel.Text = "UltraUXThemePatcher has been uninstalled!"
                ActionLabel.Refresh()
                '' Install UXTSB
                Dim SourcePathUXTSB As String = ResFolder & "\UXThemeDLL"
                ActionLabel.Text = "Installing UXThemeSignatureBypass..."
                ActionLabel.Refresh()
                If ModifyWizard.BitnessSystem = "64bit" Then
                    FileIO.FileSystem.CopyDirectory(SourcePathUXTSB, WinDir)
                ElseIf ModifyWizard.BitnessSystem = "32bit" Then
                    FileIO.FileSystem.CopyFile(SourcePathUXTSB & "\UxThemeSignatureBypass32.dll", WinDir & "\UxThemeSignatureBypass32.dll")
                End If
                '' UXTSB set registry key to load file with Windows
                Dim RegKey As RegistryKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\Microsoft\Windows NT\CurrentVersion\Windows", True)
                Dim RegKey32 As RegistryKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\Wow6432Node\Microsoft\Windows NT\CurrentVersion\Windows", True)
                If ModifyWizard.BitnessSystem = "64bit" Then
                    If RegKey.GetValue("AppInit_DLLs") <> String.Empty Then
                        RegKey.SetValue("AppInit_DLLs", RegKey.GetValue("AppInit_DLLs") & ", " & WinDir & "\UxThemeSignatureBypass64.dll")
                    Else
                        RegKey.SetValue("AppInit_DLLs", WinDir & "\UxThemeSignatureBypass64.dll")
                    End If
                    If RegKey.GetValue("LoadAppInit_DLLs") <> 1 Then
                        RegKey.SetValue("LoadAppInit_DLLs", 1)
                    End If
                    RegKey.Close()
                    '' UXTSB\Registry\SysWOW64
                    If RegKey32.GetValue("AppInit_DLLs") <> String.Empty Then
                        RegKey32.SetValue("AppInit_DLLs", RegKey.GetValue("AppInit_DLLs") & ", " & WinDir & "\UxThemeSignatureBypass32.dll")
                    Else
                        RegKey32.SetValue("AppInit_DLLs", WinDir & "\UxThemeSignatureBypass32.dll")
                    End If
                    If RegKey32.GetValue("LoadAppInit_DLLs") <> 1 Then
                        RegKey32.SetValue("LoadAppInit_DLLs", 1)
                    End If
                    RegKey32.Close()
                ElseIf ModifyWizard.BitnessSystem = "32bit" Then
                    If RegKey.GetValue("AppInit_DLLs") <> String.Empty Then
                        RegKey.SetValue("AppInit_DLLs", RegKey.GetValue("AppInit_DLLs") & ", " & WinDir & "\UxThemeSignatureBypass32.dll")
                    Else
                        RegKey.SetValue("AppInit_DLLs", WinDir & "\UxThemeSignatureBypass32.dll")
                    End If
                    If RegKey.GetValue("LoadAppInit_DLLs") <> 1 Then
                        RegKey.SetValue("LoadAppInit_DLLs", 1)
                    End If
                    RegKey.Close()
                End If
                ProgressBarMod.PerformStep()
                ProgressBarMod.Refresh()
                ActionLabel.Text = "UXThemeSignatureBypass has been installed!"
                ActionLabel.Refresh()
            ElseIf ModifyWizard.InstUX = "UXTSB" Then
                ActionLabel.Text = "Removing UXThemeSignatureBypass..."
                ActionLabel.Refresh()
                '' Removing UXTSB after reboot
                If ModifyWizard.BitnessSystem = "64bit" Then
                    Dim RegKeyDelete As RegistryKey = Registry.LocalMachine.OpenSubKey("SYSTEM\CurrentControlSet\Control\Session Manager", True)
                    If RegKeyDelete.GetValue("PendingFileRenameOperations") IsNot Nothing Then
                        Dim RegKeyValue As String() = RegKeyDelete.GetValue("PendingFileRenameOperations")
                        Dim RegKeyValueString As String = String.Join(Chr(0), RegKeyValue)
                        Dim DeleteFiles As String() = {RegKeyValueString, "\??\" & WinDir & "\UxThemeSignatureBypass64.dll" & Chr(0), "\??\" & WinDir & "\UxThemeSignatureBypass32.dll" & Chr(0)}
                        RegKeyDelete.SetValue("PendingFileRenameOperations", DeleteFiles, RegistryValueKind.MultiString)
                    Else
                        Dim DeleteFiles As String() = {"\??\" & WinDir & "\UxThemeSignatureBypass64.dll" & Chr(0), "\??\" & WinDir & "\UxThemeSignatureBypass32.dll" & Chr(0)}
                        RegKeyDelete.SetValue("PendingFileRenameOperations", DeleteFiles, RegistryValueKind.MultiString)
                    End If
                    RegKeyDelete.Close()
                ElseIf ModifyWizard.BitnessSystem = "32bit" Then
                    Dim RegKeyDelete As RegistryKey = Registry.LocalMachine.OpenSubKey("SYSTEM\CurrentControlSet\Control\Session Manager", True)
                    If RegKeyDelete.GetValue("PendingFileRenameOperations") IsNot Nothing Then
                        Dim RegKeyValue As String() = RegKeyDelete.GetValue("PendingFileRenameOperations")
                        Dim RegKeyValueString As String = String.Join(Chr(0), RegKeyValue)
                        Dim DeleteFiles As String() = {RegKeyValueString, "\??\" & WinDir & "\UxThemeSignatureBypass32.dll" & Chr(0)}
                        RegKeyDelete.SetValue("PendingFileRenameOperations", DeleteFiles, RegistryValueKind.MultiString)
                    Else
                        Dim DeleteFiles As String() = {"\??\" & WinDir & "\UxThemeSignatureBypass32.dll" & Chr(0)}
                        RegKeyDelete.SetValue("PendingFileRenameOperations", DeleteFiles, RegistryValueKind.MultiString)
                    End If
                    RegKeyDelete.Close()
                End If
                '' UXTSB set registry key to load file with Windows
                Dim RegKey As RegistryKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\Microsoft\Windows NT\CurrentVersion\Windows", True)
                Dim RegKey32 As RegistryKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\Wow6432Node\Microsoft\Windows NT\CurrentVersion\Windows", True)
                If ModifyWizard.BitnessSystem = "64bit" Then
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
                ElseIf ModifyWizard.BitnessSystem = "32bit" Then
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
                ProgressBarMod.PerformStep()
                ProgressBarMod.Refresh()
                ActionLabel.Text = "UXThemeSignatureBypass has been removed!"
                ActionLabel.Refresh()
                '' Install UltraUXThemePatcher
                Dim UltraUXSetup As New Process
                ActionLabel.Text = "Installing UltraUXThemePatcher..."
                ActionLabel.Refresh()
                With UltraUXSetup
                    .StartInfo.FileName = ResFolder & "\UXTheme\UltraUXThemePatcher_4.4.4.exe"
                    .Start()
                    .WaitForExit()
                End With
                ProgressBarMod.PerformStep()
                ProgressBarMod.Refresh()
                ActionLabel.Text = "UltraUXThemePatcher has been installed!"
                ActionLabel.Refresh()
            End If
        End If

        ' Change Quero Toolbar
        If ModifyWizard.ChangeQueroChecked = True And Not FileIO.FileSystem.DirectoryExists(ProgramFiles & "\Quero Toolbar") Then
            Dim QueroSetup As New Process
            ActionLabel.Text = "Installing Quero Toolbar..."
            ActionLabel.Refresh()
            '' Change install directory to ProgramFiles folder
            Dim QueroSetupInfFilePath As String = ResFolder & "\Quero\settings.inf"
            Dim QueroSetupInfFile As String() = File.ReadAllLines(QueroSetupInfFilePath)
            QueroSetupInfFile(2) = "Dir=" & ProgramFiles & "\Quero Toolbar"
            File.WriteAllLines(QueroSetupInfFilePath, QueroSetupInfFile)
            '' Start Quero Toolbar setup
            If ModifyWizard.BitnessSystem = "64bit" Then
                With QueroSetup
                    .StartInfo.FileName = ResFolder & "\Quero\QueroToolbarInstaller_x64.exe"
                    .StartInfo.Arguments = "/SILENT /LOADINF=""" & QueroSetupInfFilePath & """"
                    .Start()
                    .WaitForExit()
                End With
            ElseIf ModifyWizard.BitnessSystem = "32bit" Then
                With QueroSetup
                    .StartInfo.FileName = ResFolder & "\Quero\QueroToolbarInstaller_x86.exe"
                    .StartInfo.Arguments = "/SILENT /LOADINF=""" & QueroSetupInfFilePath & """"
                    .Start()
                    .WaitForExit()
                End With
            End If
            File.WriteAllBytes(ResFolder & "\modify.xht", My.Resources.ModifyQueroEnglish)
            ProgressBarMod.PerformStep()
            ProgressBarMod.Refresh()
            ActionLabel.Text = "Quero Toolbar has been installed!"
            ActionLabel.Refresh()
        ElseIf ModifyWizard.ChangeQueroChecked = True And FileIO.FileSystem.DirectoryExists(ProgramFiles & "\Quero Toolbar") Then
            Dim QueroSetup As New Process
            ActionLabel.Text = "Uninstalling Quero Toolbar..."
            ActionLabel.Refresh()
            '' Restore address bar of Internet Explorer
            Dim HklmIENoNavBar As RegistryKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\Policies\Microsoft\Internet Explorer\Toolbars\Restrictions", True)
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
            With QueroSetup
                .StartInfo.FileName = ProgramFiles & "\Quero Toolbar\unins000.exe"
                .StartInfo.Arguments = "/SILENT /SUPPRESSMSGBOXES"
                .Start()
                .WaitForExit()
            End With
            ProgressBarMod.PerformStep()
            ProgressBarMod.Refresh()
            ActionLabel.Text = "Quero Toolbar has been uninstalled!"
            ActionLabel.Refresh()
        End If

        ' Change UIRibbon.dll & UIRibbonRes.dll
        If ModifyWizard.InstSysFiles = True And ModifyWizard.ChangeSysFilesChecked Then
            Dim SysFilesSetup As New Process
            ActionLabel.Text = "Restoring UIRibbon.dll && UIRibbonRes.dll..."
            ActionLabel.Refresh()
            If ModifyWizard.BitnessSystem = "64bit" Then
                '' Replace files for 64-bit
                If FileIO.FileSystem.FileExists(WinDir & "\System32\UIRibbon.dll.bak") And FileIO.FileSystem.FileExists(WinDir & "\System32\UIRibbonRes.dll.bak") And FileIO.FileSystem.FileExists(WinDir & "\System32\" & ModifyWizard.Language & "\UIRibbon.dll.mui.bak") And FileIO.FileSystem.FileExists(WinDir & "\SysWOW64\UIRibbon.dll.bak") And FileIO.FileSystem.FileExists(WinDir & "\SysWOW64\UIRibbonRes.dll.bak") And FileIO.FileSystem.FileExists(WinDir & "\SysWOW64\" & ModifyWizard.Language & "\UIRibbon.dll.mui.bak") Then
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
                        .StartInfo.Arguments = "/f " & WinDir & "\System32\" & ModifyWizard.Language & "\UIRibbon.dll.mui /a"
                        .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                        .Start()
                        .WaitForExit()
                    End With
                    With SysFilesSetup
                        .StartInfo.FileName = WinDir & "\System32\icacls.exe"
                        .StartInfo.Arguments = WinDir & "\System32\" & ModifyWizard.Language & "\UIRibbon.dll.mui /grant *S-1-5-32-545:F *S-1-5-32-544:F"
                        .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                        .Start()
                        .WaitForExit()
                    End With
                    With SysFilesSetup
                        .StartInfo.FileName = WinDir & "\System32\takeown.exe"
                        .StartInfo.Arguments = "/f " & WinDir & "\System32\" & ModifyWizard.Language & "\UIRibbon.dll.mui.bak /a"
                        .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                        .Start()
                        .WaitForExit()
                    End With
                    With SysFilesSetup
                        .StartInfo.FileName = WinDir & "\System32\icacls.exe"
                        .StartInfo.Arguments = WinDir & "\System32\" & ModifyWizard.Language & "\UIRibbon.dll.mui.bak /grant *S-1-5-32-545:F *S-1-5-32-544:F"
                        .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                        .Start()
                        .WaitForExit()
                    End With
                    FileIO.FileSystem.DeleteFile(WinDir & "\System32\" & ModifyWizard.Language & "\UIRibbon.dll.mui")
                    FileIO.FileSystem.RenameFile(WinDir & "\System32\" & ModifyWizard.Language & "\UIRibbon.dll.mui.bak", "UIRibbon.dll.mui")
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
                        .StartInfo.Arguments = WinDir & "\System32\" & ModifyWizard.Language & "\UIRibbon.dll.mui /setowner ""NT SERVICE\TrustedInstaller"""
                        .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                        .Start()
                        .WaitForExit()
                    End With
                    With SysFilesSetup
                        .StartInfo.FileName = WinDir & "\System32\icacls.exe"
                        .StartInfo.Arguments = WinDir & "\System32\" & ModifyWizard.Language & "\UIRibbon.dll.mui /reset"
                        .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                        .Start()
                        .WaitForExit()
                    End With
                    With SysFilesSetup
                        .StartInfo.FileName = WinDir & "\System32\icacls.exe"
                        .StartInfo.Arguments = WinDir & "\System32\" & ModifyWizard.Language & "\UIRibbon.dll.mui /grant *S-1-5-18:RX *S-1-5-32-544:RX *S-1-5-32-545:RX *S-1-15-2-1:RX ""NT SERVICE\TrustedInstaller"":F /inheritance:r"
                        .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                        .Start()
                        .WaitForExit()
                    End With
                    '' UIRibbon.dll, UIRibbonRes.dll\SysWOW64
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
                        .StartInfo.Arguments = "/f " & WinDir & "\SysWOW64\" & ModifyWizard.Language & "\UIRibbon.dll.mui /a"
                        .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                        .Start()
                        .WaitForExit()
                    End With
                    With SysFilesSetup
                        .StartInfo.FileName = WinDir & "\System32\icacls.exe"
                        .StartInfo.Arguments = WinDir & "\SysWOW64\" & ModifyWizard.Language & "\UIRibbon.dll.mui /grant *S-1-5-32-545:F *S-1-5-32-544:F"
                        .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                        .Start()
                        .WaitForExit()
                    End With
                    With SysFilesSetup
                        .StartInfo.FileName = WinDir & "\System32\takeown.exe"
                        .StartInfo.Arguments = "/f " & WinDir & "\SysWOW64\" & ModifyWizard.Language & "\UIRibbon.dll.mui.bak /a"
                        .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                        .Start()
                        .WaitForExit()
                    End With
                    With SysFilesSetup
                        .StartInfo.FileName = WinDir & "\System32\icacls.exe"
                        .StartInfo.Arguments = WinDir & "\SysWOW64\" & ModifyWizard.Language & "\UIRibbon.dll.mui.bak /grant *S-1-5-32-545:F *S-1-5-32-544:F"
                        .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                        .Start()
                        .WaitForExit()
                    End With
                    FileIO.FileSystem.DeleteFile(WinDir & "\SysWOW64\" & ModifyWizard.Language & "\UIRibbon.dll.mui")
                    FileIO.FileSystem.RenameFile(WinDir & "\SysWOW64\" & ModifyWizard.Language & "\UIRibbon.dll.mui.bak", "UIRibbon.dll.mui")
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
                        .StartInfo.Arguments = WinDir & "\SysWOW64\" & ModifyWizard.Language & "\UIRibbon.dll.mui /setowner ""NT SERVICE\TrustedInstaller"""
                        .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                        .Start()
                        .WaitForExit()
                    End With
                    With SysFilesSetup
                        .StartInfo.FileName = WinDir & "\System32\icacls.exe"
                        .StartInfo.Arguments = WinDir & "\SysWOW64\" & ModifyWizard.Language & "\UIRibbon.dll.mui /reset"
                        .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                        .Start()
                        .WaitForExit()
                    End With
                    With SysFilesSetup
                        .StartInfo.FileName = WinDir & "\System32\icacls.exe"
                        .StartInfo.Arguments = WinDir & "\SysWOW64\" & ModifyWizard.Language & "\UIRibbon.dll.mui /grant *S-1-5-18:RX *S-1-5-32-544:RX *S-1-5-32-545:RX *S-1-15-2-1:RX ""NT SERVICE\TrustedInstaller"":F /inheritance:r"
                        .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                        .Start()
                        .WaitForExit()
                    End With
                    ProgressBarMod.PerformStep()
                    ProgressBarMod.Refresh()
                    ActionLabel.Text = "UIRibbon.dll && UIRibbonRes.dll have been restored!"
                    ActionLabel.Refresh()
                    '' Start userinit.exe
                    ModifyModifyExplorerStartInfoWindowEnglish.Show()
                    Dim StartExplorerTask As New Process
                    With StartExplorerTask
                        .StartInfo.FileName = WinDir & "\System32\userinit.exe"
                        .Start()
                        .WaitForExit()
                    End With
                    ModifyModifyExplorerStartInfoWindowEnglish.Dispose()
                Else
                    MessageBox.Show("Backup files for UIRibbon.dll && UIRibbonRes.dll missing!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    ProgressBarMod.PerformStep()
                    ProgressBarMod.Refresh()
                    ActionLabel.Text = "Installation will be continued..."
                    ActionLabel.Refresh()
                End If
            ElseIf ModifyWizard.BitnessSystem = "32bit" Then
                If FileIO.FileSystem.FileExists(WinDir & "\System32\UIRibbon.dll.bak") And FileIO.FileSystem.FileExists(WinDir & "\System32\UIRibbonRes.dll.bak") And FileIO.FileSystem.FileExists(WinDir & "\System32\" & ModifyWizard.Language & "\UIRibbon.dll.mui.bak") Then
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
                        .StartInfo.Arguments = "/f " & WinDir & "\System32\" & ModifyWizard.Language & "\UIRibbon.dll.mui /a"
                        .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                        .Start()
                        .WaitForExit()
                    End With
                    With SysFilesSetup
                        .StartInfo.FileName = WinDir & "\System32\icacls.exe"
                        .StartInfo.Arguments = WinDir & "\System32\" & ModifyWizard.Language & "\UIRibbon.dll.mui /grant *S-1-5-32-545:F *S-1-5-32-544:F"
                        .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                        .Start()
                        .WaitForExit()
                    End With
                    With SysFilesSetup
                        .StartInfo.FileName = WinDir & "\System32\takeown.exe"
                        .StartInfo.Arguments = "/f " & WinDir & "\System32\" & ModifyWizard.Language & "\UIRibbon.dll.mui.bak /a"
                        .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                        .Start()
                        .WaitForExit()
                    End With
                    With SysFilesSetup
                        .StartInfo.FileName = WinDir & "\System32\icacls.exe"
                        .StartInfo.Arguments = WinDir & "\System32\" & ModifyWizard.Language & "\UIRibbon.dll.mui.bak /grant *S-1-5-32-545:F *S-1-5-32-544:F"
                        .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                        .Start()
                        .WaitForExit()
                    End With
                    FileIO.FileSystem.DeleteFile(WinDir & "\System32\" & ModifyWizard.Language & "\UIRibbon.dll.mui")
                    FileIO.FileSystem.RenameFile(WinDir & "\System32\" & ModifyWizard.Language & "\UIRibbon.dll.mui.bak", "UIRibbon.dll.mui")
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
                        .StartInfo.Arguments = WinDir & "\System32\" & ModifyWizard.Language & "\UIRibbon.dll.mui /setowner ""NT SERVICE\TrustedInstaller"""
                        .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                        .Start()
                        .WaitForExit()
                    End With
                    With SysFilesSetup
                        .StartInfo.FileName = WinDir & "\System32\icacls.exe"
                        .StartInfo.Arguments = WinDir & "\System32\" & ModifyWizard.Language & "\UIRibbon.dll.mui /reset"
                        .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                        .Start()
                        .WaitForExit()
                    End With
                    With SysFilesSetup
                        .StartInfo.FileName = WinDir & "\System32\icacls.exe"
                        .StartInfo.Arguments = WinDir & "\System32\" & ModifyWizard.Language & "\UIRibbon.dll.mui /grant *S-1-5-18:RX *S-1-5-32-544:RX *S-1-5-32-545:RX *S-1-15-2-1:RX ""NT SERVICE\TrustedInstaller"":F /inheritance:r"
                        .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                        .Start()
                        .WaitForExit()
                    End With
                    ProgressBarMod.PerformStep()
                    ProgressBarMod.Refresh()
                    ActionLabel.Text = "UIRibbon.dll && UIRibbonRes.dll have been restored!"
                    ActionLabel.Refresh()
                    '' Start userinit.exe
                    ModifyModifyExplorerStartInfoWindowEnglish.Show()
                    Dim StartExplorerTask As New Process
                    With StartExplorerTask
                        .StartInfo.FileName = WinDir & "\System32\userinit.exe"
                        .Start()
                        .WaitForExit()
                    End With
                    ModifyModifyExplorerStartInfoWindowEnglish.Dispose()
                Else
                    MessageBox.Show("Backup files for UIRibbon.dll && UIRibbonRes.dll missing!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    ProgressBarMod.PerformStep()
                    ProgressBarMod.Refresh()
                    ActionLabel.Text = "Installation will be continued..."
                    ActionLabel.Refresh()
                End If
            End If
        ElseIf ModifyWizard.InstSysFiles = False And ModifyWizard.ChangeSysFilesChecked Then
            Dim SysFilesSetup As New Process
            Dim SourcePathSysFiles As String = ResFolder & "\UIRibbon"
            ActionLabel.Text = "Replacing UIRibbon.dll && UIRibbonRes.dll..."
            ActionLabel.Refresh()
            '' End explorer.exe task
            Dim EndExplorerTask As New Process
            With EndExplorerTask
                .StartInfo.FileName = WinDir & "\System32\taskkill.exe"
                .StartInfo.Arguments = "/f /im explorer.exe"
                .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                .Start()
                .WaitForExit()
            End With
            '' Replace files for 64-bit
            If ModifyWizard.BitnessSystem = "64bit" Then
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
                FileIO.FileSystem.RenameFile(WinDir & "\System32\UIRibbon.dll", "UIRibbon.dll.bak")
                FileIO.FileSystem.RenameFile(WinDir & "\System32\UIRibbonRes.dll", "UIRibbonRes.dll.bak")
                With SysFilesSetup
                    .StartInfo.FileName = WinDir & "\System32\takeown.exe"
                    .StartInfo.Arguments = "/f " & WinDir & "\System32\" & ModifyWizard.Language & "\UIRibbon.dll.mui /a"
                    .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                    .Start()
                    .WaitForExit()
                End With
                With SysFilesSetup
                    .StartInfo.FileName = WinDir & "\System32\icacls.exe"
                    .StartInfo.Arguments = WinDir & "\System32\" & ModifyWizard.Language & "\UIRibbon.dll.mui /grant *S-1-5-32-545:F *S-1-5-32-544:F"
                    .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                    .Start()
                    .WaitForExit()
                End With
                FileIO.FileSystem.RenameFile(WinDir & "\System32\" & ModifyWizard.Language & "\UIRibbon.dll.mui", "UIRibbon.dll.mui.bak")
                FileIO.FileSystem.CopyDirectory(SourcePathSysFiles & "\system32", WinDir & "\System32")
                '' New files change to TrustedInstaller
                With SysFilesSetup
                    .StartInfo.FileName = WinDir & "\System32\icacls.exe"
                    .StartInfo.Arguments = WinDir & "\System32\UIRibbon.dll /setowner ""NT SERVICE\TrustedInstaller"""
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
                    .StartInfo.Arguments = WinDir & "\System32\UIRibbonRes.dll /grant *S-1-5-18:RX *S-1-5-32-544:RX *S-1-5-32-545:RX *S-1-15-2-1:RX ""NT SERVICE\TrustedInstaller"":F /inheritance:r"
                    .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                    .Start()
                    .WaitForExit()
                End With
                With SysFilesSetup
                    .StartInfo.FileName = WinDir & "\System32\icacls.exe"
                    .StartInfo.Arguments = WinDir & "\System32\" & ModifyWizard.Language & "\UIRibbon.dll.mui /setowner ""NT SERVICE\TrustedInstaller"""
                    .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                    .Start()
                    .WaitForExit()
                End With
                With SysFilesSetup
                    .StartInfo.FileName = WinDir & "\System32\icacls.exe"
                    .StartInfo.Arguments = WinDir & "\System32\" & ModifyWizard.Language & "\UIRibbon.dll.mui /grant *S-1-5-18:RX *S-1-5-32-544:RX *S-1-5-32-545:RX *S-1-15-2-1:RX ""NT SERVICE\TrustedInstaller"":F /inheritance:r"
                    .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                    .Start()
                    .WaitForExit()
                End With
                '' UIRibbon.dll, UIRibbonRes.dll\SysWOW64
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
                FileIO.FileSystem.RenameFile(WinDir & "\SysWOW64\UIRibbon.dll", "UIRibbon.dll.bak")
                FileIO.FileSystem.RenameFile(WinDir & "\SysWOW64\UIRibbonRes.dll", "UIRibbonRes.dll.bak")
                With SysFilesSetup
                    .StartInfo.FileName = WinDir & "\System32\takeown.exe"
                    .StartInfo.Arguments = "/f " & WinDir & "\SysWOW64\" & ModifyWizard.Language & "\UIRibbon.dll.mui /a"
                    .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                    .Start()
                    .WaitForExit()
                End With
                With SysFilesSetup
                    .StartInfo.FileName = WinDir & "\System32\icacls.exe"
                    .StartInfo.Arguments = WinDir & "\SysWOW64\" & ModifyWizard.Language & "\UIRibbon.dll.mui /grant *S-1-5-32-545:F *S-1-5-32-544:F"
                    .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                    .Start()
                    .WaitForExit()
                End With
                FileIO.FileSystem.RenameFile(WinDir & "\SysWOW64\" & ModifyWizard.Language & "\UIRibbon.dll.mui", "UIRibbon.dll.mui.bak")
                FileIO.FileSystem.CopyDirectory(SourcePathSysFiles & "\syswow64", WinDir & "\SysWOW64")
                '' New files\SysWOW64 change to TrustedInstaller
                With SysFilesSetup
                    .StartInfo.FileName = WinDir & "\System32\icacls.exe"
                    .StartInfo.Arguments = WinDir & "\SysWOW64\UIRibbon.dll /grant *S-1-5-18:RX *S-1-5-32-544:RX *S-1-5-32-545:RX *S-1-15-2-1:RX ""NT SERVICE\TrustedInstaller"":F /inheritance:r"
                    .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                    .Start()
                    .WaitForExit()
                End With
                With SysFilesSetup
                    .StartInfo.FileName = WinDir & "\System32\icacls.exe"
                    .StartInfo.Arguments = WinDir & "\SysWOW64\UIRibbon.dll /setowner ""NT SERVICE\TrustedInstaller"""
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
                    .StartInfo.Arguments = WinDir & "\SysWOW64\UIRibbonRes.dll /setowner ""NT SERVICE\TrustedInstaller"""
                    .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                    .Start()
                    .WaitForExit()
                End With
                With SysFilesSetup
                    .StartInfo.FileName = WinDir & "\System32\icacls.exe"
                    .StartInfo.Arguments = WinDir & "\SysWOW64\" & ModifyWizard.Language & "\UIRibbon.dll.mui /grant *S-1-5-18:RX *S-1-5-32-544:RX *S-1-5-32-545:RX *S-1-15-2-1:RX ""NT SERVICE\TrustedInstaller"":F /inheritance:r"
                    .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                    .Start()
                    .WaitForExit()
                End With
                With SysFilesSetup
                    .StartInfo.FileName = WinDir & "\System32\icacls.exe"
                    .StartInfo.Arguments = WinDir & "\SysWOW64\" & ModifyWizard.Language & "\UIRibbon.dll.mui /setowner ""NT SERVICE\TrustedInstaller"""
                    .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                    .Start()
                    .WaitForExit()
                End With
                ProgressBarMod.PerformStep()
                ProgressBarMod.Refresh()
            ElseIf ModifyWizard.BitnessSystem = "32bit" Then
                '' Replace files for 32-bit
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
                FileIO.FileSystem.RenameFile(WinDir & "\System32\UIRibbon.dll", "UIRibbon.dll.bak")
                FileIO.FileSystem.RenameFile(WinDir & "\System32\UIRibbonRes.dll", "UIRibbonRes.dll.bak")
                With SysFilesSetup
                    .StartInfo.FileName = WinDir & "\System32\takeown.exe"
                    .StartInfo.Arguments = "/f " & WinDir & "\System32\" & ModifyWizard.Language & "\UIRibbon.dll.mui /a"
                    .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                    .Start()
                    .WaitForExit()
                End With
                With SysFilesSetup
                    .StartInfo.FileName = WinDir & "\System32\icacls.exe"
                    .StartInfo.Arguments = WinDir & "\System32\" & ModifyWizard.Language & "\UIRibbon.dll.mui /grant *S-1-5-32-545:F *S-1-5-32-544:F"
                    .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                    .Start()
                    .WaitForExit()
                End With
                FileIO.FileSystem.RenameFile(WinDir & "\System32\" & ModifyWizard.Language & "\UIRibbon.dll.mui", "UIRibbon.dll.mui.bak")
                FileIO.FileSystem.CopyDirectory(SourcePathSysFiles & "\syswow64", WinDir & "\System32")
                '' New files change to TrustedInstaller
                With SysFilesSetup
                    .StartInfo.FileName = WinDir & "\System32\icacls.exe"
                    .StartInfo.Arguments = WinDir & "\System32\UIRibbon.dll /grant *S-1-5-18:RX *S-1-5-32-544:RX *S-1-5-32-545:RX *S-1-15-2-1:RX ""NT SERVICE\TrustedInstaller"":F /inheritance:r"
                    .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                    .Start()
                    .WaitForExit()
                End With
                With SysFilesSetup
                    .StartInfo.FileName = WinDir & "\System32\icacls.exe"
                    .StartInfo.Arguments = WinDir & "\System32\UIRibbon.dll /setowner ""NT SERVICE\TrustedInstaller"""
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
                    .StartInfo.Arguments = WinDir & "\System32\UIRibbonRes.dll /setowner ""NT SERVICE\TrustedInstaller"""
                    .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                    .Start()
                    .WaitForExit()
                End With
                With SysFilesSetup
                    .StartInfo.FileName = WinDir & "\System32\icacls.exe"
                    .StartInfo.Arguments = WinDir & "\System32\" & ModifyWizard.Language & "\UIRibbon.dll.mui /grant *S-1-5-18:RX *S-1-5-32-544:RX *S-1-5-32-545:RX *S-1-15-2-1:RX ""NT SERVICE\TrustedInstaller"":F /inheritance:r"
                    .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                    .Start()
                    .WaitForExit()
                End With
                With SysFilesSetup
                    .StartInfo.FileName = WinDir & "\System32\icacls.exe"
                    .StartInfo.Arguments = WinDir & "\System32\" & ModifyWizard.Language & "\UIRibbon.dll.mui /setowner ""NT SERVICE\TrustedInstaller"""
                    .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                    .Start()
                    .WaitForExit()
                End With
                ProgressBarMod.PerformStep()
                ProgressBarMod.Refresh()
                ActionLabel.Text = "UIRibbon.dll && UIRibbonRes.dll have been replaced!"
                ActionLabel.Refresh()
            End If
            '' Start userinit.exe
            ModifyModifyExplorerStartInfoWindowEnglish.Show()
            Dim StartExplorerTask As New Process
            With StartExplorerTask
                .StartInfo.FileName = WinDir & "\System32\userinit.exe"
                .Start()
                .WaitForExit()
            End With
            ModifyModifyExplorerStartInfoWindowEnglish.Dispose()
        End If

        ' Installing sounds
        If ModifyWizard.InstallSoundsChecked = True Then
            ActionLabel.Text = "Installing sounds..."
            ActionLabel.Refresh()
            FileIO.FileSystem.CopyDirectory(ResFolder & "\Sounds", WinDir & "\Media\Windows 8 Release Preview")
            FileIO.FileSystem.CopyFile(ResFolder & "\Theme\Themes\aerorp_sounds.theme", WinDir & "\Resources\Themes\aerorp_sounds.theme")
            ProgressBarMod.PerformStep()
            ProgressBarMod.Refresh()
            ActionLabel.Text = "Finished!"
            ActionLabel.Refresh()
        End If

        ' Change 8GadgetPack
        If ModifyWizard.ChangeGadgetsChecked = True Then
            If ModifyWizard.InstGadgets = True And FileIO.FileSystem.FileExists(ProgramFiles & "\Windows Sidebar\8GadgetPack.exe") Then
                ActionLabel.Text = "Uninstalling 8GadgetPack..."
                ActionLabel.Refresh()
                Dim GadgetsSetup As New Process
                With GadgetsSetup
                    .StartInfo.FileName = WinDir & "\System32\msiexec.exe"
                    .StartInfo.Arguments = "/X{B6AF19AD-2D5B-44DC-9272-EC91965123E8} /passive"
                    .Start()
                    .WaitForExit()
                End With
                ProgressBarMod.PerformStep()
                ProgressBarMod.Refresh()
                ActionLabel.Text = "8GadgetPack has been uninstalled!"
                ActionLabel.Refresh()
            ElseIf ModifyWizard.InstGadgets = False Or (ModifyWizard.InstGadgets = True And Not FileIO.FileSystem.FileExists(ProgramFiles & "\Windows Sidebar\8GadgetPack.exe")) Then
                ActionLabel.Text = "Installing 8GadgetPack..."
                ActionLabel.Refresh()
                Dim ChangeGadgetsProcess As New Process
                With ChangeGadgetsProcess
                    .StartInfo.FileName = WinDir & "\System32\msiexec.exe"
                    .StartInfo.Arguments = "/package """ & ResFolder & "\Gadgets\8GadgetPack370Setup.msi"" /passive"
                    .Start()
                    .WaitForExit()
                End With
                ProgressBarMod.PerformStep()
                ProgressBarMod.Refresh()
                ActionLabel.Text = "8GadgetPack has been installed!"
                ActionLabel.Refresh()
            End If
        End If

        'Replacing theme blue/white address bar
        If ModifyWizard.ChangeAddBarStyle = True Then
            If ModifyWizard.AddressBarStyle = "blue" Then
                ActionLabel.Text = "Changing theme (white address bar)..."
                ActionLabel.Refresh()
                FileIO.FileSystem.DeleteFile(WinDir & "\Resources\Themes\aerorp\aerolite.msstyles")
                FileIO.FileSystem.CopyFile(ResFolder & "\Theme\Themes\aerorp\aero_nonblue.msstyles", WinDir & "\Resources\Themes\aerorp\aerolite.msstyles")
            ElseIf ModifyWizard.AddressBarStyle = "white" Then
                ActionLabel.Text = "Changing theme (blue address bar)..."
                ActionLabel.Refresh()
                FileIO.FileSystem.DeleteFile(WinDir & "\Resources\Themes\aerorp\aerolite.msstyles")
                FileIO.FileSystem.CopyFile(ResFolder & "\Theme\Themes\aerorp\aero_blue.msstyles", WinDir & "\Resources\Themes\aerorp\aerolite.msstyles")
            End If
            ProgressBarMod.PerformStep()
            ProgressBarMod.Refresh()
            ActionLabel.Text = "Theme has been changed!"
            ActionLabel.Refresh()
        End If

        ' If theme change needed and restart not needed, change to Windows 8 RP theme now
        If ModifyWizard.Language = "en-GB" Then
            ActionLabel.Text = "Finalising modifications..."
        ElseIf ModifyWizard.Language = "en-US" Then
            ActionLabel.Text = "Finalizing modifications..."
        End If
        ActionLabel.Refresh()
        If ModifyWizard.IsRestartNeeded = False And ModifyWizard.IsThemeChangeNeeded = True Then
            Me.TopMost = True
            If ModifyWizard.InstallSoundsChecked = True Or ModifyWizard.InstSounds = True Then
                Dim ThemeChangeProcess As New Process
                With ThemeChangeProcess
                    .StartInfo.FileName = WinDir & "\Resources\Themes\aerorp_sounds.theme"
                    .Start()
                End With
                Wait(5)
            ElseIf ModifyWizard.InstallSoundsChecked = False And ModifyWizard.InstSounds = False Then
                Dim ThemeChangeProcess As New Process
                With ThemeChangeProcess
                    .StartInfo.FileName = WinDir & "\Resources\Themes\aerorp.theme"
                    .Start()
                End With
                Wait(5)
            End If
            Me.TopMost = False
        ElseIf ModifyWizard.IsRestartNeeded = True And ModifyWizard.IsThemeChangeNeeded = True Then
            '' If theme or UXTheme patcher changed then autostart theme change after reboot
            Dim ThemeAutorunRegKey As RegistryKey = Registry.CurrentUser.OpenSubKey("Software\Microsoft\Windows\CurrentVersion\RunOnce", True)
            If ModifyWizard.InstallSoundsChecked = True Or ModifyWizard.InstSounds = True Then
                ThemeAutorunRegKey.SetValue("aerorp_sounds.theme", WinDir & "\Resources\Themes\aerorp_sounds.theme", RegistryValueKind.String)
                ThemeAutorunRegKey.Close()
            ElseIf ModifyWizard.InstallSoundsChecked = False And ModifyWizard.InstSounds = False Then
                ThemeAutorunRegKey.SetValue("aerorp.theme", WinDir & "\Resources\Themes\aerorp.theme", RegistryValueKind.String)
                ThemeAutorunRegKey.Close()
            End If
        End If

        ' If restart needed and Quero installed, run modify.xht with autostart
        If ModifyWizard.ChangeQueroChecked = True And ModifyWizard.InstQuero = False Then
            If ModifyWizard.IsRestartNeeded = True Then
                Dim AutorunRegKey = Registry.CurrentUser.OpenSubKey("Software\Microsoft\Windows\CurrentVersion\RunOnce", True)
                AutorunRegKey.SetValue("modify.xht", """" & ProgramFiles & "\Internet Explorer\iexplore.exe"" " & """" & ResFolder & "\modify.xht""", RegistryValueKind.String)
                AutorunRegKey.Close()
            End If
        End If
        ProgressBarMod.PerformStep()
        ProgressBarMod.Refresh()

        ' Finishing modifications
        ActionLabel.Text = "Modification(s) has/have been successfully completed!"
        ActionLabel.Refresh()
        '' Change Next button and mark setup as finished
        NextButton.Text = "&Finish"
        NextButton.Enabled = True
        IsFinished = True
    End Sub

    Private Sub NextButton_Click(sender As Object, e As EventArgs) Handles NextButton.Click
        Me.Close()
    End Sub

    Private Sub mod_modify_en_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        If IsFinished = False Then
            MessageBox.Show("The setup cannot be finished right now.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        ElseIf IsFinished = True Then
            If e.CloseReason = CloseReason.UserClosing Then
                If ModifyWizard.IsRestartNeeded = True Then
                    ' Delete OldNewExplorer directory after reboot
                    If ModifyWizard.InstONE = True And ModifyWizard.ChangeONEChecked = True Then
                        Dim DeleteFilesRegKey As RegistryKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce", True)
                        DeleteFilesRegKey.SetValue("Remove OldNewExplorer", """" & WinDir & "\System32\cmd.exe"" /c ""rmdir /s /q """ & ProgramFiles & "\OldNewExplorer""""")
                        DeleteFilesRegKey.Close()
                    End If

                    ' Restart prompt
                    Dim ShutDownProcess As New Process
                    If MessageBox.Show("You need to restart your system to finish the modification(s). Do you want to restart now?", "Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) = System.Windows.Forms.DialogResult.Yes Then
                        With ShutDownProcess
                            .StartInfo.FileName = WinDir & "\System32\shutdown.exe"
                            .StartInfo.Arguments = "/r /t 0"
                            .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                            .Start()
                        End With
                        ModifyWizard.Close()
                    ElseIf System.Windows.Forms.DialogResult.No Then
                        ModifyWizard.Close()
                    End If
                Else
                    If ModifyWizard.ChangeQueroChecked = True And FileIO.FileSystem.DirectoryExists(ProgramFiles & "\Quero Toolbar") Then
                        Dim ProgDir As String = ProgramFiles & "\Windows8.1to8RPConv"
                        Dim ResFolder As String = ProgDir & "\res"
                        Dim RunModifyXhtmlDocument As New Process
                        With RunModifyXhtmlDocument
                            .StartInfo.FileName = ProgramFiles & "\Internet Explorer\iexplore.exe"
                            .StartInfo.Arguments = """" & ResFolder & "\modify.xht"""
                            .Start()
                        End With
                    End If
                    ModifyWizard.Close()
                End If
            End If
        End If
    End Sub
End Class