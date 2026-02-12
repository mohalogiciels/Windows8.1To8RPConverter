Imports Microsoft.Win32
Imports System.ComponentModel
Imports System.IO
Imports System.IO.Compression
Imports System.Security.AccessControl
Imports System.Security.Principal

Public Class ModifyModify
    Private WithEvents BgWorker As BackgroundWorker
    Dim ResFolder As String = ModifyWizard.ProgDir & "\res"
    Dim ProgramFiles As String = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles)
    Dim ProgramFilesX86 As String = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86)
    Dim WinDir As String = Environment.GetFolderPath(Environment.SpecialFolder.Windows)
    Dim IsFinished As Boolean = False
    Dim ProgBarCounter As Integer = 0

    Private Sub ModifyModifyEnglish_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown
        ' Set progress bar to marquee
        With ProgressBarMod
            .Style = ProgressBarStyle.Marquee
            .MarqueeAnimationSpeed = 100
            .Refresh()
        End With

        ' Initialise progress bar
        InitialiseProgressBar()

        ' Initialise BackgroundWorker
        InitialiseBackgroundWorker()

        ' Check if theme change is needed
        If ModifyWizard.IsThemeChangeNeeded = True Or ModifyWizard.ChangeUXMethodChecked = True Then
            '' Change theme now
            Me.TopMost = True
            BgWorker.RunWorkerAsync("ChangeToDefaultTheme")
            Wait(5)
            Me.TopMost = False
            ModifyLauncher()
        Else
            '' Start modifying without theme change
            ModifyLauncher()
        End If
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

        If ModifyWizard.Change7TaskTwChecked = True Then
            ProgBarCounter += 10
        End If

        If ModifyWizard.ChangeAddBarStyle = True Then
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
                '' Change to standard theme
                Using ThemeChangeProcess As New Process
                    With ThemeChangeProcess
                        .StartInfo.FileName = WinDir & "\Resources\Themes\aero.theme"
                        .Start()
                    End With
                End Using
            Case "CloseExplorer"
                '' Close explorer.exe
                Using EndExplorerTask As New Process
                    With EndExplorerTask
                        .StartInfo.FileName = WinDir & "\System32\taskkill.exe"
                        .StartInfo.Arguments = "/f /im explorer.exe"
                        .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                        .Start()
                        .WaitForExit()
                    End With
                End Using
            Case "ChangeONE"
                Dim SourcePathONE As String = ResFolder & "\OldNewExplorer"
                Dim DestinationPathONE As String = ProgramFiles & "\OldNewExplorer"
                If ModifyWizard.ModifyInstallONE = False Then
                    '' Uninstall OldNewExplorer
                    If ActionLabel.InvokeRequired Then
                        ActionLabel.Invoke(Sub()
                                               ActionLabel.Text = "Removing OldNewExplorer..."
                                               ActionLabel.Refresh()
                                           End Sub)
                    Else
                        ActionLabel.Text = "Removing OldNewExplorer..."
                        ActionLabel.Refresh()
                    End If
                    '' Unreg OldNewExplorer DLLs
                    If ModifyWizard.BitnessSystem = "64Bit" Then
                        Using UninstONE As New Process
                            With UninstONE
                                .StartInfo.FileName = WinDir & "\System32\regsvr32.exe"
                                .StartInfo.Arguments = "/u /s """ & DestinationPathONE & "\OldNewExplorer64.dll"""
                                .Start()
                                .WaitForExit()
                            End With
                        End Using
                    End If
                    Using UninstONE As New Process
                        With UninstONE
                            .StartInfo.FileName = WinDir & "\System32\regsvr32.exe"
                            .StartInfo.Arguments = "/u /s """ & DestinationPathONE & "\OldNewExplorer32.dll"""
                            .Start()
                            .WaitForExit()
                        End With
                    End Using
                    If FileIO.FileSystem.FileExists(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) & "\OldNewExplorer Configuration.lnk") Then
                        FileIO.FileSystem.DeleteFile(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) & "\OldNewExplorer Configuration.lnk")
                    End If
                    If ProgressBarMod.InvokeRequired Then
                        ProgressBarMod.Invoke(Sub()
                                                  ProgressBarMod.PerformStep()
                                                  ProgressBarMod.Refresh()
                                              End Sub)
                    Else
                        ProgressBarMod.PerformStep()
                        ProgressBarMod.Refresh()
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
                    '' Install OldNewExplorer
                    If ActionLabel.InvokeRequired Then
                        ActionLabel.Invoke(Sub()
                                               ActionLabel.Text = "Installing OldNewExplorer..."
                                               ActionLabel.Refresh()
                                           End Sub)
                    Else
                        ActionLabel.Text = "Installing OldNewExplorer..."
                        ActionLabel.Refresh()
                    End If
                    FileIO.FileSystem.CopyDirectory(SourcePathONE, DestinationPathONE)
                    If ModifyWizard.BitnessSystem = "64Bit" Then
                        Using ONESetup As New Process
                            With ONESetup
                                .StartInfo.FileName = WinDir & "\System32\regsvr32.exe"
                                .StartInfo.Arguments = "/s """ & DestinationPathONE & "\OldNewExplorer64.dll"""
                                .Start()
                                .WaitForExit()
                            End With
                        End Using
                    End If
                    Using ONESetup As New Process
                        With ONESetup
                            .StartInfo.FileName = WinDir & "\System32\regsvr32.exe"
                            .StartInfo.Arguments = "/s """ & DestinationPathONE & "\OldNewExplorer32.dll"""
                            .Start()
                            .WaitForExit()
                        End With
                    End Using
                    '' Create shortcut on desktop
                    Using ObjectShell As Object = CreateObject("WScript.Shell")
                        Using ObjectLink As Object = ObjectShell.CreateShortcut(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) & "\OldNewExplorer Configuration.lnk")
                            Try
                                With ObjectLink
                                    .TargetPath = ProgramFiles & "\OldNewExplorer\OldNewExplorerCfg.exe"
                                    .WorkingDirectory = ProgramFiles & "\OldNewExplorer"
                                    .WindowStyle = 1
                                    .Save()
                                End With
                            Catch ex As Exception
                                MessageBox.Show("Shortcut to OldNewExplorer configuration panel could not be created on desktop!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                            End Try
                        End Using
                    End Using
                    If ProgressBarMod.InvokeRequired Then
                        ProgressBarMod.Invoke(Sub()
                                                  ProgressBarMod.PerformStep()
                                                  ProgressBarMod.Refresh()
                                              End Sub)
                    Else
                        ProgressBarMod.PerformStep()
                        ProgressBarMod.Refresh()
                    End If
                    If ActionLabel.InvokeRequired Then
                        ActionLabel.Invoke(Sub()
                                               ActionLabel.Text = "OldNewExplorer has been installed!"
                                               ActionLabel.Refresh()
                                           End Sub)
                    Else
                        ActionLabel.Text = "OldNewExplorer has been installed!"
                        ActionLabel.Refresh()
                    End If
                End If
            Case "ChangeUXMethod"
                If ModifyWizard.ModifyChangeToUXMethod = "UXTSB" Then
                    '' Uninstall UltraUX
                    If ActionLabel.InvokeRequired Then
                        ActionLabel.Invoke(Sub()
                                               ActionLabel.Text = "Uninstalling UltraUXThemePatcher..."
                                               ActionLabel.Refresh()
                                           End Sub)
                    Else
                        ActionLabel.Text = "Uninstalling UltraUXThemePatcher..."
                        ActionLabel.Refresh()
                    End If
                    Using UninstUltraUX As New Process
                        With UninstUltraUX
                            If ModifyWizard.BitnessSystem = "64Bit" Then
                                .StartInfo.FileName = ProgramFilesX86 & "\UltraUXThemePatcher\Uninstall.exe"
                            ElseIf ModifyWizard.BitnessSystem = "32Bit" Then
                                .StartInfo.FileName = ProgramFiles & "\UltraUXThemePatcher\Uninstall.exe"
                            End If
                            .Start()
                            .WaitForExit()
                        End With
                    End Using
                    '' WaitForExit for process called Un.exe because Uninstall.exe closes
                    Do
                        Application.DoEvents()
                    Loop Until Process.GetProcessesByName("Un").Count = 0
                    If ProgressBarMod.InvokeRequired Then
                        ProgressBarMod.Invoke(Sub()
                                                  ProgressBarMod.PerformStep()
                                                  ProgressBarMod.Refresh()
                                              End Sub)
                    Else
                        ProgressBarMod.PerformStep()
                        ProgressBarMod.Refresh()
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
                    '' Install UXTSB
                    If ActionLabel.InvokeRequired Then
                        ActionLabel.Invoke(Sub()
                                               ActionLabel.Text = "Installing UXThemeSignatureBypass..."
                                               ActionLabel.Refresh()
                                           End Sub)
                    Else
                        ActionLabel.Text = "Installing UXThemeSignatureBypass..."
                        ActionLabel.Refresh()
                    End If
                    If ModifyWizard.BitnessSystem = "64Bit" Then
                        FileIO.FileSystem.CopyFile(ResFolder & "\UXThemeDLL\UxThemeSignatureBypass64.dll", WinDir & "\System32\UxThemeSignatureBypass.dll")
                        FileIO.FileSystem.CopyFile(ResFolder & "\UXThemeDLL\UxThemeSignatureBypass32.dll", WinDir & "\SysWOW64\UxThemeSignatureBypass.dll")
                    ElseIf ModifyWizard.BitnessSystem = "32Bit" Then
                        FileIO.FileSystem.CopyFile(ResFolder & "\UXThemeDLL\UxThemeSignatureBypass32.dll", WinDir & "\System32\UxThemeSignatureBypass.dll")
                    End If
                    '' UXTSB set registry key to load file with Windows
                    Using RegKey As RegistryKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\Microsoft\Windows NT\CurrentVersion\Windows", True)
                        If RegKey.GetValue("AppInit_DLLs") <> String.Empty Then
                            RegKey.SetValue("AppInit_DLLs", RegKey.GetValue("AppInit_DLLs") & ", " & WinDir & "\System32\UxThemeSignatureBypass.dll")
                        Else
                            RegKey.SetValue("AppInit_DLLs", WinDir & "\System32\UxThemeSignatureBypass.dll")
                        End If
                        If RegKey.GetValue("LoadAppInit_DLLs") <> 1 Then
                            RegKey.SetValue("LoadAppInit_DLLs", 1)
                        End If
                        RegKey.Close()
                    End Using
                    '' UXTSB\Registry\SysWOW64
                    If ModifyWizard.BitnessSystem = "64Bit" Then
                        Using RegKey32 As RegistryKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\Wow6432Node\Microsoft\Windows NT\CurrentVersion\Windows", True)
                            If RegKey32.GetValue("AppInit_DLLs") <> String.Empty Then
                                RegKey32.SetValue("AppInit_DLLs", RegKey32.GetValue("AppInit_DLLs") & ", " & WinDir & "\SysWOW64\UxThemeSignatureBypass.dll")
                            Else
                                RegKey32.SetValue("AppInit_DLLs", WinDir & "\SysWOW64\UxThemeSignatureBypass.dll")
                            End If
                            If RegKey32.GetValue("LoadAppInit_DLLs") <> 1 Then
                                RegKey32.SetValue("LoadAppInit_DLLs", 1)
                            End If
                            RegKey32.Close()
                        End Using
                    End If
                    If ProgressBarMod.InvokeRequired Then
                        ProgressBarMod.Invoke(Sub()
                                                  ProgressBarMod.PerformStep()
                                                  ProgressBarMod.Refresh()
                                              End Sub)
                    Else
                        ProgressBarMod.PerformStep()
                        ProgressBarMod.Refresh()
                    End If
                    If ActionLabel.InvokeRequired Then
                        ActionLabel.Invoke(Sub()
                                               ActionLabel.Text = "UXThemeSignatureBypass has been installed!"
                                               ActionLabel.Refresh()
                                           End Sub)
                    Else
                        ActionLabel.Text = "UXThemeSignatureBypass has been installed!"
                        ActionLabel.Refresh()
                    End If
                ElseIf ModifyWizard.ModifyChangeToUXMethod = "UltraUX" Then
                    If ActionLabel.InvokeRequired Then
                        ActionLabel.Invoke(Sub()
                                               ActionLabel.Text = "Removing UXThemeSignatureBypass..."
                                               ActionLabel.Refresh()
                                           End Sub)
                    Else
                        ActionLabel.Text = "Removing UXThemeSignatureBypass..."
                        ActionLabel.Refresh()
                    End If
                    '' Removing UXTSB after reboot
                    Using RegKeyDelete As RegistryKey = Registry.LocalMachine.OpenSubKey("SYSTEM\CurrentControlSet\Control\Session Manager", True)
                        If RegKeyDelete.GetValue("PendingFileRenameOperations") IsNot Nothing Or RegKeyDelete.GetValue("PendingFileRenameOperations") <> String.Empty Then
                            Dim RegKeyValue As String() = RegKeyDelete.GetValue("PendingFileRenameOperations")
                            Dim RegKeyValueString As String = String.Join(Chr(0), RegKeyValue)
                            Dim DeleteFiles As String() = {RegKeyValueString, "\??\" & WinDir & "\System32\UxThemeSignatureBypass.dll" & Chr(0), "\??\" & WinDir & "\SysWOW64\UxThemeSignatureBypass.dll" & Chr(0)}
                            Dim DeleteFiles32 As String() = {RegKeyValueString, "\??\" & WinDir & "\System32\UxThemeSignatureBypass.dll" & Chr(0)}
                            If ModifyWizard.BitnessSystem = "64Bit" Then
                                RegKeyDelete.SetValue("PendingFileRenameOperations", DeleteFiles, RegistryValueKind.MultiString)
                            ElseIf ModifyWizard.BitnessSystem = "32Bit" Then
                                RegKeyDelete.SetValue("PendingFileRenameOperations", DeleteFiles32, RegistryValueKind.MultiString)
                            End If
                        Else
                            Dim DeleteFiles As String() = {"\??\" & WinDir & "\System32\UxThemeSignatureBypass.dll" & Chr(0), "\??\" & WinDir & "\SysWOW64\UxThemeSignatureBypass.dll" & Chr(0)}
                            Dim DeleteFiles32 As String() = {"\??\" & WinDir & "\System32\UxThemeSignatureBypass.dll" & Chr(0)}
                            If ModifyWizard.BitnessSystem = "64Bit" Then
                                RegKeyDelete.SetValue("PendingFileRenameOperations", DeleteFiles, RegistryValueKind.MultiString)
                            ElseIf ModifyWizard.BitnessSystem = "32Bit" Then
                                RegKeyDelete.SetValue("PendingFileRenameOperations", DeleteFiles32, RegistryValueKind.MultiString)
                            End If
                        End If
                        RegKeyDelete.Close()
                    End Using
                    '' UXTSB -> remove registry key to load file with Windows
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
                    '' UXTSB\Registry\SysWOW64
                    If ModifyWizard.BitnessSystem = "64Bit" Then
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
                    If ProgressBarMod.InvokeRequired Then
                        ProgressBarMod.Invoke(Sub()
                                                  ProgressBarMod.PerformStep()
                                                  ProgressBarMod.Refresh()
                                              End Sub)
                    Else
                        ProgressBarMod.PerformStep()
                        ProgressBarMod.Refresh()
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
                    '' Install UltraUXThemePatcher
                    If ActionLabel.InvokeRequired Then
                        ActionLabel.Invoke(Sub()
                                               ActionLabel.Text = "Installing UltraUXThemePatcher..."
                                               ActionLabel.Refresh()
                                           End Sub)
                    Else
                        ActionLabel.Text = "Installing UltraUXThemePatcher..."
                        ActionLabel.Refresh()
                    End If
                    Using UltraUXSetup As New Process
                        With UltraUXSetup
                            .StartInfo.FileName = ResFolder & "\UXTheme\UltraUXThemePatcher_4.5.0.exe"
                            .Start()
                            .WaitForExit()
                        End With
                    End Using
                    If ProgressBarMod.InvokeRequired() Then
                        ProgressBarMod.Invoke(Sub()
                                                  ProgressBarMod.PerformStep()
                                                  ProgressBarMod.Refresh()
                                              End Sub)
                    Else
                        ProgressBarMod.PerformStep()
                        ProgressBarMod.Refresh()
                    End If
                    If ActionLabel.InvokeRequired Then
                        ActionLabel.Invoke(Sub()
                                               ActionLabel.Text = "UltraUXThemePatcher has been installed!"
                                               ActionLabel.Refresh()
                                           End Sub)
                    Else
                        ActionLabel.Text = "UltraUXThemePatcher has been installed!"
                        ActionLabel.Refresh()
                    End If
                End If
            Case "ChangeQuero"
                If ModifyWizard.ModifyInstallQuero = True Then
                    If ActionLabel.InvokeRequired Then
                        ActionLabel.Invoke(Sub()
                                               ActionLabel.Text = "Installing Quero Toolbar..."
                                               ActionLabel.Refresh()
                                           End Sub)
                    Else
                        ActionLabel.Text = "Installing Quero Toolbar..."
                        ActionLabel.Refresh()
                    End If
                    '' Change install directory to ProgramFiles folder
                    Dim QueroSetupInfFilePath As String = ResFolder & "\Quero\settings.inf"
                    Dim QueroSetupInfFile As String() = File.ReadAllLines(QueroSetupInfFilePath)
                    QueroSetupInfFile(2) = "Dir=" & ProgramFiles & "\Quero Toolbar"
                    File.WriteAllLines(QueroSetupInfFilePath, QueroSetupInfFile)
                    Erase QueroSetupInfFile
                    '' Start Quero Toolbar setup
                    Using QueroSetup As New Process
                        With QueroSetup
                            If ModifyWizard.BitnessSystem = "64Bit" Then
                                .StartInfo.FileName = ResFolder & "\Quero\QueroToolbarInstaller_x64.exe"
                            ElseIf ModifyWizard.BitnessSystem = "32Bit" Then
                                .StartInfo.FileName = ResFolder & "\Quero\QueroToolbarInstaller_x86.exe"
                            End If
                            .StartInfo.Arguments = "/VERYSILENT /LOADINF=""" & QueroSetupInfFilePath & """"
                            .Start()
                            .WaitForExit()
                        End With
                    End Using
                    File.WriteAllBytes(ModifyWizard.ProgDir & "\modify.xht", My.Resources.ResourceManager.GetObject("ModifyQuero_" & ModifyWizard.LanguageForLocalisedStrings))
                    If ProgressBarMod.InvokeRequired Then
                        ProgressBarMod.Invoke(Sub()
                                                  ProgressBarMod.PerformStep()
                                                  ProgressBarMod.Refresh()
                                              End Sub)
                    Else
                        ProgressBarMod.PerformStep()
                        ProgressBarMod.Refresh()
                    End If
                    If ActionLabel.InvokeRequired Then
                        ActionLabel.Invoke(Sub()
                                               ActionLabel.Text = "Quero Toolbar has been installed!"
                                               ActionLabel.Refresh()
                                           End Sub)
                    Else
                        ActionLabel.Text = "Quero Toolbar has been installed!"
                        ActionLabel.Refresh()
                    End If
                Else
                    If ActionLabel.InvokeRequired Then
                        ActionLabel.Invoke(Sub()
                                               ActionLabel.Text = "Uninstalling Quero Toolbar..."
                                               ActionLabel.Refresh()
                                           End Sub)
                    Else
                        ActionLabel.Text = "Uninstalling Quero Toolbar..."
                        ActionLabel.Refresh()
                    End If
                    '' Restore address bar of Internet Explorer
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
                    '' Run uninstaller
                    Using QueroUninst As New Process
                        With QueroUninst
                            .StartInfo.FileName = ProgramFiles & "\Quero Toolbar\unins000.exe"
                            .StartInfo.Arguments = "/VERYSILENT /SUPPRESSMSGBOXES"
                            .Start()
                            .WaitForExit()
                        End With
                    End Using
                    If ProgressBarMod.InvokeRequired Then
                        ProgressBarMod.Invoke(Sub()
                                                  ProgressBarMod.PerformStep()
                                                  ProgressBarMod.Refresh()
                                              End Sub)
                    Else
                        ProgressBarMod.PerformStep()
                        ProgressBarMod.Refresh()
                    End If
                    If ActionLabel.InvokeRequired Then
                        ActionLabel.Invoke(Sub()
                                               ActionLabel.Text = "Quero Toolbar has been uninstalled!"
                                               ActionLabel.Refresh()
                                           End Sub)
                    Else
                        ActionLabel.Text = "Quero Toolbar has been uninstalled!"
                        ActionLabel.Refresh()
                    End If
                End If
            Case "ChangeSysFiles"
                Dim EveryoneSidAsUserName As String = New SecurityIdentifier(WellKnownSidType.WorldSid, Nothing).Translate(GetType(NTAccount)).Value
                Dim EveryoneAuditRule As FileSystemAuditRule = New FileSystemAuditRule(EveryoneSidAsUserName, FileSystemRights.CreateFiles + FileSystemRights.CreateDirectories + FileSystemRights.WriteAttributes + FileSystemRights.WriteExtendedAttributes + FileSystemRights.Delete + FileSystemRights.ChangePermissions + FileSystemRights.TakeOwnership, AuditFlags.Success + AuditFlags.Failure)
                Dim UiRibbonAuditRules As FileSecurity = Nothing
                If ModifyWizard.ModifyInstallSysFiles = False Then
                    '' Restore system files
                    If ActionLabel.InvokeRequired Then
                        ActionLabel.Invoke(Sub()
                                               ActionLabel.Text = "Restoring UIRibbon.dll && UIRibbonRes.dll..."
                                               ActionLabel.Refresh()
                                           End Sub)
                    Else
                        ActionLabel.Text = "Restoring UIRibbon.dll && UIRibbonRes.dll..."
                        ActionLabel.Refresh()
                    End If
                    '' Check if backup files exist
                    If ModifyWizard.BitnessSystem = "64Bit" Then
                        If Not FileIO.FileSystem.FileExists(WinDir & "\System32\UIRibbon.dll.bak") Or Not FileIO.FileSystem.FileExists(WinDir & "\System32\UIRibbonRes.dll.bak") Or Not FileIO.FileSystem.FileExists(WinDir & "\System32\" & ModifyWizard.ProgLanguage & "\UIRibbon.dll.mui.bak") Or Not FileIO.FileSystem.FileExists(WinDir & "\SysWOW64\UIRibbon.dll.bak") Or Not FileIO.FileSystem.FileExists(WinDir & "\SysWOW64\UIRibbonRes.dll.bak") Or Not FileIO.FileSystem.FileExists(WinDir & "\SysWOW64\" & ModifyWizard.ProgLanguage & "\UIRibbon.dll.mui.bak") Then
                            MessageBox.Show("One or more backup files could not be found on this system! Skipping restore of system files...", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            If ProgressBarMod.InvokeRequired Then
                                ProgressBarMod.Invoke(Sub()
                                                          ProgressBarMod.PerformStep()
                                                          ProgressBarMod.Refresh()
                                                      End Sub)
                            Else
                                ProgressBarMod.PerformStep()
                                ProgressBarMod.Refresh()
                            End If
                            Exit Select
                        End If
                    ElseIf ModifyWizard.BitnessSystem = "32Bit" Then
                        If Not FileIO.FileSystem.FileExists(WinDir & "\System32\UIRibbon.dll.bak") Or Not FileIO.FileSystem.FileExists(WinDir & "\System32\UIRibbonRes.dll.bak") Or Not FileIO.FileSystem.FileExists(WinDir & "\System32\" & ModifyWizard.ProgLanguage & "\UIRibbon.dll.mui.bak") Then
                            MessageBox.Show("One or more backup files could not be found on this system! Skipping restore of system files...", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            If ProgressBarMod.InvokeRequired Then
                                ProgressBarMod.Invoke(Sub()
                                                          ProgressBarMod.PerformStep()
                                                          ProgressBarMod.Refresh()
                                                      End Sub)
                            Else
                                ProgressBarMod.PerformStep()
                                ProgressBarMod.Refresh()
                            End If
                            Exit Select
                        End If
                    End If
                    '' Restore files now
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
                        .StartInfo.Arguments = "/f " & WinDir & "\System32\" & ModifyWizard.ProgLanguage & "\UIRibbon.dll.mui /a"
                        .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                        .Start()
                        .WaitForExit()
                    End With
                    With SysFilesSetup
                        .StartInfo.FileName = WinDir & "\System32\icacls.exe"
                        .StartInfo.Arguments = WinDir & "\System32\" & ModifyWizard.ProgLanguage & "\UIRibbon.dll.mui /grant *S-1-5-32-545:F *S-1-5-32-544:F"
                        .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                        .Start()
                        .WaitForExit()
                    End With
                    With SysFilesSetup
                        .StartInfo.FileName = WinDir & "\System32\takeown.exe"
                        .StartInfo.Arguments = "/f " & WinDir & "\System32\" & ModifyWizard.ProgLanguage & "\UIRibbon.dll.mui.bak /a"
                        .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                        .Start()
                        .WaitForExit()
                    End With
                    With SysFilesSetup
                        .StartInfo.FileName = WinDir & "\System32\icacls.exe"
                        .StartInfo.Arguments = WinDir & "\System32\" & ModifyWizard.ProgLanguage & "\UIRibbon.dll.mui.bak /grant *S-1-5-32-545:F *S-1-5-32-544:F"
                        .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                        .Start()
                        .WaitForExit()
                    End With
                    FileIO.FileSystem.DeleteFile(WinDir & "\System32\" & ModifyWizard.ProgLanguage & "\UIRibbon.dll.mui")
                    FileIO.FileSystem.RenameFile(WinDir & "\System32\" & ModifyWizard.ProgLanguage & "\UIRibbon.dll.mui.bak", "UIRibbon.dll.mui")
                    '' Original files -> change to original FileSystemSecurity
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
                        .StartInfo.Arguments = WinDir & "\System32\" & ModifyWizard.ProgLanguage & "\UIRibbon.dll.mui /setowner ""NT SERVICE\TrustedInstaller"""
                        .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                        .Start()
                        .WaitForExit()
                    End With
                    With SysFilesSetup
                        .StartInfo.FileName = WinDir & "\System32\icacls.exe"
                        .StartInfo.Arguments = WinDir & "\System32\" & ModifyWizard.ProgLanguage & "\UIRibbon.dll.mui /reset"
                        .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                        .Start()
                        .WaitForExit()
                    End With
                    With SysFilesSetup
                        .StartInfo.FileName = WinDir & "\System32\icacls.exe"
                        .StartInfo.Arguments = WinDir & "\System32\" & ModifyWizard.ProgLanguage & "\UIRibbon.dll.mui /grant *S-1-5-18:RX *S-1-5-32-544:RX *S-1-5-32-545:RX *S-1-15-2-1:RX ""NT SERVICE\TrustedInstaller"":F /inheritance:r"
                        .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                        .Start()
                        .WaitForExit()
                    End With
                    UiRibbonAuditRules = File.GetAccessControl(WinDir & "\System32\" & ModifyWizard.ProgLanguage & "\UIRibbon.dll.mui", AccessControlSections.Audit)
                    UiRibbonAuditRules.SetAuditRule(EveryoneAuditRule)
                    File.SetAccessControl(WinDir & "\System32\" & ModifyWizard.ProgLanguage & "\UIRibbon.dll.mui", UiRibbonAuditRules)
                    UiRibbonAuditRules = Nothing
                    '' If System is 64-bit, repeat for files in SysWOW64
                    If ModifyWizard.BitnessSystem = "64Bit" Then
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
                            .StartInfo.Arguments = "/f " & WinDir & "\SysWOW64\" & ModifyWizard.ProgLanguage & "\UIRibbon.dll.mui /a"
                            .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                            .Start()
                            .WaitForExit()
                        End With
                        With SysFilesSetup
                            .StartInfo.FileName = WinDir & "\System32\icacls.exe"
                            .StartInfo.Arguments = WinDir & "\SysWOW64\" & ModifyWizard.ProgLanguage & "\UIRibbon.dll.mui /grant *S-1-5-32-545:F *S-1-5-32-544:F"
                            .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                            .Start()
                            .WaitForExit()
                        End With
                        With SysFilesSetup
                            .StartInfo.FileName = WinDir & "\System32\takeown.exe"
                            .StartInfo.Arguments = "/f " & WinDir & "\SysWOW64\" & ModifyWizard.ProgLanguage & "\UIRibbon.dll.mui.bak /a"
                            .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                            .Start()
                            .WaitForExit()
                        End With
                        With SysFilesSetup
                            .StartInfo.FileName = WinDir & "\System32\icacls.exe"
                            .StartInfo.Arguments = WinDir & "\SysWOW64\" & ModifyWizard.ProgLanguage & "\UIRibbon.dll.mui.bak /grant *S-1-5-32-545:F *S-1-5-32-544:F"
                            .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                            .Start()
                            .WaitForExit()
                        End With
                        FileIO.FileSystem.DeleteFile(WinDir & "\SysWOW64\" & ModifyWizard.ProgLanguage & "\UIRibbon.dll.mui")
                        FileIO.FileSystem.RenameFile(WinDir & "\SysWOW64\" & ModifyWizard.ProgLanguage & "\UIRibbon.dll.mui.bak", "UIRibbon.dll.mui")
                        '' Original files in SysWOW64 -> change to original FileSystemSecurity
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
                            .StartInfo.Arguments = WinDir & "\SysWOW64\" & ModifyWizard.ProgLanguage & "\UIRibbon.dll.mui /setowner ""NT SERVICE\TrustedInstaller"""
                            .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                            .Start()
                            .WaitForExit()
                        End With
                        With SysFilesSetup
                            .StartInfo.FileName = WinDir & "\System32\icacls.exe"
                            .StartInfo.Arguments = WinDir & "\SysWOW64\" & ModifyWizard.ProgLanguage & "\UIRibbon.dll.mui /reset"
                            .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                            .Start()
                            .WaitForExit()
                        End With
                        With SysFilesSetup
                            .StartInfo.FileName = WinDir & "\System32\icacls.exe"
                            .StartInfo.Arguments = WinDir & "\SysWOW64\" & ModifyWizard.ProgLanguage & "\UIRibbon.dll.mui /grant *S-1-5-18:RX *S-1-5-32-544:RX *S-1-5-32-545:RX *S-1-15-2-1:RX ""NT SERVICE\TrustedInstaller"":F /inheritance:r"
                            .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                            .Start()
                            .WaitForExit()
                        End With
                        UiRibbonAuditRules = File.GetAccessControl(WinDir & "\SysWOW64\" & ModifyWizard.ProgLanguage & "\UIRibbon.dll.mui", AccessControlSections.Audit)
                        UiRibbonAuditRules.SetAuditRule(EveryoneAuditRule)
                        File.SetAccessControl(WinDir & "\SysWOW64\" & ModifyWizard.ProgLanguage & "\UIRibbon.dll.mui", UiRibbonAuditRules)
                        UiRibbonAuditRules = Nothing
                    End If
                    SysFilesSetup.Dispose()
                    If ProgressBarMod.InvokeRequired Then
                        ProgressBarMod.Invoke(Sub()
                                                  ProgressBarMod.PerformStep()
                                                  ProgressBarMod.Refresh()
                                              End Sub)
                    Else
                        ProgressBarMod.PerformStep()
                        ProgressBarMod.Refresh()
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
                Else
                    '' Replace system files
                    If ActionLabel.InvokeRequired Then
                        ActionLabel.Invoke(Sub()
                                               ActionLabel.Text = "Replacing UIRibbon.dll && UIRibbonRes.dll..."
                                               ActionLabel.Refresh()
                                           End Sub)
                    Else
                        ActionLabel.Text = "Replacing UIRibbon.dll && UIRibbonRes.dll..."
                        ActionLabel.Refresh()
                    End If
                    '' Replace the files UIRibbon.dll, UIRibbonRes.dll and <UILanguage>\UIRibbon.dll.mui
                    Dim SourcePathSysFiles As String = ResFolder & "\UIRibbon"
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
                    FileIO.FileSystem.RenameFile(WinDir & "\System32\UIRibbon.dll", "UIRibbon.dll.bak")
                    FileIO.FileSystem.RenameFile(WinDir & "\System32\UIRibbonRes.dll", "UIRibbonRes.dll.bak")
                    With SysFilesSetup
                        .StartInfo.FileName = WinDir & "\System32\takeown.exe"
                        .StartInfo.Arguments = "/f " & WinDir & "\System32\" & ModifyWizard.ProgLanguage & "\UIRibbon.dll.mui /a"
                        .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                        .Start()
                        .WaitForExit()
                    End With
                    With SysFilesSetup
                        .StartInfo.FileName = WinDir & "\System32\icacls.exe"
                        .StartInfo.Arguments = WinDir & "\System32\" & ModifyWizard.ProgLanguage & "\UIRibbon.dll.mui /grant *S-1-5-32-545:F *S-1-5-32-544:F"
                        .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                        .Start()
                        .WaitForExit()
                    End With
                    FileIO.FileSystem.RenameFile(WinDir & "\System32\" & ModifyWizard.ProgLanguage & "\UIRibbon.dll.mui", "UIRibbon.dll.mui.bak")
                    If ModifyWizard.BitnessSystem = "64Bit" Then
                        FileIO.FileSystem.CopyFile(SourcePathSysFiles & "\system32\UIRibbon.dll", WinDir & "\System32\UIRibbon.dll")
                        FileIO.FileSystem.CopyFile(SourcePathSysFiles & "\system32\UIRibbonRes.dll", WinDir & "\System32\UIRibbonRes.dll")
                        FileIO.FileSystem.CopyFile(SourcePathSysFiles & "\system32\" & ModifyWizard.ProgLanguage & "\UIRibbon.dll.mui", WinDir & "\System32\" & ModifyWizard.ProgLanguage & "\UIRibbon.dll.mui")
                    ElseIf ModifyWizard.BitnessSystem = "32Bit" Then
                        FileIO.FileSystem.CopyFile(SourcePathSysFiles & "\syswow64\UIRibbon.dll", WinDir & "\System32\UIRibbon.dll")
                        FileIO.FileSystem.CopyFile(SourcePathSysFiles & "\syswow64\UIRibbonRes.dll", WinDir & "\System32\UIRibbonRes.dll")
                        FileIO.FileSystem.CopyFile(SourcePathSysFiles & "\syswow64\" & ModifyWizard.ProgLanguage & "\UIRibbon.dll.mui", WinDir & "\System32\" & ModifyWizard.ProgLanguage & "\UIRibbon.dll.mui")
                    End If
                    '' New files -> change to original FileSystemSecurity
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
                        .StartInfo.Arguments = WinDir & "\System32\" & ModifyWizard.ProgLanguage & "\UIRibbon.dll.mui /setowner ""NT SERVICE\TrustedInstaller"""
                        .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                        .Start()
                        .WaitForExit()
                    End With
                    With SysFilesSetup
                        .StartInfo.FileName = WinDir & "\System32\icacls.exe"
                        .StartInfo.Arguments = WinDir & "\System32\" & ModifyWizard.ProgLanguage & "\UIRibbon.dll.mui /reset"
                        .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                        .Start()
                        .WaitForExit()
                    End With
                    With SysFilesSetup
                        .StartInfo.FileName = WinDir & "\System32\icacls.exe"
                        .StartInfo.Arguments = WinDir & "\System32\" & ModifyWizard.ProgLanguage & "\UIRibbon.dll.mui /grant *S-1-5-18:RX *S-1-5-32-544:RX *S-1-5-32-545:RX *S-1-15-2-1:RX ""NT SERVICE\TrustedInstaller"":F /inheritance:r"
                        .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                        .Start()
                        .WaitForExit()
                    End With
                    UiRibbonAuditRules = File.GetAccessControl(WinDir & "\System32\" & ModifyWizard.ProgLanguage & "\UIRibbon.dll.mui", AccessControlSections.Audit)
                    UiRibbonAuditRules.SetAuditRule(EveryoneAuditRule)
                    File.SetAccessControl(WinDir & "\System32\" & ModifyWizard.ProgLanguage & "\UIRibbon.dll.mui", UiRibbonAuditRules)
                    UiRibbonAuditRules = Nothing
                    '' If System is 64-bit, repeat for files in SysWOW64
                    If ModifyWizard.BitnessSystem = "64Bit" Then
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
                            .StartInfo.Arguments = "/f " & WinDir & "\SysWOW64\" & ModifyWizard.ProgLanguage & "\UIRibbon.dll.mui /a"
                            .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                            .Start()
                            .WaitForExit()
                        End With
                        With SysFilesSetup
                            .StartInfo.FileName = WinDir & "\System32\icacls.exe"
                            .StartInfo.Arguments = WinDir & "\SysWOW64\" & ModifyWizard.ProgLanguage & "\UIRibbon.dll.mui /grant *S-1-5-32-545:F *S-1-5-32-544:F"
                            .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                            .Start()
                            .WaitForExit()
                        End With
                        FileIO.FileSystem.RenameFile(WinDir & "\SysWOW64\" & ModifyWizard.ProgLanguage & "\UIRibbon.dll.mui", "UIRibbon.dll.mui.bak")
                        FileIO.FileSystem.CopyFile(SourcePathSysFiles & "\syswow64\UIRibbon.dll", WinDir & "\SysWOW64\UIRibbon.dll")
                        FileIO.FileSystem.CopyFile(SourcePathSysFiles & "\syswow64\UIRibbonRes.dll", WinDir & "\SysWOW64\UIRibbonRes.dll")
                        FileIO.FileSystem.CopyFile(SourcePathSysFiles & "\syswow64\" & ModifyWizard.ProgLanguage & "\UIRibbon.dll.mui", WinDir & "\SysWOW64\" & ModifyWizard.ProgLanguage & "\UIRibbon.dll.mui")
                        '' New files in SysWOW64 -> change to original FileSystemSecurity
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
                            .StartInfo.Arguments = WinDir & "\SysWOW64\" & ModifyWizard.ProgLanguage & "\UIRibbon.dll.mui /setowner ""NT SERVICE\TrustedInstaller"""
                            .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                            .Start()
                            .WaitForExit()
                        End With
                        With SysFilesSetup
                            .StartInfo.FileName = WinDir & "\System32\icacls.exe"
                            .StartInfo.Arguments = WinDir & "\SysWOW64\" & ModifyWizard.ProgLanguage & "\UIRibbon.dll.mui /reset"
                            .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                            .Start()
                            .WaitForExit()
                        End With
                        With SysFilesSetup
                            .StartInfo.FileName = WinDir & "\System32\icacls.exe"
                            .StartInfo.Arguments = WinDir & "\SysWOW64\" & ModifyWizard.ProgLanguage & "\UIRibbon.dll.mui /grant *S-1-5-18:RX *S-1-5-32-544:RX *S-1-5-32-545:RX *S-1-15-2-1:RX ""NT SERVICE\TrustedInstaller"":F /inheritance:r"
                            .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                            .Start()
                            .WaitForExit()
                        End With
                        UiRibbonAuditRules = File.GetAccessControl(WinDir & "\SysWOW64\" & ModifyWizard.ProgLanguage & "\UIRibbon.dll.mui", AccessControlSections.Audit)
                        UiRibbonAuditRules.SetAuditRule(EveryoneAuditRule)
                        File.SetAccessControl(WinDir & "\SysWOW64\" & ModifyWizard.ProgLanguage & "\UIRibbon.dll.mui", UiRibbonAuditRules)
                        UiRibbonAuditRules = Nothing
                    End If
                    SysFilesSetup.Dispose()
                    If ProgressBarMod.InvokeRequired Then
                        ProgressBarMod.Invoke(Sub()
                                                  ProgressBarMod.PerformStep()
                                                  ProgressBarMod.Refresh()
                                              End Sub)
                    Else
                        ProgressBarMod.PerformStep()
                        ProgressBarMod.Refresh()
                    End If
                    If ActionLabel.InvokeRequired Then
                        ActionLabel.Invoke(Sub()
                                               ActionLabel.Text = "UIRibbon.dll && UIRibbonRes.dll have been replaced!"
                                               ActionLabel.Refresh()
                                           End Sub)
                    Else
                        ActionLabel.Text = "UIRibbon.dll && UIRibbonRes.dll have been replaced!"
                        ActionLabel.Refresh()
                    End If
                End If
            Case "ChangeGadgets"
                If ModifyWizard.ModifyInstallGadgets = False Then
                    If ActionLabel.InvokeRequired Then
                        ActionLabel.Invoke(Sub()
                                               ActionLabel.Text = "Uninstalling 8GadgetPack..."
                                               ActionLabel.Refresh()
                                           End Sub)
                    Else
                        ActionLabel.Text = "Uninstalling 8GadgetPack..."
                        ActionLabel.Refresh()
                    End If
                    Using GadgetsSetup As New Process
                        With GadgetsSetup
                            .StartInfo.FileName = WinDir & "\System32\msiexec.exe"
                            .StartInfo.Arguments = "/X{B6AF19AD-2D5B-44DC-9272-EC91965123E8} /quiet"
                            .Start()
                            .WaitForExit()
                        End With
                    End Using
                    If ProgressBarMod.InvokeRequired Then
                        ProgressBarMod.Invoke(Sub()
                                                  ProgressBarMod.PerformStep()
                                                  ProgressBarMod.Refresh()
                                              End Sub)
                    Else
                        ProgressBarMod.PerformStep()
                        ProgressBarMod.Refresh()
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
                    If ActionLabel.InvokeRequired Then
                        ActionLabel.Invoke(Sub()
                                               ActionLabel.Text = "Installing 8GadgetPack..."
                                               ActionLabel.Refresh()
                                           End Sub)
                    Else
                        ActionLabel.Text = "Installing 8GadgetPack..."
                        ActionLabel.Refresh()
                    End If
                    Using ChangeGadgetsProcess As New Process
                        With ChangeGadgetsProcess
                            .StartInfo.FileName = WinDir & "\System32\msiexec.exe"
                            .StartInfo.Arguments = "/package """ & ResFolder & "\Gadgets\8GadgetPack370Setup.msi"" /quiet"
                            .Start()
                            .WaitForExit()
                        End With
                    End Using
                    If ProgressBarMod.InvokeRequired Then
                        ProgressBarMod.Invoke(Sub()
                                                  ProgressBarMod.PerformStep()
                                                  ProgressBarMod.Refresh()
                                              End Sub)
                    Else
                        ProgressBarMod.PerformStep()
                        ProgressBarMod.Refresh()
                    End If
                    If ActionLabel.InvokeRequired Then
                        ActionLabel.Invoke(Sub()
                                               ActionLabel.Text = "8GadgetPack has been installed!"
                                               ActionLabel.Refresh()
                                           End Sub)
                    Else
                        ActionLabel.Text = "8GadgetPack has been installed!"
                        ActionLabel.Refresh()
                    End If
                End If
            Case "Change7TaskTw"
                If ModifyWizard.ModifyInstall7TaskTw = True Then
                    If ActionLabel.InvokeRequired Then
                        ActionLabel.Invoke(Sub()
                                               ActionLabel.Text = "Installing 7+ Taskbar Tweaker..."
                                               ActionLabel.Refresh()
                                           End Sub)
                    Else
                        ActionLabel.Text = "Installing 7+ Taskbar Tweaker..."
                        ActionLabel.Refresh()
                    End If
                    Using TaskTwSetup As New Process
                        With TaskTwSetup
                            .StartInfo.FileName = ResFolder & "\7TaskTw\7tt_setup.exe"
                            .StartInfo.Arguments = "/S"
                            .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                            .Start()
                            .WaitForExit()
                        End With
                    End Using
                    '' Apply settings with registry file 7TaskTw.reg
                    Using ApplyRegFile As New Process
                        With ApplyRegFile
                            .StartInfo.FileName = WinDir & "\System32\reg.exe"
                            .StartInfo.Arguments = "import """ & ResFolder & "\7TaskTw\settings.reg"""
                            .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                            .Start()
                            .WaitForExit()
                        End With
                    End Using
                    If ProgressBarMod.InvokeRequired Then
                        ProgressBarMod.Invoke(Sub()
                                                  ProgressBarMod.PerformStep()
                                                  ProgressBarMod.Refresh()
                                              End Sub)
                    Else
                        ProgressBarMod.PerformStep()
                        ProgressBarMod.Refresh()
                    End If
                    If ActionLabel.InvokeRequired Then
                        ActionLabel.Invoke(Sub()
                                               ActionLabel.Text = "7+ Taskbar Tweaker has been installed!"
                                               ActionLabel.Refresh()
                                           End Sub)
                    Else
                        ActionLabel.Text = "7+ Taskbar Tweaker has been installed!"
                        ActionLabel.Refresh()
                    End If
                Else
                    If ActionLabel.InvokeRequired Then
                        ActionLabel.Invoke(Sub()
                                               ActionLabel.Text = "Uninstalling 7+ Taskbar Tweaker..."
                                               ActionLabel.Refresh()
                                           End Sub)
                    Else
                        ActionLabel.Text = "Uninstalling 7+ Taskbar Tweaker..."
                        ActionLabel.Refresh()
                    End If
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
                End If
            Case "StartExplorer"
                '' Start explorer with userinit.exe after changing system files
                Using StartExplorerTask As New Process
                    With StartExplorerTask
                        .StartInfo.FileName = WinDir & "\System32\userinit.exe"
                        .Start()
                        .WaitForExit()
                    End With
                End Using
        End Select

        ' CancelAsync of BackgroundWorker after task finishes
        BgWorker.CancelAsync()
    End Sub

    Private Sub ModifyLauncher()
        ' Initialise progress bar
        With ProgressBarMod
            .Style = ProgressBarStyle.Continuous
            .Value = 0
            .Minimum = 0
            .Maximum = ProgBarCounter
            .Refresh()
        End With

        ' Close explorer.exe
        BgWorker.RunWorkerAsync("CloseExplorer")
        WaitUntilTaskFinishes()

        ' Start modifying now
        Try
            RunModify()
        Catch ex As Exception
            MessageBox.Show("An error occurred: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub RunModify()
        ' Change OldNewExplorer
        If ModifyWizard.ChangeONEChecked = True Then
            BgWorker.RunWorkerAsync("ChangeONE")
            WaitUntilTaskFinishes()
        End If

        ' Change UXTheme patcher
        If ModifyWizard.ChangeUXMethodChecked = True Then
            BgWorker.RunWorkerAsync("ChangeUXMethod")
            WaitUntilTaskFinishes()
        End If

        ' Change Quero Toolbar
        If ModifyWizard.ChangeQueroChecked = True Then
            BgWorker.RunWorkerAsync("ChangeQuero")
            WaitUntilTaskFinishes()
        End If

        ' Change UIRibbon.dll & UIRibbonRes.dll
        If ModifyWizard.ChangeSysFilesChecked = True Then
            BgWorker.RunWorkerAsync("ChangeSysFiles")
            WaitUntilTaskFinishes()
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
            BgWorker.RunWorkerAsync("ChangeGadgets")
            WaitUntilTaskFinishes()
        End If

        ' Change 7+ Taskbar Tweaker
        If ModifyWizard.Change7TaskTwChecked = True Then
            BgWorker.RunWorkerAsync("Change7TaskTw")
            WaitUntilTaskFinishes()
        End If

        'Replacing theme blue/white address bar
        If ModifyWizard.ChangeAddBarStyle = True Then
            If ModifyWizard.ChangeToSelectedStyle = "WhiteAddressBar" Then
                ActionLabel.Text = "Changing theme (white address bar)..."
                ActionLabel.Refresh()
                FileIO.FileSystem.DeleteFile(WinDir & "\Resources\Themes\aerorp\aerolite.msstyles")
                FileIO.FileSystem.CopyFile(ResFolder & "\Theme\Themes\aerorp\aero_nonblue.msstyles", WinDir & "\Resources\Themes\aerorp\aerolite.msstyles")
            ElseIf ModifyWizard.ChangeToSelectedStyle = "BlueAddressBar" Then
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
        If ModifyWizard.ProgLanguage = "en-GB" Then
            ActionLabel.Text = "Finalising modifications..."
        ElseIf ModifyWizard.ProgLanguage = "en-US" Then
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
        End If

        ' Set autostart registry keys
        If ModifyWizard.IsRestartNeeded = True And (ModifyWizard.IsThemeChangeNeeded = True Or ModifyWizard.ChangeUXMethodChecked = True) Then
            If Registry.CurrentUser.OpenSubKey("Software\Microsoft\Windows\CurrentVersion\RunOnce", False) Is Nothing Then
                Using CurrentVersionRegKey As RegistryKey = Registry.CurrentUser.OpenSubKey("Software\Microsoft\Windows\CurrentVersion", True)
                    CurrentVersionRegKey.CreateSubKey("RunOnce")
                    CurrentVersionRegKey.Close()
                End Using
            End If
            Using AutorunRegKey As RegistryKey = Registry.CurrentUser.OpenSubKey("Software\Microsoft\Windows\CurrentVersion\RunOnce", True)
                AutorunRegKey.SetValue("Run firstrun.exe", """" & ModifyWizard.ProgDir & "\firstrun.exe"" /modify", RegistryValueKind.String)
                If ModifyWizard.IsThemeChangeNeeded = True Or ModifyWizard.ChangeUXMethodChecked = True Then
                    '' If theme or UXTheme patcher changed then autostart theme change after reboot
                    If ModifyWizard.InstSounds = True Or ModifyWizard.InstallSoundsChecked = True Then
                        AutorunRegKey.SetValue("Run firstrun.exe", AutorunRegKey.GetValue("Run firstrun.exe") & " /applytheme /sounds")
                    Else
                        AutorunRegKey.SetValue("Run firstrun.exe", AutorunRegKey.GetValue("Run firstrun.exe") & " /applytheme")
                    End If
                End If
                '' If restart needed and Quero installed, run modify.xht with autostart
                If ModifyWizard.ChangeQueroChecked = True And ModifyWizard.ModifyInstallQuero = True Then
                    AutorunRegKey.SetValue("Run firstrun.exe", AutorunRegKey.GetValue("Run firstrun.exe") & " /quero")
                End If
                AutorunRegKey.Close()
            End Using
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

    Private Sub ModifyModify_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        If IsFinished = False Then
            MessageBox.Show("The setup cannot be finished right now.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        ElseIf IsFinished = True Then
            If e.CloseReason = CloseReason.UserClosing Then
                If ModifyWizard.IsRestartNeeded = True Then
                    '' Delete OldNewExplorer directory after reboot when uninstalled
                    If ModifyWizard.ChangeONEChecked = True And ModifyWizard.ModifyInstallONE = False Then
                        Using DeleteFilesRegKey As RegistryKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce", True)
                            DeleteFilesRegKey.SetValue("Remove OldNewExplorer", """" & WinDir & "\System32\cmd.exe"" /c ""rmdir /s /q """ & ProgramFiles & "\OldNewExplorer""""")
                            DeleteFilesRegKey.Close()
                        End Using
                    End If
                    '' Restart prompt
                    If MessageBox.Show("You need to restart your system to finish the modification(s). Do you want to restart now?", "Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) = System.Windows.Forms.DialogResult.Yes Then
                        Using ShutDownProcess As New Process
                            With ShutDownProcess
                                .StartInfo.FileName = WinDir & "\System32\shutdown.exe"
                                .StartInfo.Arguments = "/r /t 3"
                                .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                                .Start()
                            End With
                        End Using
                        ModifyWizard.Close()
                    ElseIf System.Windows.Forms.DialogResult.No Then
                        ModifyWizard.Close()
                    End If
                Else
                    '' Start explorer.exe
                    BgWorker.RunWorkerAsync("StartExplorer")
                    WaitUntilTaskFinishes()
                    '' Open modify.xht in Internet Explorer
                    If ModifyWizard.ChangeQueroChecked = True And ModifyWizard.ModifyInstallQuero = True Then
                        Dim ResFolder As String = ModifyWizard.ProgDir & "\res"
                        Using RunModifyXhtmlDocument As New Process
                            With RunModifyXhtmlDocument
                                .StartInfo.FileName = ProgramFiles & "\Internet Explorer\iexplore.exe"
                                .StartInfo.Arguments = """" & ResFolder & "\modify.xht"""
                                .Start()
                            End With
                        End Using
                    End If
                    ModifyWizard.Close()
                End If
            End If
        End If
    End Sub
End Class