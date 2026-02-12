Imports Microsoft.Win32
Imports System.ComponentModel
Imports System.Configuration
Imports System.IO
Imports System.IO.Compression
Imports System.Security.AccessControl
Imports System.Security.Principal

Public Class SetupInstall
    Private RestorePoint As Object = GetObject("winmgmts:\\.\root\default:Systemrestore")
    Private WithEvents BgWorker As BackgroundWorker
    Dim RestorePointCreated As Boolean = False
    Dim CloseProgram As Boolean = True
    Dim ShowMessageBoxForRestartOnExit As Boolean = False
    Dim ResFolder As String = SetupWizard.ProgDir & "\res"
    Dim LocalAppData As String = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData)
    Dim ProgramFiles As String = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles)
    Dim WinDir As String = Environment.GetFolderPath(Environment.SpecialFolder.Windows)

    Private Sub SetupInstall_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown
        ' Prepare DetailsTextBox
        DetailsTextBox.ScrollBars = ScrollBars.Vertical

        ' Initialise progress bar -> Set as marquee
        With ProgressBarSetup
            .Style = ProgressBarStyle.Marquee
            .MarqueeAnimationSpeed = 100
            .Refresh()
        End With

        ' Wait 2 seconds prior starting setup
        Wait(2)

        ' Initialise BackgroundWorker
        InitialiseBackgroundWorker()

        ' Close explorer.exe
        BgWorker.RunWorkerAsync("CloseExplorer")
        WaitUntilTaskFinishes()

        ' Create system restore point
        ActionLabel.Text = "Creating a system restore point..."
        ActionLabel.Refresh()
        DetailsTextBox.AppendText(ActionLabel.Text)
        CreateRestorePoint()
    End Sub

    Private Sub InitialiseBackgroundWorker()
        BgWorker = New BackgroundWorker
        BgWorker.WorkerSupportsCancellation = True
    End Sub

    Private Sub Wait(ByVal SecondsToWait As Integer)
        For Index As Integer = 0 To SecondsToWait * 100
            System.Threading.Thread.Sleep(10)
            Application.DoEvents()
        Next
    End Sub

    Private Sub WaitUntilTaskFinishes()
        Do
            Application.DoEvents()
        Loop While BgWorker.IsBusy
    End Sub

    Private Sub CreateLogFile(Optional ByVal SaveTo As String = Nothing)
        If SaveTo = Nothing Then
            SaveTo = SetupWizard.ProgDir
        End If
        File.AppendAllText(SaveTo & "\win8rpconv_install.log", "Installation date: " & My.Computer.Clock.LocalTime.ToShortDateString & vbCrLf)
        File.AppendAllText(SaveTo & "\win8rpconv_install.log", "Setup options: [Theme: """ & SetupWizard.AddressBarStyle & """, InstAeroGlass: """ & SetupWizard.InstAeroGlass & """, InstONE: """ & SetupWizard.InstONE & """, InstUX: """ & SetupWizard.InstUX & """, InstQuero: """ & SetupWizard.InstQuero & """, InstSysFiles: """ & SetupWizard.InstSysFiles & """, InstSounds: """ & SetupWizard.InstSounds & """, Inst7TaskTw: """ & SetupWizard.Inst7TaskTw & """]" & vbCrLf & "Install log:" & vbCrLf)
        File.AppendAllLines(SaveTo & "\win8rpconv_install.log", DetailsTextBox.Lines)
    End Sub

    Private Sub ExtractFiles(ByVal FileFromResources As Byte(), ByVal ExtractZipTo As String)
        Using ZipFileToOpen As New MemoryStream(FileFromResources)
            Using ResZipFile As New ZipArchive(ZipFileToOpen, ZipArchiveMode.Read)
                ResZipFile.ExtractToDirectory(ExtractZipTo)
            End Using
        End Using
    End Sub

    Private Sub BgWorker_DoWork(sender As Object, e As DoWorkEventArgs) Handles BgWorker.DoWork
        Dim WhatToDo As String = Convert.ToString(e.Argument)
        Select Case WhatToDo
            Case "CloseExplorer"
                Using EndExplorerTask As New Process
                    With EndExplorerTask
                        .StartInfo.FileName = WinDir & "\System32\taskkill.exe"
                        .StartInfo.Arguments = "/f /im explorer.exe"
                        .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                        .Start()
                        .WaitForExit()
                    End With
                End Using
            Case "CreateRestorePoint"
                '' Creating system restore point
                If RestorePoint.CreateRestorePoint("Installed Windows 8.1 to 8 RP Converter", 0, 100) = 0 Then
                    RestorePointCreated = True
                Else
                    RestorePointCreated = False
                End If
            Case "ExtractResourcesFolder"
                '' Extract resources file from program resources
                ExtractFiles(My.Resources.ResourceFile, ResFolder)
            Case "InstAeroGlass"
                '' Change install directory to user’s root drive
                Dim AeroGlassSetupInfFilePath As String = ResFolder & "\AeroGlass\settings.inf"
                Dim AeroGlassSetupInfFile As String() = File.ReadAllLines(AeroGlassSetupInfFilePath)
                AeroGlassSetupInfFile(2) = "Dir=" & Environ("SystemDrive") & "\AeroGlass"
                File.WriteAllLines(AeroGlassSetupInfFilePath, AeroGlassSetupInfFile)
                Erase AeroGlassSetupInfFile
                '' Start Aero Glass setup
                Using AeroGlassSetup As New Process
                    With AeroGlassSetup
                        .StartInfo.FileName = ResFolder & "\AeroGlass\setup-w8.1-1.4.6.exe"
                        .StartInfo.Arguments = "/VERYSILENT /LOADINF=""" & AeroGlassSetupInfFilePath & """"
                        .Start()
                        .WaitForExit()
                    End With
                End Using
                Using AeroRegKey As RegistryKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\Microsoft\Windows\DWM", True)
                    AeroRegKey.SetValue("CustomThemeAtlas", "")
                    AeroRegKey.Close()
                End Using
            Case "InstONE"
                Dim SourcePathONE As String = ResFolder & "\OldNewExplorer"
                Dim DestinationPathONE As String = ProgramFiles & "\OldNewExplorer"
                FileIO.FileSystem.CopyDirectory(SourcePathONE, DestinationPathONE)
                If SetupWizard.BitnessSystem = "64Bit" Then
                    If DetailsTextBox.InvokeRequired Then
                        DetailsTextBox.Invoke(Sub() DetailsTextBox.AppendText("Registering file " & DestinationPathONE & "\OldNewExplorer64.dll..."))
                    Else
                        DetailsTextBox.AppendText("Registering file " & DestinationPathONE & "\OldNewExplorer64.dll...")
                    End If
                    Using ONESetup As New Process
                        With ONESetup
                            .StartInfo.FileName = WinDir & "\System32\regsvr32.exe"
                            .StartInfo.Arguments = "/s """ & DestinationPathONE & "\OldNewExplorer64.dll"""
                            .Start()
                            .WaitForExit()
                        End With
                    End Using
                    If DetailsTextBox.InvokeRequired Then
                        DetailsTextBox.Invoke(Sub()
                                                  DetailsTextBox.AppendText(" Finished!" & vbCrLf)
                                                  DetailsTextBox.AppendText("Registering file " & DestinationPathONE & "\OldNewExplorer32.dll...")
                                              End Sub)
                    Else
                        DetailsTextBox.AppendText(" Finished!" & vbCrLf)
                        DetailsTextBox.AppendText("Registering file " & DestinationPathONE & "\OldNewExplorer32.dll...")
                    End If
                    Using ONESetup As New Process
                        With ONESetup
                            .StartInfo.FileName = WinDir & "\System32\regsvr32.exe"
                            .StartInfo.Arguments = "/s """ & DestinationPathONE & "\OldNewExplorer32.dll"""
                            .Start()
                            .WaitForExit()
                        End With
                    End Using
                    If DetailsTextBox.InvokeRequired Then
                        DetailsTextBox.Invoke(Sub() DetailsTextBox.AppendText(" Finished!" & vbCrLf))
                    Else
                        DetailsTextBox.AppendText(" Finished!" & vbCrLf)
                    End If
                ElseIf SetupWizard.BitnessSystem = "32Bit" Then
                    If DetailsTextBox.InvokeRequired Then
                        DetailsTextBox.Invoke(Sub() DetailsTextBox.AppendText("Registering file " & DestinationPathONE & "\OldNewExplorer32.dll..."))
                    Else
                        DetailsTextBox.AppendText("Registering file " & DestinationPathONE & "\OldNewExplorer32.dll...")
                    End If
                    Using ONESetup As New Process
                        With ONESetup
                            .StartInfo.FileName = WinDir & "\System32\regsvr32.exe"
                            .StartInfo.Arguments = "/s """ & DestinationPathONE & "\OldNewExplorer32.dll"""
                            .Start()
                            .WaitForExit()
                        End With
                    End Using
                    If DetailsTextBox.InvokeRequired Then
                        DetailsTextBox.Invoke(Sub() DetailsTextBox.AppendText(" Finished!" & vbCrLf))
                    Else
                        DetailsTextBox.AppendText(" Finished!" & vbCrLf)
                    End If
                End If
                '' Create shortcut on desktop
                If DetailsTextBox.InvokeRequired Then
                    DetailsTextBox.Invoke(Sub() DetailsTextBox.AppendText("Creating shortcut on desktop..."))
                Else
                    DetailsTextBox.AppendText("Creating shortcut on desktop...")
                End If
                Using ObjectShell As Object = CreateObject("WScript.Shell")
                    Using ObjectLink As Object = ObjectShell.CreateShortcut(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) & "\OldNewExplorer Configuration.lnk")
                        Try
                            With ObjectLink
                                .TargetPath = ProgramFiles & "\OldNewExplorer\OldNewExplorerCfg.exe"
                                .WorkingDirectory = ProgramFiles & "\OldNewExplorer"
                                .WindowStyle = 1
                                .Save()
                            End With
                            If DetailsTextBox.InvokeRequired Then
                                DetailsTextBox.Invoke(Sub() DetailsTextBox.AppendText(" Finished!" & vbCrLf))
                            Else
                                DetailsTextBox.AppendText(" Finished!" & vbCrLf)
                            End If
                        Catch ex As Exception
                            MessageBox.Show("Shortcut to OldNewExplorer configuration program could not be created!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                            If DetailsTextBox.InvokeRequired Then
                                DetailsTextBox.Invoke(Sub() DetailsTextBox.AppendText(" Failed!" & vbCrLf))
                            Else
                                DetailsTextBox.AppendText(" Failed!" & vbCrLf)
                            End If
                        End Try
                    End Using
                End Using
            Case "InstUltraUX"
                Using UltraUXSetup As New Process
                    With UltraUXSetup
                        .StartInfo.FileName = ResFolder & "\UltraUX\UltraUXThemePatcher_4.5.0.exe"
                        .Start()
                        .WaitForExit()
                    End With
                End Using
            Case "InstUXTSB"
                If SetupWizard.BitnessSystem = "64Bit" Then
                    If DetailsTextBox.InvokeRequired Then
                        DetailsTextBox.Invoke(Sub() DetailsTextBox.AppendText("Copying file UxThemeSignatureBypass64.dll to " & WinDir & "\System32\UxThemeSignatureBypass.dll..."))
                    Else
                        DetailsTextBox.AppendText("Copying file UxThemeSignatureBypass64.dll to " & WinDir & "\System32\UxThemeSignatureBypass.dll...")
                    End If
                    FileIO.FileSystem.CopyFile(ResFolder & "\UXThemeDLL\UxThemeSignatureBypass64.dll", WinDir & "\System32\UxThemeSignatureBypass.dll")
                    If DetailsTextBox.InvokeRequired Then
                        DetailsTextBox.Invoke(Sub() DetailsTextBox.AppendText(" Finished!" & vbCrLf))
                    Else
                        DetailsTextBox.AppendText(" Finished!" & vbCrLf)
                    End If
                    If DetailsTextBox.InvokeRequired Then
                        DetailsTextBox.Invoke(Sub() DetailsTextBox.AppendText("Copying file UxThemeSignatureBypass32.dll to " & WinDir & "\SysWOW64\UxThemeSignatureBypass.dll..."))
                    Else
                        DetailsTextBox.AppendText("Copying file UxThemeSignatureBypass32.dll to " & WinDir & "\SysWOW64\UxThemeSignatureBypass.dll...")
                    End If
                    FileIO.FileSystem.CopyFile(ResFolder & "\UXThemeDLL\UxThemeSignatureBypass32.dll", WinDir & "\SysWOW64\UxThemeSignatureBypass.dll")
                    If DetailsTextBox.InvokeRequired Then
                        DetailsTextBox.Invoke(Sub() DetailsTextBox.AppendText(" Finished!" & vbCrLf))
                    Else
                        DetailsTextBox.AppendText(" Finished!" & vbCrLf)
                    End If
                ElseIf SetupWizard.BitnessSystem = "32Bit" Then
                    If DetailsTextBox.InvokeRequired Then
                        DetailsTextBox.Invoke(Sub() DetailsTextBox.AppendText("Copying file UxThemeSignatureBypass32.dll to " & WinDir & "\System32\UxThemeSignatureBypass.dll..."))
                    Else
                        DetailsTextBox.AppendText("Copying file UxThemeSignatureBypass32.dll to " & WinDir & "\System32\UxThemeSignatureBypass.dll...")
                    End If
                    FileIO.FileSystem.CopyFile(ResFolder & "\UXThemeDLL\UxThemeSignatureBypass32.dll", WinDir & "\System32\UxThemeSignatureBypass.dll")
                    If DetailsTextBox.InvokeRequired Then
                        DetailsTextBox.Invoke(Sub() DetailsTextBox.AppendText(" Finished!" & vbCrLf))
                    Else
                        DetailsTextBox.AppendText(" Finished!" & vbCrLf)
                    End If
                End If
                '' UXTSB\Registry
                If DetailsTextBox.InvokeRequired Then
                    DetailsTextBox.Invoke(Sub() DetailsTextBox.AppendText("Registering files..."))
                Else
                    DetailsTextBox.AppendText("Registering files...")
                End If
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
                If SetupWizard.BitnessSystem = "64Bit" Then
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
            Case "InstQuero"
                '' Change install directory to ProgramFiles folder
                Dim QueroSetupInfFilePath As String = ResFolder & "\Quero\settings.inf"
                Dim QueroSetupInfFile As String() = File.ReadAllLines(QueroSetupInfFilePath)
                QueroSetupInfFile(2) = "Dir=" & ProgramFiles & "\Quero Toolbar"
                File.WriteAllLines(QueroSetupInfFilePath, QueroSetupInfFile)
                Erase QueroSetupInfFile
                '' Start Quero Toolbar setup
                Using QueroSetup As New Process
                    With QueroSetup
                        If SetupWizard.BitnessSystem = "64Bit" Then
                            .StartInfo.FileName = ResFolder & "\Quero\QueroToolbarInstaller_x64.exe"
                        ElseIf SetupWizard.BitnessSystem = "32Bit" Then
                            .StartInfo.FileName = ResFolder & "\Quero\QueroToolbarInstaller_x86.exe"
                        End If
                        .StartInfo.Arguments = "/VERYSILENT /LOADINF=""" & QueroSetupInfFilePath & """"
                        .Start()
                        .WaitForExit()
                    End With
                End Using
            Case "InstSysFiles"
                Dim SourcePathSysFiles As String = ResFolder & "\UIRibbon"
                Dim SysFilesSetup As New Process
                Dim EveryoneSidAsUserName As String = New SecurityIdentifier(WellKnownSidType.WorldSid, Nothing).Translate(GetType(NTAccount)).Value
                Dim EveryoneAuditRule As FileSystemAuditRule = New FileSystemAuditRule(EveryoneSidAsUserName, FileSystemRights.CreateFiles + FileSystemRights.CreateDirectories + FileSystemRights.WriteAttributes + FileSystemRights.WriteExtendedAttributes + FileSystemRights.Delete + FileSystemRights.ChangePermissions + FileSystemRights.TakeOwnership, AuditFlags.Success + AuditFlags.Failure)
                Dim UiRibbonAuditRules As New FileSecurity
                '' Replace the files UIRibbon.dll, UIRibbonRes.dll and <UILanguage>\UIRibbon.dll.mui
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
                If DetailsTextBox.InvokeRequired Then
                    DetailsTextBox.Invoke(Sub() DetailsTextBox.AppendText("Creating backup of " & WinDir & "\System32\UIRibbon.dll..."))
                Else
                    DetailsTextBox.AppendText("Creating backup of " & WinDir & "\System32\UIRibbon.dll...")
                End If
                FileIO.FileSystem.RenameFile(WinDir & "\System32\UIRibbon.dll", "UIRibbon.dll.bak")
                If DetailsTextBox.InvokeRequired Then
                    DetailsTextBox.Invoke(Sub()
                                              DetailsTextBox.AppendText(" Finished!" & vbCrLf)
                                              DetailsTextBox.AppendText("Creating backup of " & WinDir & "\System32\UIRibbonRes.dll...")
                                          End Sub)
                Else
                    DetailsTextBox.AppendText(" Finished!" & vbCrLf)
                    DetailsTextBox.AppendText("Creating backup of " & WinDir & "\System32\UIRibbonRes.dll...")
                End If
                FileIO.FileSystem.RenameFile(WinDir & "\System32\UIRibbonRes.dll", "UIRibbonRes.dll.bak")
                If DetailsTextBox.InvokeRequired Then
                    DetailsTextBox.Invoke(Sub() DetailsTextBox.AppendText(" Finished!" & vbCrLf))
                Else
                    DetailsTextBox.AppendText(" Finished!" & vbCrLf)
                End If
                With SysFilesSetup
                    .StartInfo.FileName = WinDir & "\System32\takeown.exe"
                    .StartInfo.Arguments = "/f " & WinDir & "\System32\" & SetupWizard.ProgLanguage & "\UIRibbon.dll.mui /a"
                    .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                    .Start()
                    .WaitForExit()
                End With
                With SysFilesSetup
                    .StartInfo.FileName = WinDir & "\System32\icacls.exe"
                    .StartInfo.Arguments = WinDir & "\System32\" & SetupWizard.ProgLanguage & "\UIRibbon.dll.mui /grant *S-1-5-32-545:F *S-1-5-32-544:F"
                    .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                    .Start()
                    .WaitForExit()
                End With
                If DetailsTextBox.InvokeRequired Then
                    DetailsTextBox.Invoke(Sub() DetailsTextBox.AppendText("Creating backup of " & WinDir & "\System32\" & SetupWizard.ProgLanguage & "\UIRibbon.dll.mui..."))
                Else
                    DetailsTextBox.AppendText("Creating backup of " & WinDir & "\System32\" & SetupWizard.ProgLanguage & "\UIRibbon.dll.mui...")
                End If
                FileIO.FileSystem.RenameFile(WinDir & "\System32\" & SetupWizard.ProgLanguage & "\UIRibbon.dll.mui", "UIRibbon.dll.mui.bak")
                If DetailsTextBox.InvokeRequired Then
                    DetailsTextBox.Invoke(Sub()
                                              DetailsTextBox.AppendText(" Finished!" & vbCrLf)
                                              DetailsTextBox.AppendText("Replacing files...")
                                          End Sub)
                Else
                    DetailsTextBox.AppendText(" Finished!" & vbCrLf)
                    DetailsTextBox.AppendText("Replacing files...")
                End If
                If SetupWizard.BitnessSystem = "64Bit" Then
                    FileIO.FileSystem.CopyFile(SourcePathSysFiles & "\system32\UIRibbon.dll", WinDir & "\System32\UIRibbon.dll")
                    FileIO.FileSystem.CopyFile(SourcePathSysFiles & "\system32\UIRibbonRes.dll", WinDir & "\System32\UIRibbonRes.dll")
                    FileIO.FileSystem.CopyFile(SourcePathSysFiles & "\system32\" & SetupWizard.ProgLanguage & "\UIRibbon.dll.mui", WinDir & "\System32\" & SetupWizard.ProgLanguage & "\UIRibbon.dll.mui")
                ElseIf SetupWizard.BitnessSystem = "32Bit" Then
                    FileIO.FileSystem.CopyFile(SourcePathSysFiles & "\syswow64\UIRibbon.dll", WinDir & "\System32\UIRibbon.dll")
                    FileIO.FileSystem.CopyFile(SourcePathSysFiles & "\syswow64\UIRibbonRes.dll", WinDir & "\System32\UIRibbonRes.dll")
                    FileIO.FileSystem.CopyFile(SourcePathSysFiles & "\syswow64\" & SetupWizard.ProgLanguage & "\UIRibbon.dll.mui", WinDir & "\System32\" & SetupWizard.ProgLanguage & "\UIRibbon.dll.mui")
                End If
                If DetailsTextBox.InvokeRequired Then
                    DetailsTextBox.Invoke(Sub() DetailsTextBox.AppendText(" Finished!" & vbCrLf))
                Else
                    DetailsTextBox.AppendText(" Finished!" & vbCrLf)
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
                    .StartInfo.Arguments = WinDir & "\System32\" & SetupWizard.ProgLanguage & "\UIRibbon.dll.mui /setowner ""NT SERVICE\TrustedInstaller"""
                    .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                    .Start()
                    .WaitForExit()
                End With
                With SysFilesSetup
                    .StartInfo.FileName = WinDir & "\System32\icacls.exe"
                    .StartInfo.Arguments = WinDir & "\System32\" & SetupWizard.ProgLanguage & "\UIRibbon.dll.mui /reset"
                    .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                    .Start()
                    .WaitForExit()
                End With
                With SysFilesSetup
                    .StartInfo.FileName = WinDir & "\System32\icacls.exe"
                    .StartInfo.Arguments = WinDir & "\System32\" & SetupWizard.ProgLanguage & "\UIRibbon.dll.mui /grant *S-1-5-18:RX *S-1-5-32-544:RX *S-1-5-32-545:RX *S-1-15-2-1:RX ""NT SERVICE\TrustedInstaller"":F /inheritance:r"
                    .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                    .Start()
                    .WaitForExit()
                End With
                UiRibbonAuditRules = File.GetAccessControl(WinDir & "\System32\" & SetupWizard.ProgLanguage & "\UIRibbon.dll.mui", AccessControlSections.Audit)
                UiRibbonAuditRules.SetAuditRule(EveryoneAuditRule)
                File.SetAccessControl(WinDir & "\System32\" & SetupWizard.ProgLanguage & "\UIRibbon.dll.mui", UiRibbonAuditRules)
                UiRibbonAuditRules = Nothing
                '' If System is 64-bit, repeat for files in SysWOW64
                If SetupWizard.BitnessSystem = "64Bit" Then
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
                    If DetailsTextBox.InvokeRequired Then
                        DetailsTextBox.Invoke(Sub() DetailsTextBox.AppendText("Creating backup of " & WinDir & "\SysWOW64\UIRibbon.dll..."))
                    Else
                        DetailsTextBox.AppendText("Creating backup of " & WinDir & "\SysWOW64\UIRibbon.dll...")
                    End If
                    FileIO.FileSystem.RenameFile(WinDir & "\SysWOW64\UIRibbon.dll", "UIRibbon.dll.bak")
                    If DetailsTextBox.InvokeRequired Then
                        DetailsTextBox.Invoke(Sub()
                                                  DetailsTextBox.AppendText(" Finished!" & vbCrLf)
                                                  DetailsTextBox.AppendText("Creating backup of " & WinDir & "\SysWOW64\UIRibbonRes.dll...")
                                              End Sub)
                    Else
                        DetailsTextBox.AppendText(" Finished!" & vbCrLf)
                        DetailsTextBox.AppendText("Creating backup of " & WinDir & "\SysWOW64\UIRibbonRes.dll...")
                    End If
                    FileIO.FileSystem.RenameFile(WinDir & "\SysWOW64\UIRibbonRes.dll", "UIRibbonRes.dll.bak")
                    If DetailsTextBox.InvokeRequired Then
                        DetailsTextBox.Invoke(Sub() DetailsTextBox.AppendText(" Finished!" & vbCrLf))
                    Else
                        DetailsTextBox.AppendText(" Finished!" & vbCrLf)
                    End If
                    With SysFilesSetup
                        .StartInfo.FileName = WinDir & "\System32\takeown.exe"
                        .StartInfo.Arguments = "/f " & WinDir & "\SysWOW64\" & SetupWizard.ProgLanguage & "\UIRibbon.dll.mui /a"
                        .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                        .Start()
                        .WaitForExit()
                    End With
                    With SysFilesSetup
                        .StartInfo.FileName = WinDir & "\System32\icacls.exe"
                        .StartInfo.Arguments = WinDir & "\SysWOW64\" & SetupWizard.ProgLanguage & "\UIRibbon.dll.mui /grant *S-1-5-32-545:F *S-1-5-32-544:F"
                        .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                        .Start()
                        .WaitForExit()
                    End With
                    If DetailsTextBox.InvokeRequired Then
                        DetailsTextBox.Invoke(Sub() DetailsTextBox.AppendText("Creating backup of " & WinDir & "\SysWOW64\" & SetupWizard.ProgLanguage & "\UIRibbon.dll.mui..."))
                    Else
                        DetailsTextBox.AppendText("Creating backup of " & WinDir & "\SysWOW64\" & SetupWizard.ProgLanguage & "\UIRibbon.dll.mui...")
                    End If
                    FileIO.FileSystem.RenameFile(WinDir & "\SysWOW64\" & SetupWizard.ProgLanguage & "\UIRibbon.dll.mui", "UIRibbon.dll.mui.bak")
                    If DetailsTextBox.InvokeRequired Then
                        DetailsTextBox.Invoke(Sub()
                                                  DetailsTextBox.AppendText(" Finished!" & vbCrLf)
                                                  DetailsTextBox.AppendText("Replacing files...")
                                              End Sub)
                    Else
                        DetailsTextBox.AppendText(" Finished!" & vbCrLf)
                        DetailsTextBox.AppendText("Replacing files...")
                    End If
                    FileIO.FileSystem.CopyFile(SourcePathSysFiles & "\syswow64\UIRibbon.dll", WinDir & "\SysWOW64\UIRibbon.dll")
                    FileIO.FileSystem.CopyFile(SourcePathSysFiles & "\syswow64\UIRibbonRes.dll", WinDir & "\SysWOW64\UIRibbonRes.dll")
                    FileIO.FileSystem.CopyFile(SourcePathSysFiles & "\syswow64\" & SetupWizard.ProgLanguage & "\UIRibbon.dll.mui", WinDir & "\SysWOW64\" & SetupWizard.ProgLanguage & "\UIRibbon.dll.mui")
                    If DetailsTextBox.InvokeRequired Then
                        DetailsTextBox.Invoke(Sub() DetailsTextBox.AppendText(" Finished!" & vbCrLf))
                    Else
                        DetailsTextBox.AppendText(" Finished!" & vbCrLf)
                    End If
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
                        .StartInfo.Arguments = WinDir & "\SysWOW64\" & SetupWizard.ProgLanguage & "\UIRibbon.dll.mui /setowner ""NT SERVICE\TrustedInstaller"""
                        .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                        .Start()
                        .WaitForExit()
                    End With
                    With SysFilesSetup
                        .StartInfo.FileName = WinDir & "\System32\icacls.exe"
                        .StartInfo.Arguments = WinDir & "\SysWOW64\" & SetupWizard.ProgLanguage & "\UIRibbon.dll.mui /reset"
                        .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                        .Start()
                        .WaitForExit()
                    End With
                    With SysFilesSetup
                        .StartInfo.FileName = WinDir & "\System32\icacls.exe"
                        .StartInfo.Arguments = WinDir & "\SysWOW64\" & SetupWizard.ProgLanguage & "\UIRibbon.dll.mui /grant *S-1-5-18:RX *S-1-5-32-544:RX *S-1-5-32-545:RX *S-1-15-2-1:RX ""NT SERVICE\TrustedInstaller"":F /inheritance:r"
                        .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                        .Start()
                        .WaitForExit()
                    End With
                    UiRibbonAuditRules = File.GetAccessControl(WinDir & "\SysWOW64\" & SetupWizard.ProgLanguage & "\UIRibbon.dll.mui", AccessControlSections.Audit)
                    UiRibbonAuditRules.SetAuditRule(EveryoneAuditRule)
                    File.SetAccessControl(WinDir & "\SysWOW64\" & SetupWizard.ProgLanguage & "\UIRibbon.dll.mui", UiRibbonAuditRules)
                    UiRibbonAuditRules = Nothing
                End If
                SysFilesSetup.Dispose()
            Case "InstGadgets"
                Using GadgetsSetup As New Process
                    With GadgetsSetup
                        .StartInfo.FileName = WinDir & "\System32\msiexec.exe"
                        .StartInfo.Arguments = "/package """ & ResFolder & "\Gadgets\8GadgetPack370Setup.msi"" /quiet"
                        .Start()
                        .WaitForExit()
                    End With
                End Using
            Case "Inst7TaskTw"
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
        End Select

        ' CancelAsync of BackgroundWorker after task finishes
        BgWorker.CancelAsync()
    End Sub

    Private Sub CreateRestorePoint()
        ' Check if system protection is enabled
        Using SystemRestoreRegKey As RegistryKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\Microsoft\Windows NT\CurrentVersion\SystemRestore", False)
            If SystemRestoreRegKey.GetValue("RPSessionInterval") = 1 Then
                '' If enabled then create system restore point
                If RestorePoint IsNot Nothing Then
                    BgWorker.RunWorkerAsync("CreateRestorePoint")
                    WaitUntilTaskFinishes()
                    If RestorePointCreated = False Then
                        If MessageBox.Show("System restore point has not been created! It is not recommended to install this without creating a restore point. Do you want to proceed with the installation anyway?", "System restore point not created", MessageBoxButtons.YesNo, MessageBoxIcon.Error) = System.Windows.Forms.DialogResult.Yes Then
                            SetupLauncher()
                        ElseIf System.Windows.Forms.DialogResult.No Then
                            SetupWizard.Close()
                        End If
                    Else
                        ' Start setup now
                        SetupLauncher()
                    End If
                End If
            Else
                '' If not, ask if it should be enabled now
                If MessageBox.Show("System protection is disabled. Do you want to enable it now to create a system restore point?", "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = System.Windows.Forms.DialogResult.Yes Then
                    If RestorePoint IsNot Nothing Then
                        If RestorePoint.Enable(Environ("SystemDrive") & "\") = 0 Then
                            BgWorker.WorkerSupportsCancellation = True
                            BgWorker.RunWorkerAsync("CreateRestorePoint")
                            WaitUntilTaskFinishes()
                        Else
                            If MessageBox.Show("System protection could not be enabled! It is not recommended to install it without creating a restore point. Do you want to proceed with the installation anyway?", "System protection not enabled", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) = System.Windows.Forms.DialogResult.Yes Then
                                RestorePointCreated = False
                                SetupLauncher()
                            ElseIf System.Windows.Forms.DialogResult.No Then
                                SetupWizard.Close()
                            End If
                        End If
                    End If
                ElseIf System.Windows.Forms.DialogResult.No Then
                    If MessageBox.Show("System restore point has not been created! It is not recommended to install it without creating a restore point. Do you want to proceed with the installation anyway?", "System restore point not created", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) = System.Windows.Forms.DialogResult.Yes Then
                        RestorePointCreated = False
                        SetupLauncher()
                    ElseIf System.Windows.Forms.DialogResult.No Then
                        SetupWizard.Close()
                    End If
                End If
            End If
        End Using
    End Sub

    Private Sub SetupLauncher()
        ' Check if system restore point has been created
        If RestorePointCreated = True Then
            '' Change progress bar
            With ProgressBarSetup
                .Style = ProgressBarStyle.Continuous
                .Value = 0
                .Minimum = 0
                .Maximum = 10
                .Refresh()
                .PerformStep()
                .Refresh()
            End With
            ActionLabel.Text = "Restore point has been successfully created!"
            ActionLabel.Refresh()
            DetailsTextBox.AppendText(" Finished!" & vbCrLf)
            CloseButton.Enabled = False
            CloseProgram = False
        Else
            ActionLabel.Text = "Restore point not created! Starting setup..."
            ActionLabel.Refresh()
            With ProgressBarSetup
                .MarqueeAnimationSpeed = 0
                .Value = 0
                .Refresh()
            End With
            DetailsTextBox.AppendText(" Failed or denied!" & vbCrLf)
            CloseButton.Enabled = False
            CloseProgram = False
        End If

        ' Waiting 2 seconds prior starting setup
        Wait(2)

        ' Run setup now
        Try
            RunSetup()
        Catch ex As Exception
            '' If error occurs, create log file
            MessageBox.Show("An error occured during installation: " & ex.Message & " Please check the log file created on the desktop to check what happened, and try it again.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            DetailsTextBox.AppendText(vbCrLf & "ERROR: " & ex.Message & vbCrLf)
            '' Enable close button
            CloseProgram = True
            CloseButton.Enabled = True
            CloseButton.Text = "&Close"
            '' Create log file
            CreateLogFile(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory))
            '' Delete ProgramFiles directory
            If FileIO.FileSystem.DirectoryExists(SetupWizard.ProgDir) Then
                FileIO.FileSystem.DeleteDirectory(SetupWizard.ProgDir, FileIO.DeleteDirectoryOption.DeleteAllContents)
            End If
            '' Close program
            SetupWizard.Close()
        End Try
    End Sub

    Private Sub InitialiseProgressBar()
        ' Initialise progress bar
        Dim ProgBarCount As Integer = 40
        If SetupWizard.InstAeroGlass = True Then
            ProgBarCount += 10
        End If
        If SetupWizard.InstONE = True Then
            ProgBarCount += 10
        End If
        If SetupWizard.InstUX <> "NoUXPatch" Then
            ProgBarCount += 10
        End If
        If SetupWizard.InstQuero = True Then
            ProgBarCount += 10
        End If
        If SetupWizard.InstSysFiles = True Then
            ProgBarCount += 10
        End If
        If SetupWizard.InstSounds = True Then
            ProgBarCount += 10
        End If
        If SetupWizard.InstGadgets = True Then
            ProgBarCount += 10
        End If
        If SetupWizard.Inst7TaskTw = True Then
            ProgBarCount += 10
        End If

        ' Set up progress bar
        With ProgressBarSetup
            .Style = ProgressBarStyle.Continuous
            .Value = 0
            .Maximum = ProgBarCount
            .Refresh()
        End With
    End Sub

    Private Sub RunSetup()
        InitialiseProgressBar()
        ' Extract resources file
        ActionLabel.Text = "Extracting resources file..."
        ActionLabel.Refresh()
        DetailsTextBox.AppendText(ActionLabel.Text)
        BgWorker.RunWorkerAsync("ExtractResourcesFolder")
        WaitUntilTaskFinishes()
        '' Check if file has been extracted -> folder has been created
        If FileIO.FileSystem.DirectoryExists(ResFolder) Then
            ProgressBarSetup.PerformStep()
            ProgressBarSetup.Refresh()
            ActionLabel.Text = "Resources file has been extracted!"
            ActionLabel.Refresh()
            DetailsTextBox.AppendText(" Finished!" & vbCrLf)
        Else
            MessageBox.Show("Resources file could not be extracted!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            SetupWizard.Close()
        End If

        ' Install Aero Glass
        If SetupWizard.AeroGlassInstalled = False Or (SetupWizard.AeroGlassInstalled = True And SetupWizard.InstAeroGlass = True) Then
            ActionLabel.Text = "Installing Aero Glass..."
            ActionLabel.Refresh()
            DetailsTextBox.AppendText(ActionLabel.Text)
            BgWorker.RunWorkerAsync("InstAeroGlass")
            WaitUntilTaskFinishes()
            ProgressBarSetup.PerformStep()
            ProgressBarSetup.Refresh()
            ActionLabel.Text = "Aero Glass has been installed!"
            ActionLabel.Refresh()
            DetailsTextBox.AppendText(" Finished!" & vbCrLf)
        End If

        ' Install OldNewExplorer
        If SetupWizard.InstONE = True Then
            ActionLabel.Text = "Installing OldNewExplorer..."
            ActionLabel.Refresh()
            DetailsTextBox.AppendText(ActionLabel.Text & vbCrLf)
            BgWorker.RunWorkerAsync("InstONE")
            WaitUntilTaskFinishes()
            ProgressBarSetup.PerformStep()
            ProgressBarSetup.Refresh()
            ActionLabel.Text = "OldNewExplorer has been installed!"
            ActionLabel.Refresh()
        End If

        ' Install UXTheme patcher
        If SetupWizard.InstUX = "UltraUX" Then
            '' Install UltraUX
            ActionLabel.Text = "Installing UltraUXThemePatcher..."
            ActionLabel.Refresh()
            DetailsTextBox.AppendText(ActionLabel.Text)
            BgWorker.RunWorkerAsync("InstUltraUX")
            WaitUntilTaskFinishes()
            ProgressBarSetup.PerformStep()
            ProgressBarSetup.Refresh()
            ActionLabel.Text = "UltraUXThemePatcher has been installed!"
            ActionLabel.Refresh()
            DetailsTextBox.AppendText(" Finished!" & vbCrLf)
        ElseIf SetupWizard.InstUX = "UXTSB" Then
            '' Install UXTSB
            ActionLabel.Text = "Installing UXThemeSignatureBypass..."
            ActionLabel.Refresh()
            DetailsTextBox.AppendText(ActionLabel.Text & vbCrLf)
            BgWorker.RunWorkerAsync("InstUXTSB")
            WaitUntilTaskFinishes()
            ProgressBarSetup.PerformStep()
            ProgressBarSetup.Refresh()
            ActionLabel.Text = "UXThemeSignatureBypass has been installed!"
            ActionLabel.Refresh()
            DetailsTextBox.AppendText(" Finished!" & vbCrLf)
        End If

        ' Install Quero Toolbar
        If SetupWizard.InstQuero = True Then
            ActionLabel.Text = "Installing Quero Toolbar..."
            ActionLabel.Refresh()
            DetailsTextBox.AppendText(ActionLabel.Text)
            BgWorker.RunWorkerAsync("InstQuero")
            WaitUntilTaskFinishes()
            ProgressBarSetup.PerformStep()
            ProgressBarSetup.Refresh()
            ActionLabel.Text = "Quero Toolbar has been installed!"
            ActionLabel.Refresh()
            DetailsTextBox.AppendText(" Finished!" & vbCrLf)
        End If

        ' Replace UIRibbon.dll & UIRibbonRes.dll
        If SetupWizard.InstSysFiles = True Then
            ActionLabel.Text = "Replacing UIRibbon.dll && UIRibbonRes.dll..."
            ActionLabel.Refresh()
            DetailsTextBox.AppendText("Replacing UIRibbon.dll & UIRibbonRes.dll..." & vbCrLf)
            BgWorker.RunWorkerAsync("InstSysFiles")
            WaitUntilTaskFinishes()
            ProgressBarSetup.PerformStep()
            ProgressBarSetup.Refresh()
            ActionLabel.Text = "UIRibbon.dll && UIRibbonRes.dll have been replaced!"
            ActionLabel.Refresh()
        End If

        ' Installing theme
        ActionLabel.Text = "Installing theme..."
        ActionLabel.Refresh()
        DetailsTextBox.AppendText(ActionLabel.Text)
        FileIO.FileSystem.CopyDirectory(ResFolder & "\Theme\Themes", WinDir & "\Resources\Themes", True)
        '' Blue/white address bar
        If SetupWizard.AddressBarStyle = "BlueAddressBar" Then
            FileIO.FileSystem.RenameFile(WinDir & "\Resources\Themes\aerorp\aero_blue.msstyles", "aerolite.msstyles")
            FileIO.FileSystem.DeleteFile(WinDir & "\Resources\Themes\aerorp\aero_nonblue.msstyles")
        ElseIf SetupWizard.AddressBarStyle = "WhiteAddressBar" Then
            FileIO.FileSystem.RenameFile(WinDir & "\Resources\Themes\aerorp\aero_nonblue.msstyles", "aerolite.msstyles")
            FileIO.FileSystem.DeleteFile(WinDir & "\Resources\Themes\aerorp\aero_blue.msstyles")
        End If
        DetailsTextBox.AppendText(" Finished!" & vbCrLf)
        '' Copy wallpaper files
        DetailsTextBox.AppendText("Copying wallpaper files to " & WinDir & "\Web\Wallpaper\Windows 8 Release Preview...")
        FileIO.FileSystem.CopyDirectory(ResFolder & "\Theme\Wallpaper\Windows 8 Release Preview", WinDir & "\Web\Wallpaper\Windows 8 Release Preview", True)
        ProgressBarSetup.PerformStep()
        ProgressBarSetup.Refresh()
        ActionLabel.Text = "Theme has been installed!"
        ActionLabel.Refresh()
        DetailsTextBox.AppendText(" Finished!" & vbCrLf)

        ' Installing sounds
        If SetupWizard.InstSounds = True Then
            ActionLabel.Text = "Installing sound scheme..."
            ActionLabel.Refresh()
            DetailsTextBox.AppendText(ActionLabel.Text & vbCrLf)
            DetailsTextBox.AppendText("Copying sound files to " & WinDir & "\Media\Windows 8 Release Preview...")
            FileIO.FileSystem.CopyDirectory(ResFolder & "\Sounds", WinDir & "\Media\Windows 8 Release Preview", True)
            ProgressBarSetup.PerformStep()
            ProgressBarSetup.Refresh()
            ActionLabel.Text = "Sound scheme has been installed!"
            ActionLabel.Refresh()
            DetailsTextBox.AppendText(" Finished!" & vbCrLf)
        ElseIf SetupWizard.InstSounds = False Then
            DetailsTextBox.AppendText("Removing file " & WinDir & "\Resources\Themes\aerorp_sounds.theme...")
            FileIO.FileSystem.DeleteFile(WinDir & "\Resources\Themes\aerorp_sounds.theme")
            DetailsTextBox.AppendText(" Finished!" & vbCrLf)
        End If

        ' Installing Gadgets
        If SetupWizard.InstGadgets = True Then
            ActionLabel.Text = "Installing 8GadgetPack..."
            ActionLabel.Refresh()
            DetailsTextBox.AppendText(ActionLabel.Text)
            BgWorker.RunWorkerAsync("InstGadgets")
            WaitUntilTaskFinishes()
            ProgressBarSetup.PerformStep()
            ProgressBarSetup.Refresh()
            ActionLabel.Text = "8GadgetPack has been installed!"
            ActionLabel.Refresh()
            DetailsTextBox.AppendText(" Finished!" & vbCrLf)
        End If

        ' Installing 7+ Taskbar Tweaker
        If SetupWizard.Inst7TaskTw = True Then
            ActionLabel.Text = "Installing 7+ Taskbar Tweaker..."
            ActionLabel.Refresh()
            DetailsTextBox.AppendText(ActionLabel.Text)
            BgWorker.RunWorkerAsync("Inst7TaskTw")
            WaitUntilTaskFinishes()
            ProgressBarSetup.PerformStep()
            ProgressBarSetup.Refresh()
            ActionLabel.Text = "7+ Taskbar Tweaker has been installed!"
            ActionLabel.Refresh()
            DetailsTextBox.AppendText(" Finished!" & vbCrLf)
        End If

        ' Finalising installation
        If SetupWizard.ProgLanguage = "en-GB" Then
            ActionLabel.Text = "Finalising installation..."
        ElseIf SetupWizard.ProgLanguage = "en-US" Then
            ActionLabel.Text = "Finalizing installation..."
        End If
        ActionLabel.Refresh()
        DetailsTextBox.AppendText(ActionLabel.Text & vbCrLf)
        ' Create registry keys, copy files
        DetailsTextBox.AppendText("Copying file firstrun.xht to " & ResFolder & "...")
        '' Copy firstrun.xht from resources and put files in program files
        If SetupWizard.InstQuero = True Then
            File.WriteAllBytes(SetupWizard.ProgDir & "\firstrun.xht", My.Resources.ResourceManager.GetObject("FirstRunQuero_" & SetupWizard.LanguageForLocalisedStrings))
        ElseIf SetupWizard.InstQuero = False Then
            File.WriteAllBytes(SetupWizard.ProgDir & "\firstrun.xht", My.Resources.ResourceManager.GetObject("FirstRun_" & SetupWizard.LanguageForLocalisedStrings))
        End If
        If FileIO.FileSystem.DirectoryExists(ResFolder & "\firstrun") Then
            FileIO.FileSystem.MoveDirectory(ResFolder & "\firstrun", SetupWizard.ProgDir & "\firstrun")
        End If
        DetailsTextBox.AppendText(" Finished!" & vbCrLf)
        '' Create config file in <ProgDir>\config.xml
        DetailsTextBox.AppendText("Creating config file in path " & SetupWizard.ProgDir & "\config.xml...")
        FileIO.FileSystem.CopyFile(ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.PerUserRoamingAndLocal).FilePath, SetupWizard.ProgDir & "\config.xml")
        DetailsTextBox.AppendText(" Finished!" & vbCrLf)
        '' Create Control Panel\Programs entry
        '' EstimatedSize is theme & wallpaper without programs -> about 18,2 MB
        DetailsTextBox.AppendText("Creating Control Panel entry...")
        Using RegKeyControlPanel As RegistryKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall", True)
            RegKeyControlPanel.CreateSubKey("Windows 8.1 to 8 RP Converter")
            Using RegKeyControlPanelEntry As RegistryKey = RegKeyControlPanel.OpenSubKey("Windows 8.1 to 8 RP Converter", True)
                With RegKeyControlPanelEntry
                    .SetValue("DisplayIcon", ResFolder & "\program.ico", RegistryValueKind.String)
                    .SetValue("DisplayName", "Windows 8.1 to Windows 8 RP Converter", RegistryValueKind.String)
                    .SetValue("DisplayVersion", "1.0 beta", RegistryValueKind.String)
                    .SetValue("EstimatedSize", 18572, RegistryValueKind.DWord)
                    .SetValue("InstallLocation", SetupWizard.ProgDir, RegistryValueKind.String)
                    .SetValue("ModifyPath", """" & SetupWizard.ProgDir & "\modify.exe""", RegistryValueKind.String)
                    .SetValue("NoModify", 0, RegistryValueKind.DWord)
                    .SetValue("NoRemove", 0, RegistryValueKind.DWord)
                    .SetValue("NoRepair", 1, RegistryValueKind.DWord)
                    .SetValue("Publisher", "Moha Logiciels", RegistryValueKind.String)
                    .SetValue("UninstallString", """" & SetupWizard.ProgDir & "\uninstall.exe""", RegistryValueKind.String)
                    .SetValue("Version", "1.0.0.0", RegistryValueKind.String)
                    .Close()
                End With
                RegKeyControlPanel.Close()
            End Using
        End Using
        DetailsTextBox.AppendText(" Finished!" & vbCrLf)
        '' Copy modify.exe, uninstall.exe and firstrun.exe to program folder
        DetailsTextBox.AppendText("Copying modify.exe, uninstall.exe and firstrun.exe to " & SetupWizard.ProgDir & "...")
        FileIO.FileSystem.WriteAllBytes(SetupWizard.ProgDir & "\modify.exe", My.Resources.ModifyProgram, False)
        FileIO.FileSystem.WriteAllBytes(SetupWizard.ProgDir & "\uninstall.exe", My.Resources.UninstallProgram, False)
        FileIO.FileSystem.WriteAllBytes(SetupWizard.ProgDir & "\firstrun.exe", My.Resources.AutostartAfterRestartLauncher, False)
        DetailsTextBox.AppendText(" Finished!" & vbCrLf)
        ProgressBarSetup.PerformStep()
        ProgressBarSetup.Refresh()
        '' Set autostart reg key for firstrun.exe
        DetailsTextBox.AppendText("Creating autostart entry...")
        If Registry.CurrentUser.OpenSubKey("Software\Microsoft\Windows\CurrentVersion\RunOnce", False) Is Nothing Then
            Using CurrentVersionRegKey As RegistryKey = Registry.CurrentUser.OpenSubKey("Software\Microsoft\Windows\CurrentVersion", True)
                CurrentVersionRegKey.CreateSubKey("RunOnce")
                CurrentVersionRegKey.Close()
            End Using
        End If
        Using AutorunRegKey As RegistryKey = Registry.CurrentUser.OpenSubKey("Software\Microsoft\Windows\CurrentVersion\RunOnce", True)
            AutorunRegKey.SetValue("Run firstrun.exe", """" & SetupWizard.ProgDir & "\firstrun.exe"" /setup", RegistryValueKind.String)
            If SetupWizard.InstSounds = True Then
                AutorunRegKey.SetValue("Run firstrun.exe", AutorunRegKey.GetValue("Run firstrun.exe") & " /sounds", RegistryValueKind.String)
            End If
            AutorunRegKey.Close()
        End Using
        DetailsTextBox.AppendText(" Finished!" & vbCrLf)
        '' Remove temp files and folder
        ActionLabel.Text = "Removing temp files..."
        ActionLabel.Refresh()
        DetailsTextBox.AppendText(ActionLabel.Text)
        Dim LocalAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData)
        If FileIO.FileSystem.DirectoryExists(LocalAppData & "\SetupWindows8RPConv") Then
            FileIO.FileSystem.DeleteDirectory(LocalAppData & "\SetupWindows8RPConv", FileIO.DeleteDirectoryOption.DeleteAllContents)
            DetailsTextBox.AppendText(" Finished!" & vbCrLf)
        Else
            MessageBox.Show("Temporary files could not be found!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            DetailsTextBox.AppendText(" Failed!" & vbCrLf)
        End If
        ProgressBarSetup.PerformStep()
        ProgressBarSetup.Refresh()

        ' Finish installation
        ActionLabel.Text = "Installation has been successfully completed!"
        ActionLabel.Refresh()
        DetailsTextBox.AppendText(ActionLabel.Text & vbCrLf)
        '' Change Next button to Finish
        NextButton.Text = "&Finish"
        NextButton.Enabled = True
        CloseProgram = True
        ShowMessageBoxForRestartOnExit = True
    End Sub

    Private Sub ShowDetailsButton_Click(sender As Object, e As EventArgs) Handles ShowDetailsButton.Click
        DetailsTextBox.Visible = True
        ShowDetailsButton.Enabled = False
    End Sub

    Private Sub NextButton_Click(sender As Object, e As EventArgs) Handles NextButton.Click
        If CloseProgram = True Then
            Me.Close()
        Else
            MessageBox.Show("The setup cannot be finished right now.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End If
    End Sub

    Private Sub CloseButton_Click(sender As Object, e As EventArgs) Handles CloseButton.Click
        If CloseProgram = True Then
            Me.Close()
        Else
            MessageBox.Show("The setup cannot be finished right now.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End If
    End Sub

    Private Sub SetupInstall_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        If CloseProgram = False Then
            MessageBox.Show("The setup cannot be finished right now.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        ElseIf CloseProgram = True Then
            If e.CloseReason = CloseReason.UserClosing Then
                If ShowMessageBoxForRestartOnExit = True Then
                    ' Save log file
                    If MessageBox.Show("Do you want to save the details pane to a log file in the program’s directory?", "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = System.Windows.Forms.DialogResult.Yes Then
                        CreateLogFile()
                    End If
                    ' Restart prompt -> wait 3 seconds before restarting
                    If MessageBox.Show("You need to restart your system so that the changes take effect. Do you want to restart now?", "Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) = System.Windows.Forms.DialogResult.Yes Then
                        Using ShutDownProcess As New Process
                            With ShutDownProcess
                                .StartInfo.FileName = WinDir & "\System32\shutdown.exe"
                                .StartInfo.Arguments = "/r /t 3"
                                .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                                .Start()
                            End With
                        End Using
                        SetupWizard.Close()
                    ElseIf System.Windows.Forms.DialogResult.No Then
                        Application.Exit()
                    End If
                Else
                    Application.Exit()
                End If
            End If
        End If
    End Sub
End Class