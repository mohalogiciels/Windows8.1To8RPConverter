Imports System.IO
Imports System.IO.Compression
Imports System.ComponentModel
Imports System.Configuration
Imports Microsoft.Win32

Public Class setup_install_en
    Dim ProgramFiles = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles)
    Dim WinDir = Environment.GetFolderPath(Environment.SpecialFolder.Windows)
    Dim WithEvents BgWorker As New BackgroundWorker
    Dim IsFinished As Boolean = True
    Dim ShowMessageBoxForRestartOnExit As Boolean = False

    Private Sub Wait(ByVal seconds As Integer)
        For Index As Integer = 0 To seconds * 100
            System.Threading.Thread.Sleep(10)
            Application.DoEvents()
        Next
    End Sub

    Private Sub setup_install_en_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DetailsTextBox.ScrollBars = ScrollBars.Vertical
        DetailsTextBox.AppendText(ActionLabel.Text)
    End Sub

    Private Sub setup_install_en_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown
        DetailsTextBox.AppendText(" Finished!" & vbCrLf)

        ' Create system restore point
        ActionLabel.Text = "Creating a system restore point..."
        ActionLabel.Refresh()
        DetailsTextBox.AppendText(ActionLabel.Text)
        ProgressBarSetup.Style = ProgressBarStyle.Marquee
        ProgressBarSetup.MarqueeAnimationSpeed = 100
        BgWorker.WorkerSupportsCancellation = True
        BgWorker.RunWorkerAsync()
    End Sub

    Private Sub BgWorker_DoWork(sender As Object, e As DoWorkEventArgs) Handles BgWorker.DoWork
        ' Creating system restore point
        Dim RestPointCreate As New Process
        With RestPointCreate
            .StartInfo.FileName = WinDir & "\System32\WindowsPowerShell\v1.0\powershell.exe"
            .StartInfo.Arguments = "Checkpoint-Computer -Description 'Installation Windows 8.1 to 8RP Converter'"
            .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
            .Start()
            .WaitForExit()
        End With
        BgWorker.CancelAsync()
    End Sub

    Private Sub BgWorker_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles BgWorker.RunWorkerCompleted
        ' Change progress bar
        With ProgressBarSetup
            .Style = ProgressBarStyle.Continuous
            .Value = 0
            .Maximum = 10
            .PerformStep()
            .Refresh()
        End With
        ActionLabel.Text = "The system restore point has been successfully created!"
        ActionLabel.Refresh()
        DetailsTextBox.AppendText(" Finished!" & vbCrLf)
        CloseButton.Enabled = False
        IsFinished = False

        ' Waiting 2 seconds before starting installation
        Wait(2)

        ' If error occurs, create log file
        Dim ProgDir As String = ProgramFiles & "\Windows8.1to8RPConv"
        Try
            SetupProgram()
        Catch ex As Exception
            Dim LocalAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData)
            MessageBox.Show("An error occured during installation: " & ex.Message & " Please check the log file created on the desktop to check what happened, and try it again.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Stop)
            DetailsTextBox.AppendText(vbCrLf & "ERROR: " & ex.Message & vbCrLf)
            ' \Enable close button
            IsFinished = True
            CloseButton.Enabled = True
            CloseButton.Text = "&Close"
            ' \Create log file
            File.AppendAllText(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) & "\win8rpconv_install.log", "Installation date: " & My.Computer.Clock.LocalTime.ToShortDateString & vbCrLf)
            File.AppendAllText(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) & "\win8rpconv_install.log", "Setup options: [Theme: """ & setup.AddressBarStyle & """, InstONE: """ & setup.InstONE & """, InstUX: """ & setup.InstUX & """, InstQuero: """ & setup.InstQuero & """, InstSysFiles: """ & setup.InstSysFiles & """, InstSounds: """ & setup.InstSounds & """]" & vbCrLf & "Install log:" & vbCrLf)
            File.AppendAllLines(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) & "\win8rpconv_install.log", DetailsTextBox.Lines)
            ' \Delete ProgramFiles directory
            If FileIO.FileSystem.DirectoryExists(ProgDir) Then
                FileIO.FileSystem.DeleteDirectory(ProgDir, FileIO.DeleteDirectoryOption.DeleteAllContents)
            End If
            ' \Close program
            setup.Close()
        End Try
    End Sub

    Private Sub SetupProgram()
        ' Initialise progress bar
        Dim ProgBarCount As Integer = 50
        If setup.InstAeroGlass = True Then
            ProgBarCount += 10
        End If
        If setup.InstONE = True Then
            ProgBarCount += 10
        End If
        If setup.InstQuero = True Then
            ProgBarCount += 10
        End If
        If setup.InstSysFiles = True Then
            ProgBarCount += 10
        End If
        If setup.InstSounds = True Then
            ProgBarCount += 10
        End If
        If setup.InstGadgets = True Then
            ProgBarCount += 10
        End If

        ' Change progress bar to continuous
        With ProgressBarSetup
            .Refresh()
            .Style = ProgressBarStyle.Continuous
            .Value = 0
            .Maximum = ProgBarCount
        End With

        ' Extract resources file
        ActionLabel.Text = "Extracting resources file..."
        ActionLabel.Refresh()
        DetailsTextBox.AppendText(ActionLabel.Text)
        Dim ProgDir As String = ProgramFiles & "\Windows8.1to8RPConv"
        Dim ResFolder As String = ProgDir & "\res"
        ' \Read and extract resources file from program resources
        Dim ResZipFileBytes As Byte() = My.Resources.ResourceFile
        Using ZipFileToOpen As New MemoryStream(ResZipFileBytes)
            Using ResZipFile As New ZipArchive(ZipFileToOpen, ZipArchiveMode.Update)
                ResZipFile.ExtractToDirectory(ResFolder)
            End Using
        End Using
        ' \Check if file has been extracted -> folder has been created
        If FileIO.FileSystem.DirectoryExists(ResFolder) Then
            ProgressBarSetup.PerformStep()
            ProgressBarSetup.Refresh()
            ActionLabel.Text = "Resources file has been extracted!"
            ActionLabel.Refresh()
            DetailsTextBox.AppendText(" Finished!" & vbCrLf)
        Else
            MessageBox.Show("Resources file could not be extracted!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Stop)
            setup.Close()
        End If

        ' Install Aero Glass
        If setup.AeroGlassInstalled = False Or (setup.AeroGlassInstalled = True And setup.InstAeroGlass = True) Then
            Dim AeroGlassSetup As New Process
            ActionLabel.Text = "Installing Aero Glass..."
            ActionLabel.Refresh()
            DetailsTextBox.AppendText(ActionLabel.Text)
            ' \Change install directory to user’s root drive
            Dim AeroGlassSetupInfFilePath As String = ResFolder & "\AeroGlass\settings.inf"
            Dim AeroGlassSetupInfFile As String() = File.ReadAllLines(AeroGlassSetupInfFilePath)
            AeroGlassSetupInfFile(2) = "Dir=" & Environ("SystemDrive") & "\AeroGlass"
            File.WriteAllLines(AeroGlassSetupInfFilePath, AeroGlassSetupInfFile)
            ' \Start Aero Glass setup
            With AeroGlassSetup
                .StartInfo.FileName = ResFolder & "\AeroGlass\setup-w8.1-1.4.6.exe"
                .StartInfo.Arguments = "/silent /loadinf=""" & AeroGlassSetupInfFilePath & """"
                .Start()
                .WaitForExit()
            End With
            Dim RegKeyAero As RegistryKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\Microsoft\Windows\DWM", True)
            RegKeyAero.SetValue("CustomThemeAtlas", "")
            RegKeyAero.Close()
            ProgressBarSetup.PerformStep()
            ProgressBarSetup.Refresh()
            ActionLabel.Text = "Aero Glass has been installed!"
            ActionLabel.Refresh()
            DetailsTextBox.AppendText(" Finished!" & vbCrLf)
        End If

        ' Install OldNewExplorer
        Dim SourcePathONE As String = ResFolder & "\OldNewExplorer"
        Dim DestPathONE As String = ProgramFiles & "\OldNewExplorer"
        If setup.InstONE = True Then
            Dim ONESetup As New Process
            ActionLabel.Text = "Installing OldNewExplorer..."
            ActionLabel.Refresh()
            DetailsTextBox.AppendText(ActionLabel.Text & vbCrLf)
            FileIO.FileSystem.CopyDirectory(SourcePathONE, DestPathONE)
            If setup.BitnessSystem = "64bit" Then
                DetailsTextBox.AppendText("Registering file " & DestPathONE & "\OldNewExplorer64.dll...")
                With ONESetup
                    .StartInfo.FileName = WinDir & "\System32\regsvr32.exe"
                    .StartInfo.Arguments = "/s """ & DestPathONE & "\OldNewExplorer64.dll"""
                    .Start()
                    .WaitForExit()
                End With
                DetailsTextBox.AppendText(" Finished!" & vbCrLf)
                DetailsTextBox.AppendText("Registering file " & DestPathONE & "\OldNewExplorer32.dll...")
                With ONESetup
                    .StartInfo.FileName = WinDir & "\System32\regsvr32.exe"
                    .StartInfo.Arguments = "/s """ & DestPathONE & "\OldNewExplorer32.dll"""
                    .Start()
                    .WaitForExit()
                End With
                DetailsTextBox.AppendText(" Finished!" & vbCrLf)
            ElseIf setup.BitnessSystem = "32bit" Then
                DetailsTextBox.AppendText("Registering file " & DestPathONE & "\OldNewExplorer32.dll...")
                With ONESetup
                    .StartInfo.FileName = WinDir & "\System32\regsvr32.exe"
                    .StartInfo.Arguments = "/s """ & DestPathONE & "\OldNewExplorer32.dll"""
                    .Start()
                    .WaitForExit()
                End With
                DetailsTextBox.AppendText(" Finished!" & vbCrLf)
            End If
            ' \Create shortcut on desktop
            DetailsTextBox.AppendText("Creating shortcut on desktop...")
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
                DetailsTextBox.AppendText(" Finished!" & vbCrLf)
            Catch ex As Exception
                MessageBox.Show("Shortcut to OldNewExplorer configuration program could not be created!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                DetailsTextBox.AppendText(" Failed!" & vbCrLf)
            Finally
                Application.DoEvents()
            End Try
            ProgressBarSetup.PerformStep()
            ProgressBarSetup.Refresh()
            ActionLabel.Text = "OldNewExplorer has been installed!"
            ActionLabel.Refresh()
        End If

        ' Install UXTheme patcher
        Dim SourcePathUX As String = ResFolder & "\UXTheme"
        Dim SourcePathUXTSB As String = ResFolder & "\UXThemeDLL"
        If setup.InstUX = "UltraUX" Then
            ' \Install UltraUX
            Dim UltraUXSetup As New Process
            ActionLabel.Text = "Installing UltraUXThemePatcher..."
            ActionLabel.Refresh()
            DetailsTextBox.AppendText(ActionLabel.Text)
            With UltraUXSetup
                .StartInfo.FileName = ResFolder & "\UXTheme\UltraUXThemePatcher_4.4.4.exe"
                .Start()
                .WaitForExit()
            End With
            ProgressBarSetup.PerformStep()
            ProgressBarSetup.Refresh()
            ActionLabel.Text = "UltraUXThemePatcher has been installed!"
            ActionLabel.Refresh()
            DetailsTextBox.AppendText(" Finished!" & vbCrLf)
        ElseIf setup.InstUX = "UXTSB" Then
            ' \Install UXTSB
            ActionLabel.Text = "Installing UXThemeSignatureBypass..."
            ActionLabel.Refresh()
            DetailsTextBox.AppendText(ActionLabel.Text & vbCrLf)
            If setup.BitnessSystem = "64bit" Then
                DetailsTextBox.AppendText("Copying file UxThemeSignatureBypass32.dll to " & WinDir & "\UxThemeSignatureBypass32.dll...")
                DetailsTextBox.AppendText(" Finished!" & vbCrLf)
                DetailsTextBox.AppendText("Copying file UxThemeSignatureBypass64.dll to " & WinDir & "\UxThemeSignatureBypass64.dll...")
                DetailsTextBox.AppendText(" Finished!" & vbCrLf)
                FileIO.FileSystem.CopyDirectory(SourcePathUXTSB, WinDir)
            ElseIf setup.BitnessSystem = "32bit" Then
                FileIO.FileSystem.CopyFile(SourcePathUXTSB & "\UxThemeSignatureBypass32.dll", WinDir & "\UxThemeSignatureBypass32.dll")
                DetailsTextBox.AppendText("Copying file UxThemeSignatureBypass32.dll to " & WinDir & "\UxThemeSignatureBypass32.dll...")
                DetailsTextBox.AppendText(" Finished!" & vbCrLf)
            End If
            ' \UXTSB\Registry
            Dim RegKey As RegistryKey
            Dim RegKey32 As RegistryKey
            DetailsTextBox.AppendText("Registring files...")
            If setup.BitnessSystem = "64bit" Then
                RegKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\Microsoft\Windows NT\CurrentVersion\Windows", True)
                If RegKey.GetValue("AppInit_DLLs") <> "" Then
                    RegKey.SetValue("AppInit_DLLs", RegKey.GetValue("AppInit_DLLs") & ", " & WinDir & "\UxThemeSignatureBypass64.dll")
                Else
                    RegKey.SetValue("AppInit_DLLs", WinDir & "\UxThemeSignatureBypass64.dll")
                End If
                If RegKey.GetValue("LoadAppInit_DLLs") <> 1 Then
                    RegKey.SetValue("LoadAppInit_DLLs", 1)
                End If
                RegKey.Close()
                ' \UXTSB\Registry\SysWOW64
                RegKey32 = Registry.LocalMachine.OpenSubKey("SOFTWARE\Wow6432Node\Microsoft\Windows NT\CurrentVersion\Windows", True)
                If RegKey32.GetValue("AppInit_DLLs") <> "" Then
                    RegKey32.SetValue("AppInit_DLLs", RegKey.GetValue("AppInit_DLLs") & ", " & WinDir & "\UxThemeSignatureBypass32.dll")
                Else
                    RegKey32.SetValue("AppInit_DLLs", WinDir & "\UxThemeSignatureBypass32.dll")
                End If
                If RegKey32.GetValue("LoadAppInit_DLLs") <> 1 Then
                    RegKey32.SetValue("LoadAppInit_DLLs", 1)
                End If
                RegKey32.Close()
            ElseIf setup.BitnessSystem = "32bit" Then
                RegKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\Microsoft\Windows NT\CurrentVersion\Windows", True)
                If RegKey.GetValue("AppInit_DLLs") <> "" Then
                    RegKey.SetValue("AppInit_DLLs", RegKey.GetValue("AppInit_DLLs") & ", " & WinDir & "\UxThemeSignatureBypass32.dll")
                Else
                    RegKey.SetValue("AppInit_DLLs", WinDir & "\UxThemeSignatureBypass32.dll")
                End If
                If RegKey.GetValue("LoadAppInit_DLLs") <> 1 Then
                    RegKey.SetValue("LoadAppInit_DLLs", 1)
                End If
                RegKey.Close()
            End If
            ProgressBarSetup.PerformStep()
            ProgressBarSetup.Refresh()
            ActionLabel.Text = "UXThemeSignatureBypass has been installed!"
            ActionLabel.Refresh()
            DetailsTextBox.AppendText(" Finished!" & vbCrLf)
        End If

        ' Quero Toolbar
        If setup.InstQuero = True Then
            Dim QueroSetup As New Process
            ActionLabel.Text = "Installing Quero Toolbar..."
            ActionLabel.Refresh()
            DetailsTextBox.AppendText(ActionLabel.Text)
            ' \Change install directory to user’s root drive
            Dim QueroSetupInfFilePath As String = ResFolder & "\Quero\settings.inf"
            Dim QueroSetupInfFile As String() = File.ReadAllLines(QueroSetupInfFilePath)
            QueroSetupInfFile(2) = "Dir=" & ProgramFiles & "\Quero Toolbar"
            File.WriteAllLines(QueroSetupInfFilePath, QueroSetupInfFile)
            ' \Start Quero Toolbar setup
            If setup.BitnessSystem = "64bit" Then
                With QueroSetup
                    .StartInfo.FileName = ResFolder & "\Quero\QueroToolbarInstaller_x64.exe"
                    .StartInfo.Arguments = "/silent /loadinf=""" & QueroSetupInfFilePath & """"
                    .Start()
                    .WaitForExit()
                End With
            ElseIf setup.BitnessSystem = "32bit" Then
                With QueroSetup
                    .StartInfo.FileName = ResFolder & "\Quero\QueroToolbarInstaller_x86.exe"
                    .StartInfo.Arguments = "/silent /loadinf=""" & QueroSetupInfFilePath & """"
                    .Start()
                    .WaitForExit()
                End With
            End If
            ProgressBarSetup.PerformStep()
            ProgressBarSetup.Refresh()
            ActionLabel.Text = "Quero Toolbar has been installed!"
            ActionLabel.Refresh()
            DetailsTextBox.AppendText(" Finished!" & vbCrLf)
        End If

        ' UIRibbon.dll, UIRibbonRes.dll
        Dim SourcePathSysFiles As String = ResFolder & "\UIRibbon"
        If setup.InstSysFiles = True Then
            Dim SysFilesSetup As New Process
            ActionLabel.Text = "Replacing UIRibbon.dll and UIRibbonRes.dll..."
            ActionLabel.Refresh()
            DetailsTextBox.AppendText(ActionLabel.Text & vbCrLf)
            ' \Remove files not for the selected language from Resource folder
            Dim System32Files = FileIO.FileSystem.GetDirectories(ResFolder & "\UIRibbon\system32")
            For Each LangFolderPath In System32Files
                If LangFolderPath <> ResFolder & "\UIRibbon\system32\" & setup.Language Then
                    FileIO.FileSystem.DeleteDirectory(LangFolderPath, FileIO.DeleteDirectoryOption.DeleteAllContents)
                End If
            Next
            Dim SysWOW64Files = FileIO.FileSystem.GetDirectories(ResFolder & "\UIRibbon\syswow64")
            For Each LangFolderPath In SysWOW64Files
                If LangFolderPath <> ResFolder & "\UIRibbon\syswow64\" & setup.Language Then
                    FileIO.FileSystem.DeleteDirectory(LangFolderPath, FileIO.DeleteDirectoryOption.DeleteAllContents)
                End If
            Next
            ' \End explorer.exe task
            Dim EndExplorerTask As New Process
            With EndExplorerTask
                .StartInfo.FileName = WinDir & "\System32\taskkill.exe"
                .StartInfo.Arguments = "/f /im explorer.exe"
                .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                .Start()
                .WaitForExit()
            End With
            ' \Replace files for 64-bit
            If setup.BitnessSystem = "64bit" Then
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
                DetailsTextBox.AppendText("Creating backup of " & WinDir & "\System32\UIRibbon.dll...")
                FileIO.FileSystem.RenameFile(WinDir & "\System32\UIRibbon.dll", "UIRibbon.dll.bak")
                DetailsTextBox.AppendText(" Finished!" & vbCrLf)
                DetailsTextBox.AppendText("Creating backup of " & WinDir & "\System32\UIRibbonRes.dll...")
                FileIO.FileSystem.RenameFile(WinDir & "\System32\UIRibbonRes.dll", "UIRibbonRes.dll.bak")
                DetailsTextBox.AppendText(" Finished!" & vbCrLf)
                With SysFilesSetup
                    .StartInfo.FileName = WinDir & "\System32\takeown.exe"
                    .StartInfo.Arguments = "/f " & WinDir & "\System32\" & setup.Language & "\UIRibbon.dll.mui /a"
                    .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                    .Start()
                    .WaitForExit()
                End With
                With SysFilesSetup
                    .StartInfo.FileName = WinDir & "\System32\icacls.exe"
                    .StartInfo.Arguments = WinDir & "\System32\" & setup.Language & "\UIRibbon.dll.mui /grant *S-1-5-32-545:F *S-1-5-32-544:F"
                    .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                    .Start()
                    .WaitForExit()
                End With
                DetailsTextBox.AppendText("Creating backup of " & WinDir & "\System32\" & setup.Language & "\UIRibbon.dll.mui...")
                FileIO.FileSystem.RenameFile(WinDir & "\System32\" & setup.Language & "\UIRibbon.dll.mui", "UIRibbon.dll.mui.bak")
                DetailsTextBox.AppendText(" Finished!" & vbCrLf)
                DetailsTextBox.AppendText("Replacing files...")
                FileIO.FileSystem.CopyDirectory(SourcePathSysFiles & "\system32", WinDir & "\System32")
                DetailsTextBox.AppendText(" Finished!" & vbCrLf)
                ' \New files change to TrustedInstaller
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
                    .StartInfo.Arguments = WinDir & "\System32\" & setup.Language & "\UIRibbon.dll.mui /setowner ""NT SERVICE\TrustedInstaller"""
                    .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                    .Start()
                    .WaitForExit()
                End With
                With SysFilesSetup
                    .StartInfo.FileName = WinDir & "\System32\icacls.exe"
                    .StartInfo.Arguments = WinDir & "\System32\" & setup.Language & "\UIRibbon.dll.mui /reset"
                    .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                    .Start()
                    .WaitForExit()
                End With
                With SysFilesSetup
                    .StartInfo.FileName = WinDir & "\System32\icacls.exe"
                    .StartInfo.Arguments = WinDir & "\System32\" & setup.Language & "\UIRibbon.dll.mui /grant *S-1-5-18:RX *S-1-5-32-544:RX *S-1-5-32-545:RX *S-1-15-2-1:RX ""NT SERVICE\TrustedInstaller"":F /inheritance:r"
                    .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                    .Start()
                    .WaitForExit()
                End With
                ' \UIRibbon.dll, UIRibbonRes.dll\SysWOW64
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
                DetailsTextBox.AppendText("Creating backup of " & WinDir & "\SysWOW64\UIRibbon.dll...")
                FileIO.FileSystem.RenameFile(WinDir & "\SysWOW64\UIRibbon.dll", "UIRibbon.dll.bak")
                DetailsTextBox.AppendText(" Finished!" & vbCrLf)
                DetailsTextBox.AppendText("Creating backup of " & WinDir & "\SysWOW64\UIRibbonRes.dll...")
                FileIO.FileSystem.RenameFile(WinDir & "\SysWOW64\UIRibbonRes.dll", "UIRibbonRes.dll.bak")
                DetailsTextBox.AppendText(" Finished!" & vbCrLf)
                With SysFilesSetup
                    .StartInfo.FileName = WinDir & "\System32\takeown.exe"
                    .StartInfo.Arguments = "/f " & WinDir & "\SysWOW64\" & setup.Language & "\UIRibbon.dll.mui /a"
                    .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                    .Start()
                    .WaitForExit()
                End With
                With SysFilesSetup
                    .StartInfo.FileName = WinDir & "\System32\icacls.exe"
                    .StartInfo.Arguments = WinDir & "\SysWOW64\" & setup.Language & "\UIRibbon.dll.mui /grant *S-1-5-32-545:F *S-1-5-32-544:F"
                    .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                    .Start()
                    .WaitForExit()
                End With
                DetailsTextBox.AppendText("Creating backup of " & WinDir & "\SysWOW64\" & setup.Language & "\UIRibbon.dll.mui...")
                FileIO.FileSystem.RenameFile(WinDir & "\SysWOW64\" & setup.Language & "\UIRibbon.dll.mui", "UIRibbon.dll.mui.bak")
                DetailsTextBox.AppendText(" Finished!" & vbCrLf)
                DetailsTextBox.AppendText("Replacing files...")
                FileIO.FileSystem.CopyDirectory(SourcePathSysFiles & "\syswow64", WinDir & "\SysWOW64")
                DetailsTextBox.AppendText(" Finished!" & vbCrLf)
                ' \New files\SysWOW64 change to TrustedInstaller
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
                    .StartInfo.Arguments = WinDir & "\SysWOW64\" & setup.Language & "\UIRibbon.dll.mui /setowner ""NT SERVICE\TrustedInstaller"""
                    .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                    .Start()
                    .WaitForExit()
                End With
                With SysFilesSetup
                    .StartInfo.FileName = WinDir & "\System32\icacls.exe"
                    .StartInfo.Arguments = WinDir & "\SysWOW64\" & setup.Language & "\UIRibbon.dll.mui /reset"
                    .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                    .Start()
                    .WaitForExit()
                End With
                With SysFilesSetup
                    .StartInfo.FileName = WinDir & "\System32\icacls.exe"
                    .StartInfo.Arguments = WinDir & "\SysWOW64\" & setup.Language & "\UIRibbon.dll.mui /grant *S-1-5-18:RX *S-1-5-32-544:RX *S-1-5-32-545:RX *S-1-15-2-1:RX ""NT SERVICE\TrustedInstaller"":F /inheritance:r"
                    .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                    .Start()
                    .WaitForExit()
                End With
            ElseIf setup.BitnessSystem = "32bit" Then
                ' \Replace files for 32-bit
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
                DetailsTextBox.AppendText("Creating backup of " & WinDir & "\System32\UIRibbon.dll...")
                FileIO.FileSystem.RenameFile(WinDir & "\System32\UIRibbon.dll", "UIRibbon.dll.bak")
                DetailsTextBox.AppendText(" Finished!" & vbCrLf)
                DetailsTextBox.AppendText("Creating backup of " & WinDir & "\System32\UIRibbonRes.dll...")
                FileIO.FileSystem.RenameFile(WinDir & "\System32\UIRibbonRes.dll", "UIRibbonRes.dll.bak")
                DetailsTextBox.AppendText(" Finished!" & vbCrLf)
                With SysFilesSetup
                    .StartInfo.FileName = WinDir & "\System32\takeown.exe"
                    .StartInfo.Arguments = "/f " & WinDir & "\System32\" & setup.Language & "\UIRibbon.dll.mui /a"
                    .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                    .Start()
                    .WaitForExit()
                End With
                With SysFilesSetup
                    .StartInfo.FileName = WinDir & "\System32\icacls.exe"
                    .StartInfo.Arguments = WinDir & "\System32\" & setup.Language & "\UIRibbon.dll.mui /grant *S-1-5-32-545:F *S-1-5-32-544:F"
                    .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                    .Start()
                    .WaitForExit()
                End With
                DetailsTextBox.AppendText("Creating backup of " & WinDir & "\System32\" & setup.Language & "\UIRibbonRes.dll...")
                FileIO.FileSystem.RenameFile(WinDir & "\System32\" & setup.Language & "\UIRibbon.dll.mui", "UIRibbon.dll.mui.bak")
                DetailsTextBox.AppendText(" Finished!" & vbCrLf)
                DetailsTextBox.AppendText("Replacing files...")
                FileIO.FileSystem.CopyDirectory(SourcePathSysFiles & "\syswow64", WinDir & "\System32")
                DetailsTextBox.AppendText(" Finished!" & vbCrLf)
                ' \New files change to TrustedInstaller
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
                    .StartInfo.Arguments = WinDir & "\System32\" & setup.Language & "\UIRibbon.dll.mui /setowner ""NT SERVICE\TrustedInstaller"""
                    .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                    .Start()
                    .WaitForExit()
                End With
                With SysFilesSetup
                    .StartInfo.FileName = WinDir & "\System32\icacls.exe"
                    .StartInfo.Arguments = WinDir & "\System32\" & setup.Language & "\UIRibbon.dll.mui /reset"
                    .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                    .Start()
                    .WaitForExit()
                End With
                With SysFilesSetup
                    .StartInfo.FileName = WinDir & "\System32\icacls.exe"
                    .StartInfo.Arguments = WinDir & "\System32\" & setup.Language & "\UIRibbon.dll.mui /grant *S-1-5-18:RX *S-1-5-32-544:RX *S-1-5-32-545:RX *S-1-15-2-1:RX ""NT SERVICE\TrustedInstaller"":F /inheritance:r"
                    .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                    .Start()
                    .WaitForExit()
                End With
            End If
            ' \Start userinit.exe
            setup_install_ExplorerStartInfoWindow_en.Show()
            Dim StartExplorerTask As New Process
            With StartExplorerTask
                .StartInfo.FileName = WinDir & "\System32\userinit.exe"
                .Start()
                .WaitForExit()
            End With
            setup_install_ExplorerStartInfoWindow_en.Dispose()
            ProgressBarSetup.PerformStep()
            ProgressBarSetup.Refresh()
            ActionLabel.Text = "UIRibbon.dll and UIRibbonRes.dll have been replaced!"
            ActionLabel.Refresh()
        End If

        ' Installing theme
        ActionLabel.Text = "Installing theme..."
        ActionLabel.Refresh()
        DetailsTextBox.AppendText(ActionLabel.Text)
        FileIO.FileSystem.CopyDirectory(ResFolder & "\Theme\Themes", WinDir & "\Resources\Themes")
        ' \Blue/white address bar
        If setup.AddressBarStyle = "blue" Then
            FileIO.FileSystem.RenameFile(WinDir & "\Resources\Themes\aerorp\aero_blue.msstyles", "aerolite.msstyles")
            FileIO.FileSystem.DeleteFile(WinDir & "\Resources\Themes\aerorp\aero_nonblue.msstyles")
        ElseIf setup.AddressBarStyle = "white" Then
            FileIO.FileSystem.RenameFile(WinDir & "\Resources\Themes\aerorp\aero_nonblue.msstyles", "aerolite.msstyles")
            FileIO.FileSystem.DeleteFile(WinDir & "\Resources\Themes\aerorp\aero_blue.msstyles")
        End If
        DetailsTextBox.AppendText(" Finished!" & vbCrLf)
        ' \Copy wallpaper files
        FileIO.FileSystem.CopyDirectory(ResFolder & "\Theme\Wallpaper\Windows 8 Release Preview", WinDir & "\Web\Wallpaper\Windows 8 Release Preview")
        DetailsTextBox.AppendText("Copying wallpaper files to " & WinDir & "\Web\Wallpaper\Windows 8 Release Preview...")
        ProgressBarSetup.PerformStep()
        ProgressBarSetup.Refresh()
        ActionLabel.Text = "Theme has been installed!"
        ActionLabel.Refresh()
        DetailsTextBox.AppendText(" Finished!" & vbCrLf)

        ' Installing sounds
        If setup.InstSounds = True Then
            ActionLabel.Text = "Installing sound scheme..."
            ActionLabel.Refresh()
            DetailsTextBox.AppendText(ActionLabel.Text & vbCrLf)
            DetailsTextBox.AppendText("Copying sound files to " & WinDir & "\Media\Windows 8 Release Preview...")
            FileIO.FileSystem.CopyDirectory(ResFolder & "\Sounds", WinDir & "\Media\Windows 8 Release Preview")
            ProgressBarSetup.PerformStep()
            ProgressBarSetup.Refresh()
            ActionLabel.Text = "Sound scheme has been installed!"
            ActionLabel.Refresh()
            DetailsTextBox.AppendText(" Finished!" & vbCrLf)
        ElseIf setup.InstSounds = False Then
            DetailsTextBox.AppendText("Removing file " & WinDir & "\Resources\Themes\aerorp_sounds.theme...")
            FileIO.FileSystem.DeleteFile(WinDir & "\Resources\Themes\aerorp_sounds.theme")
            DetailsTextBox.AppendText(" Finished!" & vbCrLf)
        End If

        ' Installing Gadgets
        If setup.InstGadgets = True Then
            ActionLabel.Text = "Installing 8GadgetPack..."
            ActionLabel.Refresh()
            DetailsTextBox.AppendText(ActionLabel.Text)
            Dim GadgetsSetup As New Process
            With GadgetsSetup
                .StartInfo.FileName = WinDir & "\System32\msiexec.exe"
                .StartInfo.Arguments = "/package """ & ResFolder & "\Gadgets\8GadgetPack370Setup.msi"" /passive"
                .Start()
                .WaitForExit()
            End With
            ProgressBarSetup.PerformStep()
            ProgressBarSetup.Refresh()
            ActionLabel.Text = "8GadgetPack has been installed!"
            ActionLabel.Refresh()
            DetailsTextBox.AppendText(" Finished!" & vbCrLf)
        End If

        ' Finalising installation
        If setup.Language = "en-US" Then
            ActionLabel.Text = "Finalizing installation..."
        ElseIf setup.Language = "en-GB" Then
            ActionLabel.Text = "Finalising installation..."
        End If
        ActionLabel.Refresh()
        DetailsTextBox.AppendText(ActionLabel.Text & vbCrLf)
        ' Create registry keys, copy files
        DetailsTextBox.AppendText("Copying file firstrun.xht to " & ResFolder & "...")
        ' \Copy firstrun.xht from resources
        If setup.InstQuero = True Then
            File.WriteAllBytes(ResFolder & "\firstrun.xht", My.Resources.firstrun_en_quero)
        ElseIf setup.InstQuero = False Then
            File.WriteAllBytes(ResFolder & "\firstrun.xht", My.Resources.firstrun_en)
        End If
        DetailsTextBox.AppendText(" Finished!" & vbCrLf)
        ' \Autostart RunOnce\Set XHTML file to autostart
        Dim AutorunRegKey As RegistryKey = Registry.CurrentUser.OpenSubKey("Software\Microsoft\Windows\CurrentVersion\RunOnce", True)
        AutorunRegKey.SetValue("firstrun.xht", """" & ProgramFiles & "\Internet Explorer\iexplore.exe"" " & """" & ResFolder & "\firstrun.xht""", RegistryValueKind.String)
        AutorunRegKey.Close()
        ' \Autostart RunOnce\Set theme file to autostart
        DetailsTextBox.AppendText("Creating autostart entry for applying theme...")
        If setup.InstSounds = True Then
            AutorunRegKey = Registry.CurrentUser.OpenSubKey("Software\Microsoft\Windows\CurrentVersion\RunOnce", True)
            AutorunRegKey.SetValue("aerorp_sounds.theme", WinDir & "\Resources\Themes\aerorp_sounds.theme", RegistryValueKind.String)
            AutorunRegKey.Close()
        ElseIf setup.InstSounds = False Then
            AutorunRegKey = Registry.CurrentUser.OpenSubKey("Software\Microsoft\Windows\CurrentVersion\RunOnce", True)
            AutorunRegKey.SetValue("aerorp.theme", WinDir & "\Resources\Themes\aerorp.theme", RegistryValueKind.String)
            AutorunRegKey.Close()
        End If
        DetailsTextBox.AppendText(" Finished!" & vbCrLf)
        ' \Copy config file programs.xml
        DetailsTextBox.AppendText("Copying config file to " & ProgDir & "\programs.xml...")
        Dim UserLocalConfig = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.PerUserRoamingAndLocal)
        FileIO.FileSystem.CopyFile(UserLocalConfig.FilePath, ProgDir & "\programs.xml")
        DetailsTextBox.AppendText(" Finished!" & vbCrLf)
        ' \Create Control Panel\Programs entry
        DetailsTextBox.AppendText("Creating Control Panel entry...")
        Dim RegKeyControlPanel As RegistryKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall", True)
        RegKeyControlPanel.CreateSubKey("Windows 8.1 to 8 RP Converter")
        Dim RegKeyControlPanelEntry As RegistryKey = RegKeyControlPanel.OpenSubKey("Windows 8.1 to 8 RP Converter", True)
        With RegKeyControlPanelEntry
            .SetValue("DisplayIcon", ResFolder & "\program.ico", RegistryValueKind.String)
            .SetValue("DisplayName", "Windows 8.1 to Windows 8 RP Converter", RegistryValueKind.String)
            .SetValue("DisplayVersion", "0.1 beta", RegistryValueKind.String)
            ' important: calculate program size
            .SetValue("EstimatedSize", 18572, RegistryValueKind.DWord)
            .SetValue("InstallLocation", ProgDir, RegistryValueKind.String)
            .SetValue("ModifyPath", """" & ProgDir & "\modify.exe""", RegistryValueKind.String)
            .SetValue("NoModify", 0, RegistryValueKind.DWord)
            .SetValue("NoRemove", 0, RegistryValueKind.DWord)
            .SetValue("NoRepair", 1, RegistryValueKind.DWord)
            .SetValue("Publisher", "Moha Logiciels", RegistryValueKind.String)
            .SetValue("UninstallString", """" & ProgDir & "\uninstall.exe""", RegistryValueKind.String)
            .Close()
        End With
        RegKeyControlPanel.Close()
        DetailsTextBox.AppendText(" Finished!" & vbCrLf)
        ' \Copy modify.exe and uninstall.exe to program folder
        DetailsTextBox.AppendText("Copy modify.exe and uninstall.exe to " & ProgDir & "...")
        FileIO.FileSystem.WriteAllBytes(ProgDir & "\modify.exe", My.Resources.modify_program, False)
        FileIO.FileSystem.WriteAllBytes(ProgDir & "\uninstall.exe", My.Resources.uninstall_program, False)
        DetailsTextBox.AppendText(" Finished!" & vbCrLf)
        ProgressBarSetup.PerformStep()
        ProgressBarSetup.Refresh()
        ' \Remove temp files and folder
        ActionLabel.Text = "Removing temp files..."
        ActionLabel.Refresh()
        DetailsTextBox.AppendText(ActionLabel.Text)
        Dim LocalAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData)
        FileIO.FileSystem.DeleteDirectory(LocalAppData & "\" & My.Application.Info.AssemblyName, FileIO.DeleteDirectoryOption.DeleteAllContents)
        ProgressBarSetup.PerformStep()
        ProgressBarSetup.Refresh()
        DetailsTextBox.AppendText(" Finished!" & vbCrLf)

        ' Finish installation
        ActionLabel.Text = "Installation has been successfully completed!"
        ActionLabel.Refresh()
        DetailsTextBox.AppendText(ActionLabel.Text & vbCrLf)
        ' \Change Next button to Finish
        NextButton.Text = "&Finish"
        NextButton.Enabled = True
        IsFinished = True
        ShowMessageBoxForRestartOnExit = True
    End Sub

    Private Sub ShowDetailsButton_Click(sender As Object, e As EventArgs) Handles ShowDetailsButton.Click
        DetailsTextBox.Visible = True
        ShowDetailsButton.Enabled = False
    End Sub

    Private Sub NextButton_Click(sender As Object, e As EventArgs) Handles NextButton.Click
        If IsFinished = False Then
            MessageBox.Show("The setup cannot be finished right now.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        ElseIf IsFinished = True Then
            Me.Close()
        End If
    End Sub

    Private Sub CloseButton_Click(sender As Object, e As EventArgs) Handles CloseButton.Click
        If IsFinished = False Then
            MessageBox.Show("The setup cannot be finished right now.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        ElseIf IsFinished = True Then
            Me.Close()
        End If
    End Sub

    Private Sub setup_install_en_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        If IsFinished = False Then
            MessageBox.Show("The setup cannot be finished right now.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        ElseIf IsFinished = True Then
            If e.CloseReason = CloseReason.UserClosing = True Then
                If ShowMessageBoxForRestartOnExit = True Then
                    ' Save log file
                    Dim ProgDir As String = ProgramFiles & "\Windows8.1to8RPConv"
                    If MessageBox.Show("Do you want to save the details pane to a log file in the program’s directory?", "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                        File.AppendAllText(ProgDir & "\win8rpconv_install.log", "Installation date: " & My.Computer.Clock.LocalTime.ToShortDateString & vbCrLf)
                        File.AppendAllText(ProgDir & "\win8rpconv_install.log", "Setup options: [Theme: """ & setup.AddressBarStyle & """, InstONE: """ & setup.InstONE & """, InstUX: """ & setup.InstUX & """, InstQuero: """ & setup.InstQuero & """, InstSysFiles: """ & setup.InstSysFiles & """, InstSounds: """ & setup.InstSounds & """]" & vbCrLf & "Install log:" & vbCrLf)
                        File.AppendAllLines(ProgDir & "\win8rpconv_install.log", DetailsTextBox.Lines)
                    End If
                    ' Restart prompt
                    Dim ShutDownProcess As New Process
                    If MessageBox.Show("You need to restart your system so that the changes take effect. Do you want to restart now?", "Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) = DialogResult.Yes Then
                        With ShutDownProcess
                            .StartInfo.FileName = WinDir & "\System32\shutdown.exe"
                            .StartInfo.Arguments = "/r /t 3"
                            .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                            .Start()
                        End With
                        setup.Close()
                    ElseIf DialogResult.No Then
                        setup.Close()
                    End If
                ElseIf ShowMessageBoxForRestartOnExit = False Then
                    setup.Close()
                End If
            End If
        End If
    End Sub
End Class