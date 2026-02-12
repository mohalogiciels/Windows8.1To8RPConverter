Imports System.Drawing.Text
Imports System.IO
Imports System.Runtime.InteropServices

Public Class SplashScreen

    Private Sub SplashScreen_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim FontDirectory As String = Environment.GetFolderPath(Environment.SpecialFolder.Fonts)
        If Not FileIO.FileSystem.FileExists(FontDirectory & "\SNAP____.TTF") Then
            Using ProgramFonts As New PrivateFontCollection
                Dim FontBuffer As IntPtr = Marshal.AllocCoTaskMem(My.Resources.SnapFontFile.Length)
                Marshal.Copy(My.Resources.SnapFontFile, 0, FontBuffer, My.Resources.SnapFontFile.Length)
                ProgramFonts.AddMemoryFont(FontBuffer, My.Resources.SnapFontFile.Length)
                With VersionLabel
                    .UseCompatibleTextRendering = True
                    .Font = New Font(ProgramFonts.Families(0), 16.2, FontStyle.Regular)
                End With
                Marshal.FreeCoTaskMem(FontBuffer)
                ProgramFonts.Dispose()
            End Using
        End If
    End Sub
End Class