Imports System.Drawing.Text
Imports System.IO

Public Class SplashScreen

    Private Sub SplashScreen_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If Not FileIO.FileSystem.FileExists(Environment.GetFolderPath(Environment.SpecialFolder.Fonts) & "\SNAP____.TTF") Then
            File.WriteAllBytes(My.Computer.FileSystem.SpecialDirectories.Temp & "\SNAP____.TTF", My.Resources.SnapFontFile)
            Using LabelFonts As New PrivateFontCollection
                LabelFonts.AddFontFile(My.Computer.FileSystem.SpecialDirectories.Temp & "\SNAP____.TTF")
                VersionLabel.Font = New Font(LabelFonts.Families(0), 16.2, FontStyle.Regular)
                LabelFonts.Dispose()
            End Using
        End If
    End Sub
End Class