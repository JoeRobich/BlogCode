Imports System
Imports EnvDTE
Imports EnvDTE80
Imports EnvDTE90
Imports EnvDTE90a
Imports EnvDTE100
Imports System.Diagnostics

Public Module Settings
    Private Const FONT_SIZE_CATEGORY As String = "FontsAndColors"
    Private Const FONT_SIZE_PAGE As String = "TextEditor"
    Private Const FONT_SIZE_NAME As String = "FontSize"

    Private RootFolder As String = "[Your Root]\Documents\Visual Studio 2010\Settings\"

    Public Sub ImportDefaultSettings()
        SetColorTheme("DefaultTheme")
    End Sub

    Public Sub ImportCodingSettings()
        SetColorTheme("VibrantInkTheme")
    End Sub

    Public Sub TogglePresentationSettings()
        Dim fontSizeProperty As EnvDTE.Property = GetDTEProperty(FONT_SIZE_CATEGORY, FONT_SIZE_PAGE, FONT_SIZE_NAME)

        If fontSizeProperty.Value = 15 Then
            fontSizeProperty.Value = 10
        Else
            fontSizeProperty.Value = 15
        End If
    End Sub

    Public Sub ImportKnRFormatting()
        ImportSettingsFile("KnRFormatting")
        DTE.ExecuteCommand("Edit.FormatDocument")
    End Sub

    Public Sub ImportAllmanFormatting()
        ImportSettingsFile("AllmanFormatting")
        DTE.ExecuteCommand("Edit.FormatDocument")
    End Sub

    Private Sub SetColorTheme(ByVal fileName As String)
        Dim fontSizeProperty As EnvDTE.Property = GetDTEProperty(FONT_SIZE_CATEGORY, FONT_SIZE_PAGE, FONT_SIZE_NAME)
        Dim fontSize As Integer = fontSizeProperty.Value

        ImportSettingsFile(fileName)

        fontSizeProperty.Value = fontSize
    End Sub

    Private Function GetDTEProperty(ByVal categoryName As String, ByVal pageName As String, ByVal propertyName As String) As EnvDTE.Property
        For Each prop In DTE.Properties(categoryName, pageName)
            If prop.Name = propertyName Then
                Return prop
            End If
        Next
    End Function

    Private Sub ImportSettingsFile(ByVal FileName As String)
        FileName = IO.Path.Combine(RootFolder, FileName & ".vssettings")
        DTE.ExecuteCommand("Tools.ImportandExportSettings", "-import:""" & FileName & """")
    End Sub

End Module
