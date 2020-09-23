Attribute VB_Name = "mainModule"

Option Explicit
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function GetTextMetrics Lib "gdi32" Alias "GetTextMetricsA" (ByVal hdc As Long, lpMetrics As TEXTMETRIC) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Global Const Titre = "Pro de la Police "
Global Choix As Integer
Global gv010C As Integer
Global gv010E As Integer
Global NomDesFamilles
Global Language As String


Type TEXTMETRIC
    tmHeight As Long
    tmAscent As Long
    tmDescent As Long
    tmInternalLeading As Long
    tmExternalLeading As Long
    tmAveCharWidth As Long
    tmMaxCharWidth As Long
    tmWeight As Long
    tmOverhang As Long
    tmDigitizedAspectX As Long
    tmDigitizedAspectY As Long
    tmFirstChar As Byte
    tmLastChar As Byte
    tmDefaultChar As Byte
    tmBreakChar As Byte
    tmItalic As Byte
    tmUnderlined As Byte
    tmStruckOut As Byte
    tmPitchAndFamily As Byte
    tmCharSet As Byte
End Type
Function familleDePolice(p01D6 As TEXTMETRIC) As Variant
    Dim l01D8 As Integer
    l01D8 = (p01D6.tmPitchAndFamily) And &HF0
    Select Case l01D8
        Case 0:
            Select Case gv010E
                Case Else
                If Language = "Francais" Then
                    familleDePolice = "Non Classée!..."
                    Else
                    familleDePolice = "Not classified!..."
                End If
            End Select
        Case 16
            familleDePolice = "Roman"
        Case 32
            familleDePolice = "Swiss"
        Case 48
            familleDePolice = "Modern"
        Case 64
            familleDePolice = "Script"
        Case 80
        If Language = "Francais" Then
            familleDePolice = "Décorative"
                    Else
            familleDePolice = "Decorative"
                End If
        Case Else
            familleDePolice = Str$(l01D8) & "..."
    End Select
End Function

Public Function NombreDePoliceDansFamille() As String
If Language = "Francais" Then
    NombreDePoliceDansFamille = " " & frmPrinc.cmbPoliceHaut.ListCount & "/" & frmPrinc.cmbToutePolice.ListCount & " Police(s) Filtrée(s)"
    Else
    NombreDePoliceDansFamille = " " & frmPrinc.cmbPoliceHaut.ListCount & "/" & frmPrinc.cmbToutePolice.ListCount & " Filtered font(s)"
    End If
End Function

