VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} InvArtikelForm
   Caption         =   "Inventur - Artikel pruefen"
   ClientHeight    =   4500
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   9660
   StartUpPosition =   1
   Begin MSForms.Label lbl_Header
      Caption         =   "Inventur - Artikel pruefen"
      Height          =   510
      Left            =   0
      Top             =   0
      Width           =   9660
      ForeColor       =   &H00FFFFFF
      BackColor       =   &H00A03070
      TextAlign       =   2
      BeginProperty Font
         Name            =   "Arial"
         Size            =   13
         Bold            =   True
      EndProperty
   End
   Begin MSForms.Label lbl_key1
      Caption         =   "Artikel:"
      Height          =   270
      Left            =   60
      Top             =   600
      Width           =   1200
      BeginProperty Font
         Name            =   "Arial"
         Size            =   9
         Bold            =   True
      EndProperty
   End
   Begin MSForms.Label lbl_Artikel
      Caption         =   ""
      Height          =   270
      Left            =   1290
      Top             =   600
      Width           =   8310
   End
   Begin MSForms.Label lbl_key2
      Caption         =   "Art.-Nr.:"
      Height          =   270
      Left            =   60
      Top             =   930
      Width           =   1200
      BeginProperty Font
         Name            =   "Arial"
         Size            =   9
         Bold            =   True
      EndProperty
   End
   Begin MSForms.Label lbl_ArtNr
      Caption         =   ""
      Height          =   270
      Left            =   1290
      Top             =   930
      Width           =   3000
   End
   Begin MSForms.Label lbl_key3
      Caption         =   "EAN:"
      Height          =   270
      Left            =   4500
      Top             =   930
      Width           =   900
      BeginProperty Font
         Name            =   "Arial"
         Size            =   9
         Bold            =   True
      EndProperty
   End
   Begin MSForms.Label lbl_EAN
      Caption         =   ""
      Height          =   270
      Left            =   5460
      Top             =   930
      Width           =   4140
   End
   Begin MSForms.Label lbl_key4
      Caption         =   "VK-Preis:"
      Height          =   270
      Left            =   60
      Top             =   1260
      Width           =   1200
      BeginProperty Font
         Name            =   "Arial"
         Size            =   9
         Bold            =   True
      EndProperty
   End
   Begin MSForms.Label lbl_VK
      Caption         =   ""
      Height          =   270
      Left            =   1290
      Top             =   1260
      Width           =   2400
   End
   Begin MSForms.Label lbl_key5
      Caption         =   "EK-Preis:"
      Height          =   270
      Left            =   4500
      Top             =   1260
      Width           =   1200
      BeginProperty Font
         Name            =   "Arial"
         Size            =   9
         Bold            =   True
      EndProperty
   End
   Begin MSForms.Label lbl_EK
      Caption         =   ""
      Height          =   270
      Left            =   5760
      Top             =   1260
      Width           =   2400
   End
   Begin MSForms.Label lbl_key6
      Caption         =   "Lagerort:"
      Height          =   270
      Left            =   60
      Top             =   1590
      Width           =   1200
      BeginProperty Font
         Name            =   "Arial"
         Size            =   9
         Bold            =   True
      EndProperty
   End
   Begin MSForms.Label lbl_Lager
      Caption         =   ""
      Height          =   270
      Left            =   1290
      Top             =   1590
      Width           =   3600
   End
   Begin MSForms.Label lbl_key7
      Caption         =   "Warengruppe:"
      Height          =   270
      Left            =   5160
      Top             =   1590
      Width           =   1680
      BeginProperty Font
         Name            =   "Arial"
         Size            =   9
         Bold            =   True
      EndProperty
   End
   Begin MSForms.Label lbl_WG
      Caption         =   ""
      Height          =   270
      Left            =   6900
      Top             =   1590
      Width           =   2700
   End
   Begin MSForms.Label lbl_sep
      Caption         =   ""
      Height          =   30
      Left            =   0
      Top             =   1980
      Width           =   9660
      BackColor       =   &H00C0C0C0
   End
   Begin MSForms.Label lbl_keySoll
      Caption         =   "SOLL (System):"
      Height          =   330
      Left            =   60
      Top             =   2100
      Width           =   2400
      BeginProperty Font
         Name            =   "Arial"
         Size            =   10
         Bold            =   True
      EndProperty
   End
   Begin MSForms.Label lbl_Soll
      Caption         =   ""
      Height          =   420
      Left            =   2520
      Top             =   2040
      Width           =   1800
      TextAlign       =   2
      BeginProperty Font
         Name            =   "Arial"
         Size            =   16
         Bold            =   True
      EndProperty
   End
   Begin MSForms.Label lbl_arrow
      Caption         =   "-->"
      Height          =   330
      Left            =   4500
      Top             =   2130
      Width           =   600
      TextAlign       =   2
      BeginProperty Font
         Name            =   "Arial"
         Size            =   12
         Bold            =   True
      EndProperty
   End
   Begin MSForms.Label lbl_keyIst
      Caption         =   "TATSAECHLICH GEZAEHLT:"
      Height          =   270
      Left            =   5280
      Top             =   2040
      Width           =   4320
      ForeColor       =   &H00A03070
      BeginProperty Font
         Name            =   "Arial"
         Size            =   9
         Bold            =   True
      EndProperty
   End
   Begin MSForms.TextBox txt_Ist
      Height          =   570
      Left            =   5280
      Top             =   2400
      Width           =   4320
      BeginProperty Font
         Name            =   "Arial"
         Size            =   18
         Bold            =   True
      EndProperty
   End
   Begin MSForms.Label lbl_Diff
      Caption         =   ""
      Height          =   510
      Left            =   0
      Top             =   3180
      Width           =   9660
      TextAlign       =   2
      BeginProperty Font
         Name            =   "Arial"
         Size            =   11
         Bold            =   True
      EndProperty
   End
   Begin MSForms.CommandButton btn_Abbrechen
      Caption         =   "Abbrechen"
      Height          =   510
      Left            =   5880
      Top             =   3870
      Width           =   1560
      ForeColor       =   &H00000000
      BackColor       =   &H00D9D9D9
      BeginProperty Font
         Name            =   "Arial"
         Size            =   10
         Bold            =   False
      EndProperty
   End
   Begin MSForms.CommandButton btn_OK
      Caption         =   "Uebernehmen"
      Height          =   510
      Left            =   7740
      Top             =   3870
      Width           =   1860
      ForeColor       =   &H00FFFFFF
      BackColor       =   &H00A03070
      BeginProperty Font
         Name            =   "Arial"
         Size            =   10
         Bold            =   True
      EndProperty
   End
End
Attribute VB_Name = "InvArtikelForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub Befuellen(artikel As String, artNr As String, ean As String, _
                     vk As Double, ek As Double, lager As String, _
                     wg As String, soll As Double, quellZeile As Long)
    Me.Tag             = CStr(quellZeile)
    lbl_Artikel.Caption = artikel
    lbl_ArtNr.Caption   = artNr
    lbl_EAN.Caption     = ean
    lbl_VK.Caption      = Format(vk, "0.00") & " EUR"
    lbl_EK.Caption      = Format(ek, "0.00") & " EUR"
    lbl_Lager.Caption   = lager
    lbl_WG.Caption      = wg
    lbl_Soll.Caption    = Format(soll, "0") & " Stk"
    txt_Ist.Value       = ""
    lbl_Diff.Caption    = ""
    lbl_Diff.BackColor  = &H8000000F
    txt_Ist.SetFocus
End Sub

Private Sub txt_Ist_Change()
    If Trim(txt_Ist.Value) = "" Then
        lbl_Diff.Caption = "" : lbl_Diff.BackColor = &H8000000F : Exit Sub
    End If
    Dim ist  As Double : ist  = Val(txt_Ist.Value)
    Dim soll As Double : soll = Val(Replace(lbl_Soll.Caption, " Stk", ""))
    Dim diff As Double : diff = ist - soll
    If diff > 0 Then
        lbl_Diff.Caption   = "+" & Format(diff, "0") & " Stk  (mehr vorhanden)"
        lbl_Diff.BackColor = RGB(198, 239, 206)
        lbl_Diff.ForeColor = RGB(0, 97, 0)
    ElseIf diff < 0 Then
        lbl_Diff.Caption   = Format(diff, "0") & " Stk  (weniger vorhanden)"
        lbl_Diff.BackColor = RGB(255, 199, 206)
        lbl_Diff.ForeColor = RGB(156, 0, 6)
    Else
        lbl_Diff.Caption   = "+/-0  Bestand stimmt"
        lbl_Diff.BackColor = RGB(198, 239, 206)
        lbl_Diff.ForeColor = RGB(0, 97, 0)
    End If
End Sub

Private Sub txt_Ist_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then btn_OK_Click
    If KeyCode = 27 Then btn_Abbrechen_Click
End Sub

Private Sub btn_OK_Click()
    If Trim(txt_Ist.Value) = "" Then
        MsgBox "Bitte eine Menge eingeben.", vbExclamation : Exit Sub
    End If
    Dim ist   As Double : ist   = Val(txt_Ist.Value)
    Dim soll  As Double : soll  = Val(Replace(lbl_Soll.Caption, " Stk", ""))
    Dim diff  As Double : diff  = ist - soll
    Dim zeile As Long   : zeile = Val(Me.Tag)
    Dim ws    As Worksheet
    For Each ws In ThisWorkbook.Sheets
        If InStr(1, ws.Name, "InvSuch", vbTextCompare) > 0 Then Exit For
    Next ws
    If Not ws Is Nothing And zeile > 0 Then
        ws.Cells(zeile, 8).Value = ist
        If diff = 0 Then
            ws.Rows(zeile).Interior.Color = RGB(198, 239, 206)
        ElseIf diff > 0 Then
            ws.Rows(zeile).Interior.Color = RGB(255, 235, 156)
        Else
            ws.Rows(zeile).Interior.Color = RGB(255, 199, 206)
        End If
    End If
    Me.Hide
End Sub

Private Sub btn_Abbrechen_Click()
    Me.Hide
End Sub
