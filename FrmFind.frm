VERSION 5.00
Begin VB.Form Frmfind 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2040
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   4155
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   4155
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkWholeWords 
      Appearance      =   0  'Flat
      BackColor       =   &H00996666&
      Caption         =   "Whole Words"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   1935
   End
   Begin VB.CheckBox chkMatchcase 
      Appearance      =   0  'Flat
      BackColor       =   &H00996666&
      Caption         =   "Match Case"
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   1815
   End
   Begin VB.ComboBox cmbFind 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   2655
   End
   Begin VB.PictureBox PicSelectPonsel 
      Appearance      =   0  'Flat
      BackColor       =   &H00996666&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      ScaleHeight     =   375
      ScaleWidth      =   3900
      TabIndex        =   0
      Top             =   0
      Width           =   3900
      Begin VB.Label LblSPonsel 
         BackStyle       =   0  'Transparent
         Caption         =   "Find Text"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   120
         TabIndex        =   1
         Top             =   15
         Width           =   1335
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   1800
      Width           =   4095
   End
   Begin VB.Label LblFind 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Find"
      Height          =   255
      Left            =   3000
      MouseIcon       =   "FrmFind.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   640
      Width           =   975
   End
   Begin VB.Label LblFindNext 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Find Next"
      Height          =   255
      Left            =   3000
      MouseIcon       =   "FrmFind.frx":030A
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   1000
      Width           =   975
   End
   Begin VB.Label LblCancel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cancel"
      Height          =   255
      Left            =   3000
      MouseIcon       =   "FrmFind.frx":0614
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   1360
      Width           =   975
   End
   Begin VB.Shape ShpSend 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   300
      Left            =   3000
      Shape           =   4  'Rounded Rectangle
      Top             =   1320
      Width           =   1005
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   300
      Left            =   3000
      Shape           =   4  'Rounded Rectangle
      Top             =   600
      Width           =   1005
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   300
      Left            =   3000
      Shape           =   4  'Rounded Rectangle
      Top             =   960
      Width           =   1005
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00996666&
      BorderColor     =   &H00996666&
      FillColor       =   &H00996666&
      FillStyle       =   0  'Solid
      Height          =   1335
      Left            =   0
      Top             =   440
      Width           =   4215
   End
End
Attribute VB_Name = "Frmfind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'--> AVACO - ACCESS SIEMENS GSM CELLPHONE WITH AT+COMMAND.
'--> version 1.00
'--> Version Language : English
'--> By Agus Ramadhani
'--> avaco software
'--> http://avaco-software.tripod.com
'--> avaco@9cy.Com
'--> 2002-2003
'--> Don't forget to Vote :)

Dim A$, Found, P, Fpos

Private Sub Form_Load()
MeOnTop Me
chkMatchcase.Value = Unchecked
chkWholeWords.Value = Unchecked
cmbFind.Text = FindText$
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
LblFind.FontBold = False
LblFindNext.FontBold = False
LblCancel.FontBold = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
Cancel = 1
End Sub

Private Sub LblCancel_Click()
Me.Hide
End Sub

Private Sub LblFind_Click()
On Error GoTo err:
Label1.Caption = ""
If cmbFind.Text = "" Then
   Label1.Caption = " No target to find"
   Exit Sub
End If
A$ = cmbFind.Text
If chkMatchcase.Value Then
   Found = FrmMain.RtbListATC.Find(A$, 0, , rtfMatchCase)
Else
   Found = FrmMain.RtbListATC.Find(A$, 0)
End If
If Found <> -1 Then
   FrmMain.RtbListATC.SetFocus
Else
   Label1.Caption = " Text not found"
End If
Exit Sub
err:
Exit Sub
End Sub


Private Sub LblFindNext_Click()
On Error GoTo err
Label1.Caption = ""
If cmbFind.Text = "" Then
  Label1.Caption = " No target to find"
   Exit Sub
End If
A$ = cmbFind.Text
P = FrmMain.RtbListATC.SelStart + FrmMain.RtbListATC.SelLength + 1
If P = Fpos Then P = 0
If chkMatchcase.Value Then
   Found = FrmMain.RtbListATC.Find(A$, P, , rtfMatchCase)
Else
   Found = FrmMain.RtbListATC.Find(A$, P)
End If
If Found <> -1 Then
    FrmMain.RtbListATC.SetFocus
Else
   Fpos = P
  Label1.Caption = " The specified region has been searched"
End If
Exit Sub
err:
Exit Sub
End Sub

Private Sub LblCancel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
LblFind.FontBold = False
LblFindNext.FontBold = False
LblCancel.FontBold = True
End Sub

Private Sub LblFind_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
LblFind.FontBold = True
LblFindNext.FontBold = False
LblCancel.FontBold = False
End Sub

Private Sub LblFindNext_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
LblFind.FontBold = False
LblFindNext.FontBold = True
LblCancel.FontBold = False
End Sub

Private Sub LblSPonsel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormDrag Me
End Sub

Private Sub PicSelectPonsel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormDrag Me
End Sub
