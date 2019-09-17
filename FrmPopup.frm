VERSION 5.00
Begin VB.Form FrmPopup 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   360
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   1890
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   360
   ScaleWidth      =   1890
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   0
      Top             =   0
   End
   Begin VB.Menu Menu 
      Caption         =   "Menu"
      Begin VB.Menu MnuSend 
         Caption         =   "Send Text to AT+Command Test"
      End
      Begin VB.Menu spr 
         Caption         =   "-"
      End
      Begin VB.Menu MnuCopy 
         Caption         =   "Copy Text"
      End
      Begin VB.Menu MnuFind 
         Caption         =   "Find Text"
      End
   End
End
Attribute VB_Name = "FrmPopup"
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

Private Sub MnuCopy_Click()
Timer1.Enabled = False
LockWindowUpdate FrmMain.RtbListATC.hWnd
SendMessage FrmMain.RtbListATC.hWnd, WM_COPY, 0, 0
DoEvents
LockWindowUpdate 0&
Timer1.Enabled = True
End Sub

Private Sub MnuFind_Click()
Dim A$, N
A$ = FindText$
If A$ <> "" Then
   For N = 0 To Frmfind.cmbFind.ListCount - 1
       If A$ = Frmfind.cmbFind.List(N) Then Exit For
   Next N
   If N = Frmfind.cmbFind.ListCount Then Frmfind.cmbFind.AddItem A$, (0)
   Frmfind.cmbFind.Text = A$
End If
Frmfind.Show
End Sub

Private Sub Form_Unload(Cancel As Integer)
Frmfind.Visible = False
Cancel = 1
End Sub

Private Sub MnuSend_Click()
Dim A$, N
A$ = FindText$
If A$ <> "" Then
   For N = 0 To FrmMain.CmbATC.ListCount - 1
       If A$ = FrmMain.CmbATC.List(N) Then Exit For
   Next N
   If N = FrmMain.CmbATC.ListCount Then FrmMain.CmbATC.AddItem A$, (0)
   FrmMain.CmbATC.Text = A$
End If
End Sub
