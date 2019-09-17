VERSION 5.00
Begin VB.Form FrmSplash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3090
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   5490
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmSplash.frx":0000
   ScaleHeight     =   3090
   ScaleWidth      =   5490
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrTimer 
      Interval        =   2000
      Left            =   0
      Top             =   360
   End
End
Attribute VB_Name = "FrmSplash"
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

Private Sub Form_Click()
    Unload Me
    Load FrmMain
    FrmMain.Show
End Sub

Private Sub tmrTimer_Timer()
    tmrTimer.Enabled = False
    Unload Me
    Load FrmMain
    FrmMain.Show
End Sub

