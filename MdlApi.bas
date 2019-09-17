Attribute VB_Name = "MdlAPI"
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

Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, _
    ByVal hwndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, _
    ByVal cy As Long, ByVal wflags As Long) As Long

Declare Sub ReleaseCapture Lib "user32" ()
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Public Declare Function LockWindowUpdate Lib "user32" _
(ByVal hWnd As Long) As Long
Public Const WM_COPY = &H301
Public FindText$

Public Declare Function ShellExecute Lib _
   "shell32.dll" Alias "ShellExecuteA" _
   (ByVal hWnd As Long, _
   ByVal lpOperation As String, _
   ByVal lpFile As String, _
   ByVal lpParameters As String, _
   ByVal lpDirectory As String, _
   ByVal nShowCmd As Long) As Long
    
Public Const SW_SHOWNORMAL = 1

Public Sub FormDrag(TheForm As Form)
    ReleaseCapture
    Call SendMessage(TheForm.hWnd, &HA1, 2, 0&)
End Sub

Public Sub MeOnTop(Form As Form)
    SetWindowPos Form.hWnd, -1, 0, 0, 0, 0, 1 Or 2
End Sub


Public Sub MeDown(Form As Form)
    SetWindowPos Form.hWnd, -2, 0, 0, 0, 0, 1 Or 2
End Sub

