Attribute VB_Name = "MdlSyntax"
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

'--> This Syntax sample Code from Brian Bender | brianbender77@hotmail.com, Thanks for sample code :)

Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long
Public bInQuotes As Boolean
Const BlueKeyWords = "#Const*#Else*#ElseIf*#*Error*Ok*"
Const lBlueKeyWords = "#const*#else*#elseif*#*error*ok*"

Private Function IsArrayEmpty(arr As Variant) As Boolean
    On Error Resume Next
    If UBound(arr) > 0 Then IsArrayEmpty = False
    If err.Number > 0 Then IsArrayEmpty = True
End Function

Private Function In_Quote(sSegment As String) As Boolean
    Dim pos As Integer
    Dim start As Integer
    start = 1
    pos = 1
    Do Until pos = 0
        pos = InStr(start, sSegment, Chr(34))
        If pos > 0 Then bInQuotes = Not bInQuotes
        start = pos + 1
    Loop
    In_Quote = bInQuotes
End Function

Public Sub Colorize(RtbListATC As RichTextBox, sText As String)
    If sText = "" Then Exit Sub
    DoEvents
    Screen.MousePointer = vbHourglass
    Dim lTime As Long
    Dim arCode() As String
    Dim arSegment() As String
    Dim iLineCount As Integer
    Dim iSegment As Integer
    Dim bPartialComment As Boolean
    arCode = Split(sText, vbCrLf)
    With RtbListATC
    lTime = GetTickCount
    LockWindowUpdate .hWnd
    For iLineCount = LBound(arCode) To UBound(arCode)
        DoEvents
        If Len(Trim(arCode(iLineCount))) > 0 Then
           If Left$(Trim(arCode(iLineCount)), 1) = "Rem " Or Left$(Trim(arCode(iLineCount)), 1) = "'" Then
              .SelColor = QBColor(2)
              .SelText = arCode(iLineCount) & vbCrLf
              Else
              arSegment = Split(arCode(iLineCount), " ")
              For iSegment = LBound(arSegment) To UBound(arSegment)
              If Left$(arSegment(iSegment), 1) = "'" Then
              If Not bInQuotes Or bPartialComment Then
                 .SelColor = QBColor(2)
                 .SelText = arSegment(iSegment) & " "
                  bPartialComment = True
              Else
                 .SelText = arSegment(iSegment) & " "
              End If
              ElseIf Left$(arSegment(iSegment), 1) = "" Then
                    .SelText = arSegment(iSegment) & " "
              Else
              If bPartialComment Then
                 .SelColor = QBColor(2)
                 .SelText = arSegment(iSegment) & " "
              Else
              If InStr(1, lBlueKeyWords, LCase(arSegment(iSegment))) And Not Len(arSegment(iSegment)) = 1 Then
                 If Not bInQuotes Then
                 .SelColor = vbBlue
                 .SelText = Mid$(BlueKeyWords, InStr(1, lBlueKeyWords, LCase(arSegment(iSegment))), Len(arSegment(iSegment))) & " "
                 Else
                 .SelText = arSegment(iSegment) & " "
                 End If
                 Else
                 .SelColor = vbRed
                 .SelText = arSegment(iSegment) & " "
             End If
           End If
         End If
        Next iSegment
        If Not iLineCount = UBound(arCode) Then .SelText = vbCrLf
        End If
        Else
       .SelText = vbCrLf
        End If
          bPartialComment = False
          bInQuotes = False
        Next iLineCount
        .SelColor = QBColor(0)
    End With
    LockWindowUpdate 0&
    Screen.MousePointer = vbDefault
    lTime = GetTickCount - lTime
    FrmMain.LblTLT.Caption = "Time Load Text: " & lTime & " ms"
End Sub





