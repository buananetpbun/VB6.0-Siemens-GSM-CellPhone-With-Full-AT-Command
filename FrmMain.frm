VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form FrmMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Avaco - ATC v1.00"
   ClientHeight    =   7635
   ClientLeft      =   150
   ClientTop       =   150
   ClientWidth     =   9870
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7635
   ScaleWidth      =   9870
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3255
      Left            =   2640
      ScaleHeight     =   3255
      ScaleWidth      =   7215
      TabIndex        =   31
      Top             =   1000
      Width           =   7215
      Begin VB.Image ImgAvaco 
         Height          =   3120
         Left            =   0
         Picture         =   "FrmMain.frx":000C
         Top             =   60
         Width           =   7110
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3255
      Left            =   2640
      ScaleHeight     =   3255
      ScaleWidth      =   7215
      TabIndex        =   30
      Top             =   1000
      Visible         =   0   'False
      Width           =   7215
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Version 1.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   480
         Width           =   6735
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "If you find bugs on my program, please give me report and send from my email. Thanks."
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   120
         TabIndex        =   38
         Top             =   1920
         Width           =   6735
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "AVACO - ACCESS SIEMENS GSM CELLPHONE WITH AT+COMMAND."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   240
         Width           =   6735
      End
      Begin VB.Image Image3 
         Height          =   480
         Left            =   4680
         Picture         =   "FrmMain.frx":7815
         Top             =   2520
         Width           =   480
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   0
         Picture         =   "FrmMain.frx":80DF
         Top             =   2520
         Width           =   480
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "All rights reserved"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   1320
         Width           =   2535
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "By Agus Ramadhani"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   1080
         Width           =   2535
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright © 2002 AVACO"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   840
         Width           =   2535
      End
      Begin VB.Label LblAvacoOL 
         BackStyle       =   0  'Transparent
         Caption         =   "Http://avaco-software.tripod.com"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   600
         MouseIcon       =   "FrmMain.frx":89A9
         MousePointer    =   99  'Custom
         TabIndex        =   33
         Top             =   2760
         Width           =   2535
      End
      Begin VB.Label LblMail 
         BackStyle       =   0  'Transparent
         Caption         =   "email : Avaco@9cy.com"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   5280
         MouseIcon       =   "FrmMain.frx":8CB3
         MousePointer    =   99  'Custom
         TabIndex        =   32
         Top             =   2760
         Width           =   1815
      End
      Begin VB.Image ImgLogo 
         Height          =   2970
         Left            =   480
         Picture         =   "FrmMain.frx":8FBD
         Top             =   120
         Width           =   6465
      End
   End
   Begin VB.ComboBox CmbATC 
      BackColor       =   &H00996666&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   6480
      TabIndex        =   27
      Text            =   "AT+Command"
      Top             =   4560
      Width           =   2190
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00996666&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   0
      ScaleHeight     =   135
      ScaleWidth      =   9855
      TabIndex        =   26
      Top             =   7440
      Width           =   9850
   End
   Begin RichTextLib.RichTextBox RtbTemp 
      Height          =   375
      Left            =   120
      TabIndex        =   22
      Top             =   4440
      Visible         =   0   'False
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   661
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"FrmMain.frx":129E9
   End
   Begin VB.Timer scrolllabel1 
      Interval        =   100
      Left            =   120
      Top             =   5880
   End
   Begin VB.Timer scrolllabel2 
      Interval        =   100
      Left            =   120
      Top             =   5400
   End
   Begin VB.PictureBox PicATC 
      Appearance      =   0  'Flat
      BackColor       =   &H00996666&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2640
      ScaleHeight     =   255
      ScaleWidth      =   7230
      TabIndex        =   10
      Top             =   740
      Width           =   7230
      Begin VB.Label LblATCommand 
         BackStyle       =   0  'Transparent
         Caption         =   "Home"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   15
         Width           =   2535
      End
   End
   Begin VB.PictureBox Picmenu 
      Appearance      =   0  'Flat
      BackColor       =   &H00996666&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   120
      ScaleHeight     =   675
      ScaleWidth      =   9630
      TabIndex        =   6
      Top             =   20
      Width           =   9630
      Begin VB.Label LblHome 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Home"
         Height          =   255
         Left            =   2520
         MouseIcon       =   "FrmMain.frx":12AAB
         MousePointer    =   99  'Custom
         TabIndex        =   40
         Top             =   30
         Width           =   975
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00404040&
         X1              =   3480
         X2              =   3480
         Y1              =   270
         Y2              =   -120
      End
      Begin VB.Image Image1 
         Height          =   675
         Left            =   615
         Picture         =   "FrmMain.frx":12DB5
         Stretch         =   -1  'True
         Top             =   -15
         Width           =   1830
      End
      Begin VB.Label LblATCSupport 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "AT+Command"
         Height          =   255
         Left            =   3480
         MouseIcon       =   "FrmMain.frx":13F11
         MousePointer    =   99  'Custom
         TabIndex        =   28
         Top             =   30
         Width           =   1815
      End
      Begin VB.Image Image5 
         Height          =   735
         Left            =   0
         Picture         =   "FrmMain.frx":1421B
         Top             =   -30
         Width           =   690
      End
      Begin VB.Label LblMinimize 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Minimize"
         Height          =   255
         Left            =   6240
         MouseIcon       =   "FrmMain.frx":147EF
         MousePointer    =   99  'Custom
         TabIndex        =   24
         Top             =   30
         Width           =   1095
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00404040&
         X1              =   6240
         X2              =   6240
         Y1              =   270
         Y2              =   -120
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00404040&
         X1              =   5280
         X2              =   5280
         Y1              =   270
         Y2              =   -120
      End
      Begin VB.Label LblAbout 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "About"
         Height          =   255
         Left            =   5280
         MouseIcon       =   "FrmMain.frx":14AF9
         MousePointer    =   99  'Custom
         TabIndex        =   18
         Top             =   30
         Width           =   975
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00404040&
         X1              =   7320
         X2              =   7320
         Y1              =   270
         Y2              =   -120
      End
      Begin VB.Label LblExit 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Exit"
         Height          =   255
         Left            =   7350
         MouseIcon       =   "FrmMain.frx":14E03
         MousePointer    =   99  'Custom
         TabIndex        =   15
         Top             =   30
         Width           =   975
      End
      Begin VB.Shape ShpMenu 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         Height          =   420
         Left            =   2520
         Shape           =   4  'Rounded Rectangle
         Top             =   -120
         Width           =   5805
      End
      Begin VB.Label LblVesion 
         BackStyle       =   0  'Transparent
         Caption         =   "Version 1.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   8440
         TabIndex        =   14
         Top             =   40
         Width           =   1335
      End
      Begin VB.Label LblDesc 
         BackStyle       =   0  'Transparent
         Caption         =   "ACCESS SIEMENS GSM CELLPHONE WITH AT+COMMAND."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   5100
         TabIndex        =   7
         Top             =   435
         Width           =   4530
      End
   End
   Begin VB.PictureBox PicSShot 
      Appearance      =   0  'Flat
      BackColor       =   &H00996666&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   2580
      TabIndex        =   5
      Top             =   4240
      Width           =   2580
      Begin VB.Label LblSShot 
         BackStyle       =   0  'Transparent
         Caption         =   "Screen Shot"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   20
         Width           =   1335
      End
   End
   Begin VB.PictureBox PicSelectPonsel 
      Appearance      =   0  'Flat
      BackColor       =   &H00996666&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   2580
      TabIndex        =   4
      Top             =   740
      Width           =   2580
      Begin VB.Label LblSPonsel 
         BackStyle       =   0  'Transparent
         Caption         =   "Select Ponsel"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   20
         Width           =   1335
      End
   End
   Begin VB.TextBox TxtTestATC 
      Appearance      =   0  'Flat
      BackColor       =   &H00996666&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   2235
      Left            =   2640
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Text            =   "FrmMain.frx":1510D
      Top             =   4920
      Width           =   7220
   End
   Begin VB.ComboBox CmbCom 
      BackColor       =   &H00996666&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   2640
      TabIndex        =   2
      Text            =   "Com1"
      ToolTipText     =   "Com Port selection"
      Top             =   4560
      Width           =   900
   End
   Begin VB.ComboBox CmbSetings 
      BackColor       =   &H00996666&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   3600
      TabIndex        =   1
      Text            =   "19200,n,8,1"
      Top             =   4560
      Width           =   1455
   End
   Begin MSComctlLib.TreeView TVSelectPonsel 
      Height          =   3135
      Left            =   15
      TabIndex        =   0
      Top             =   1080
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   5530
      _Version        =   393217
      Indentation     =   0
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      HotTracking     =   -1  'True
      SingleSel       =   -1  'True
      ImageList       =   "ImgLst1"
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList ImgLst1 
      Left            =   120
      Top             =   6360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":151C3
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":15A9F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSCommLib.MSComm MSCTestCom 
      Left            =   120
      Top             =   4800
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   2
      DTREnable       =   -1  'True
   End
   Begin RichTextLib.RichTextBox RtbListATC 
      Height          =   3255
      Left            =   2640
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   1005
      Visible         =   0   'False
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   5741
      _Version        =   393217
      BorderStyle     =   0
      ScrollBars      =   3
      Appearance      =   0
      RightMargin     =   99999
      TextRTF         =   $"FrmMain.frx":17E83
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image ImgLoadSShot 
      Height          =   2550
      Left            =   440
      Top             =   4600
      Width           =   1710
   End
   Begin VB.Label Status 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Status event or error message :"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2640
      TabIndex        =   29
      Top             =   7200
      Width           =   7215
   End
   Begin VB.Line Line8 
      X1              =   0
      X2              =   9880
      Y1              =   7620
      Y2              =   7620
   End
   Begin VB.Line Line7 
      X1              =   9855
      X2              =   9855
      Y1              =   0
      Y2              =   7800
   End
   Begin VB.Line Line4 
      X1              =   0
      X2              =   9960
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   0
      Y1              =   7800
      Y2              =   0
   End
   Begin VB.Label LblPClose 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Port Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5160
      MouseIcon       =   "FrmMain.frx":17F45
      MousePointer    =   99  'Custom
      TabIndex        =   25
      Top             =   4605
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label LblTLT 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Time Load Text: 0  ms"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   7440
      TabIndex        =   23
      Top             =   4310
      Width           =   2295
   End
   Begin VB.Label LblDesc2 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   255
      Left            =   6000
      TabIndex        =   21
      Top             =   4310
      Width           =   1095
   End
   Begin VB.Label LblATC 
      BackStyle       =   0  'Transparent
      Caption         =   "AT+Command"
      Height          =   255
      Left            =   4440
      TabIndex        =   20
      Top             =   4310
      Width           =   1215
   End
   Begin VB.Label LblSA 
      BackStyle       =   0  'Transparent
      Caption         =   "Access Status "
      Height          =   255
      Left            =   2880
      TabIndex        =   19
      Top             =   4310
      Width           =   1215
   End
   Begin VB.Shape ShpDesc 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000C000&
      Height          =   150
      Left            =   5760
      Top             =   4340
      Width           =   135
   End
   Begin VB.Shape ShpATC 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      Height          =   150
      Left            =   4200
      Top             =   4340
      Width           =   135
   End
   Begin VB.Shape ShpSA 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      FillColor       =   &H00FF0000&
      Height          =   150
      Left            =   2640
      Top             =   4340
      Width           =   135
   End
   Begin VB.Label LblInfoPonsel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "AVACO"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   7200
      Width           =   855
   End
   Begin VB.Shape Shape15 
      BackColor       =   &H00996666&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00996666&
      Height          =   300
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   7200
      Width           =   2385
   End
   Begin VB.Label LblSend 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Send"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8760
      MouseIcon       =   "FrmMain.frx":1824F
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Top             =   4605
      Width           =   975
   End
   Begin VB.Label LblPOpen 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Port Open "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5160
      MouseIcon       =   "FrmMain.frx":18559
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   4605
      Width           =   1215
   End
   Begin VB.Shape ShpSend 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   300
      Left            =   8760
      Shape           =   4  'Rounded Rectangle
      Top             =   4560
      Width           =   1005
   End
   Begin VB.Shape ShpPO 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   300
      Left            =   5160
      Shape           =   4  'Rounded Rectangle
      Top             =   4560
      Width           =   1215
   End
End
Attribute VB_Name = "FrmMain"
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

Dim nATC1, nATC2, nATC3, nATC4, nATC5, nATC6, nATC7, nATC8, nATC9, nATC10, nATC11, nATC12, nATC13 As Node
Dim sFile As String
Dim bCancelFlag As Boolean
Dim DoEv
Dim StringCommand$
Dim EventMessage$
Dim ErrorMessage$
Dim iLoadComPort As Integer
Dim iGetComPort As Integer
Dim iGetDevice As Integer
Dim i As Long

Private Sub Form_Load()
Call Load_Com_port
Call Nodes
TVSelectPonsel.Refresh
End Sub

Private Sub Nodes()
TVSelectPonsel.Nodes.Clear
Set nATC1 = TVSelectPonsel.Nodes.Add(, , "nATC1", "AT+Command For...", 1)
Set nATC2 = TVSelectPonsel.Nodes.Add(, , "nATC2", "C45", 2)
Set nATC3 = TVSelectPonsel.Nodes.Add(, , "nATC3", "C35i-A", 2)
Set nATC4 = TVSelectPonsel.Nodes.Add(, , "nATC4", "C35i-B", 2)
Set nATC5 = TVSelectPonsel.Nodes.Add(, , "nATC5", "G400m", 2)
Set nATC6 = TVSelectPonsel.Nodes.Add(, , "nATC6", "M35i", 2)
Set nATC7 = TVSelectPonsel.Nodes.Add(, , "nATC7", "ME45", 2)
Set nATC8 = TVSelectPonsel.Nodes.Add(, , "nATC8", "S35i-A", 2)
Set nATC9 = TVSelectPonsel.Nodes.Add(, , "nATC9", "S35i-B", 2)
Set nATC10 = TVSelectPonsel.Nodes.Add(, , "nATC10", "S25", 2)
Set nATC11 = TVSelectPonsel.Nodes.Add(, , "nATC11", "S45", 2)
Set nATC12 = TVSelectPonsel.Nodes.Add(, , "nATC12", "SL45-A", 2)
Set nATC13 = TVSelectPonsel.Nodes.Add(, , "nATC13", "Other", 2)
    Set nATC13 = TVSelectPonsel.Nodes.Add("nATC13", tvwChild, , "Readme", 2)
End Sub



Private Sub RtbListATC_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If RtbListATC.SelLength > 0 Then
   FindText$ = RtbListATC.SelText
Else
   FindText$ = ""
End If
End Sub

Private Sub TVSelectPonsel_NodeClick(ByVal Node As MSComctlLib.Node)
On Error GoTo ErrHandler:
RtbListATC.Visible = True
Picture2.Visible = False
Picture3.Visible = False
Select Case TVSelectPonsel.SelectedItem.Text
Case "C45"
    RtbListATC.Text = ""
    sFile = App.Path + "\AT+Command\C45.avc"
    RtbTemp.LoadFile sFile
    Colorize RtbListATC, RtbTemp.Text
    RtbTemp.Text = ""
    LblATCommand.Caption = "AT+Command For : C45"
    LblInfoPonsel.Caption = "C45"
    ImgLoadSShot.Picture = LoadPicture(App.Path & "\Picture\c45.gif")
Case "C35i-A"
    RtbListATC.Text = ""
    sFile = App.Path + "\AT+Command\C35i-A.avc"
    RtbTemp.LoadFile sFile
    Colorize RtbListATC, RtbTemp.Text
    RtbTemp.Text = ""
    LblATCommand.Caption = "AT+Command For : C35i-A"
    LblInfoPonsel.Caption = "C35i-A"
    ImgLoadSShot.Picture = LoadPicture(App.Path & "\Picture\c35i.gif")
Case "C35i-B"
    RtbListATC.Text = ""
    sFile = App.Path + "\AT+Command\C35i-B.avc"
    RtbTemp.LoadFile sFile
    Colorize RtbListATC, RtbTemp.Text
    RtbTemp.Text = ""
    LblATCommand.Caption = "AT+Command For : C35i-B"
    LblInfoPonsel.Caption = "C35i-B"
    ImgLoadSShot.Picture = LoadPicture(App.Path & "\Picture\c35i.gif")
Case "M35i"
    RtbListATC.Text = ""
    sFile = App.Path + "\AT+Command\M35i.avc"
    RtbTemp.LoadFile sFile
    Colorize RtbListATC, RtbTemp.Text
    RtbTemp.Text = ""
    LblATCommand.Caption = "AT+Command For : M35i"
    LblInfoPonsel.Caption = "M35i"
    ImgLoadSShot.Picture = LoadPicture(App.Path & "\Picture\M35.gif")
Case "S35i-A"
    RtbListATC.Text = ""
    sFile = App.Path + "\AT+Command\S35i-A.avc"
    RtbTemp.LoadFile sFile
    Colorize RtbListATC, RtbTemp.Text
    RtbTemp.Text = ""
    LblATCommand.Caption = "AT+Command For : S35i"
    LblInfoPonsel.Caption = "S35i-A"
    ImgLoadSShot.Picture = LoadPicture(App.Path & "\Picture\s35i.gif")
Case "S35i-B"
    RtbListATC.Text = ""
    sFile = App.Path + "\AT+Command\S35i-B.avc"
    RtbTemp.LoadFile sFile
    Colorize RtbListATC, RtbTemp.Text
    RtbTemp.Text = ""
    LblATCommand.Caption = "AT+Command For : S35i-B"
    LblInfoPonsel.Caption = "S35i-B"
    ImgLoadSShot.Picture = LoadPicture(App.Path & "\Picture\s35i.gif")
Case "S25"
    RtbListATC.Text = ""
    sFile = App.Path + "\AT+Command\S25.avc"
    RtbTemp.LoadFile sFile
    Colorize RtbListATC, RtbTemp.Text
    RtbTemp.Text = ""
    LblATCommand.Caption = "AT+Command For : S25"
    LblInfoPonsel.Caption = "S25"
    ImgLoadSShot.Picture = LoadPicture(App.Path & "\Picture\s25.gif")
Case "S45"
    RtbListATC.Text = ""
    sFile = App.Path + "\AT+Command\S45.avc"
    RtbTemp.LoadFile sFile
    Colorize RtbListATC, RtbTemp.Text
    RtbTemp.Text = ""
    LblATCommand.Caption = "AT+Command For : S45"
    LblInfoPonsel.Caption = "S45"
    ImgLoadSShot.Picture = LoadPicture(App.Path & "\picture\S45.gif")
Case "SL45-A"
    RtbListATC.Text = ""
    sFile = App.Path + "\AT+Command\SL45-A.avc"
    RtbTemp.LoadFile sFile
    Colorize RtbListATC, RtbTemp.Text
    RtbTemp.Text = ""
    LblATCommand.Caption = "AT+Command For : SL45-A"
    LblInfoPonsel.Caption = "SL45-A"
    ImgLoadSShot.Picture = LoadPicture(App.Path & "\Picture\sl45.gif")
Case "ME45"
    RtbListATC.Text = ""
    sFile = App.Path + "\AT+Command\StatusAT.avc"
    RtbTemp.LoadFile sFile
    Colorize RtbListATC, RtbTemp.Text
    RtbTemp.Text = ""
    LblATCommand.Caption = "AT+Command For : ME45"
    LblInfoPonsel.Caption = "ME45"
    ImgLoadSShot.Picture = LoadPicture(App.Path & "\Picture\ME45.gif")
Case "G400m"
    RtbListATC.Text = ""
    sFile = App.Path + "\AT+Command\StatusAT.avc"
    RtbTemp.LoadFile sFile
    Colorize RtbListATC, RtbTemp.Text
    RtbTemp.Text = ""
    LblATCommand.Caption = "AT+Command For : G400m"
    LblInfoPonsel.Caption = "G400m"
    ImgLoadSShot.Picture = LoadPicture(App.Path & "\Picture\G400m.gif")
Case "Readme"
    RtbListATC.Text = ""
    sFile = App.Path + "\AT+Command\Readme.avc"
    RtbTemp.LoadFile sFile
    Colorize RtbListATC, RtbTemp.Text
    RtbTemp.Text = ""
    LblATCommand.Caption = "AT+Command For : Readme"
    LblInfoPonsel.Caption = "Readme"
    ImgLoadSShot.Picture = LoadPicture(App.Path & "\Picture\unknown.gif")
End Select
Exit Sub
ErrHandler:
    MsgBox "Error ! " & err.Description, vbCritical, "Error !!"
    Exit Sub
End Sub

Sub Load_Com_port()
For i = 1 To 10
CmbCom.AddItem "COM" + Str$(i)
Next i
CmbCom.ListIndex = 1
CmbSetings.AddItem "19200,n,8,1"
CmbSetings.AddItem "38400,n,8,1"
CmbSetings.AddItem "57600,n,8,1"
End Sub

Private Sub SetComPort()
On Error GoTo ErrHandler
iLoadComPort = InStr(1, CmbCom.Text, "COM", vbTextCompare)
iGetComPort = Mid$(CmbCom.Text, iLoadComPort + 4, 1)
If iLoadComPort = 0 Then iGetComPort = "0"
    iGetDevice = Mid$(CmbCom.Text, iLoadComPort + 4, Len(CmbCom.Text))
    With MSCTestCom
         .Settings = CmbSetings.Text
         .PortOpen = True
    End With
Exit Sub
ErrHandler:
    MsgBox "Error !" & err.Description, vbCritical, "Error !!"
    Exit Sub
End Sub

Private Sub MSCTestCom_OnComm()
    Select Case MSCTestCom.CommEvent
        Case comOverrun
            ErrorMessage$ = " Overrun Error"
        Case comRxOver
            ErrorMessage$ = " Receive Buffer Overflow"
        Case comRxParity
            ErrorMessage$ = " Parity Error"
        Case comCDTO
            ErrorMessage$ = " Carrier Detect Timeout"
        Case comCTSTO
            ErrorMessage$ = " CTS Timeout"
        Case comDCB
            ErrorMessage$ = " Error retrieving DCB"
        Case comDSRTO
            ErrorMessage$ = " DSR Timeout"
        Case comFrame
            ErrorMessage$ = " Framing Error"
        Case comBreak
            ErrorMessage$ = " Break Received"
            MSCTestCom.Break = False
            ErrorMessage$ = ""
        Case comTxFull
            MSCTestCom.OutBufferCount = 0
            ErrorMessage$ = ""
        Case comEvRing
            EventMessage$ = " The Phone is Ringing"
        Case comEvEOF
            EventMessage$ = " End of File Detected"
        Case comEvReceive
        Case comEvSend
        Case comEvCTS
            EventMessage$ = " Clear to send"
        Case comEvDSR
            EventMessage$ = " Change in DSR Detected"
        Case comEvCD
            EventMessage$ = " Carrier Status Toggled"
        Case Else
            ErrorMessage$ = " Unknown error or event"
    End Select
    If Len(EventMessage$) Then
        Status.Caption = "Status Event Or Error Message : " & EventMessage$ & vbCr
    End If
End Sub

'======================================================================================
Private Sub LblPOpen_Click()
LblPOpen.Visible = False
LblPClose.Visible = True
    If MSCTestCom.PortOpen = False Then
       Call SetComPort
    End If
End Sub

Private Sub LblPClose_Click()
On Error GoTo ErrHandler
    If MSCTestCom.PortOpen = True Then
       MSCTestCom.PortOpen = False
    End If
LblPClose.Visible = False
LblPOpen.Visible = True
Exit Sub
ErrHandler:
    MsgBox "Error ! " & err.Description, vbCritical, "Error !!"
    Exit Sub
End Sub

Private Sub LblSend_Click()
On Error GoTo ErrHandler
TxtTestATC.Text = ""
StringCommand$ = CmbATC.Text + vbCr
  
    MSCTestCom.InBufferCount = 0
    MSCTestCom.Output = StringCommand$
    Do
        DoEv = DoEvents()
        If bCancelFlag Then
           bCancelFlag = False
           Exit Do
       End If
       On Error GoTo ErrRepair:
      TxtTestATC.Text = TxtTestATC.Text + MSCTestCom.Input
    Loop
Exit Sub
ErrHandler:
  MsgBox "Error ! " & err.Description, vbCritical, "Error !!"
Exit Sub
ErrRepair:
Exit Sub
End Sub

Private Sub LblExit_Click()
Unload Me
End Sub

Private Sub LblAbout_Click()
RtbListATC.Visible = False
Picture2.Visible = True
Picture3.Visible = False
LblATCommand.Caption = "About"
End Sub

Private Sub LblATCSupport_Click()
RtbListATC.Visible = True
Picture2.Visible = False
Picture3.Visible = False
LblATCommand.Caption = "AT+Command"
If RtbListATC.Text = "" Then
LblATCommand.Caption = "AT+Command"
RtbListATC.Text = "<-- Please Select AT-Command On TreeView"
End If
End Sub

Private Sub LblMinimize_Click()
Me.WindowState = vbMinimized
End Sub

'=======================================================================================

Private Sub scrolllabel2_Timer()
LblInfoPonsel.Left = LblInfoPonsel.Left - 40
If LblInfoPonsel.Left <= Shape15.Left + 50 Then
scrolllabel2.Enabled = False
scrolllabel1.Enabled = True
End If
End Sub

Private Sub scrolllabel1_Timer()
LblInfoPonsel.Left = LblInfoPonsel.Left + 40
If LblInfoPonsel.Left >= Shape15.Left + Shape15.Width - 50 - LblInfoPonsel.Width Then
scrolllabel1.Enabled = False
scrolllabel2.Enabled = True
End If
End Sub

'=======================================================================================

Private Sub Font_bold_False()
LblHome.FontBold = False
Me.LblAbout.FontBold = False
Me.LblATCSupport.FontBold = False
Me.LblMinimize.FontBold = False
Me.LblExit.FontBold = False
Me.LblSend.FontBold = False
Me.LblPClose.FontBold = False
Me.LblPOpen.FontBold = False
End Sub

'======================================================================================
Private Sub LblAbout_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
LblAbout.FontBold = True
LblHome.FontBold = False
Me.LblATCSupport.FontBold = False
Me.LblMinimize.FontBold = False
Me.LblExit.FontBold = False
End Sub

Private Sub LblATCSupport_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
LblATCSupport.FontBold = True
LblHome.FontBold = False
Me.LblAbout.FontBold = False
Me.LblMinimize.FontBold = False
Me.LblExit.FontBold = False
End Sub
Private Sub LblMinimize_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
LblMinimize.FontBold = True
LblHome.FontBold = False
Me.LblAbout.FontBold = False
Me.LblATCSupport.FontBold = False
Me.LblExit.FontBold = False
End Sub

Private Sub LblExit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
LblExit.FontBold = True
LblHome.FontBold = False
Me.LblAbout.FontBold = False
Me.LblATCSupport.FontBold = False
Me.LblMinimize.FontBold = False
End Sub

Private Sub LblHome_Click()
RtbListATC.Visible = False
Picture2.Visible = False
Picture3.Visible = True
LblATCommand.Caption = "Home"
End Sub

Private Sub LblHome_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
LblHome.FontBold = True
Me.LblAbout.FontBold = False
Me.LblATCSupport.FontBold = False
Me.LblMinimize.FontBold = False
Me.LblExit.FontBold = False
End Sub

Private Sub TxtTestATC_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Font_bold_False
End Sub

Private Sub LblPClose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
LblPClose.FontBold = True
End Sub

Private Sub LblSend_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
LblSend.FontBold = True
End Sub

Private Sub LblPOpen_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
LblPOpen.FontBold = True
End Sub

Private Sub Picmenu_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Font_bold_False
End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Font_bold_False
End Sub

Private Sub Picture4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Font_bold_False
End Sub

Private Sub RtbListATC_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then PopupMenu FrmPopup.Menu
End Sub

Private Sub RtbListATC_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Font_bold_False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Font_bold_False
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Font_bold_False
End Sub

Private Sub ImgAvaco_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Font_bold_False
End Sub

Private Sub LblDesc_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormDrag Me
End Sub

Private Sub LblLogo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormDrag Me
End Sub

Private Sub Picmenu_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormDrag Me
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormDrag Me
End Sub

Private Sub LblDesc_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Font_bold_False
End Sub

Private Sub LblAvacoOL_Click()
  ShellExecute Me.hWnd, _
        vbNullString, _
        "http://Avaco-Software.tripod.com", _
        vbNullString, _
        "c:\", _
        SW_SHOWNORMAL
End Sub

Private Sub LblMail_Click()
 ShellExecute Me.hWnd, _
        vbNullString, _
        "mailto:Avaco@9cy.com", _
        vbNullString, _
        "c:\", _
        SW_SHOWNORMAL
End Sub

