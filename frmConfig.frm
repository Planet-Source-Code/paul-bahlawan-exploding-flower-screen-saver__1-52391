VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmConfig 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Exploding Flower Configuration"
   ClientHeight    =   3750
   ClientLeft      =   30
   ClientTop       =   225
   ClientWidth     =   4470
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   4470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   372
      Left            =   2280
      TabIndex        =   3
      Top             =   3000
      Width           =   972
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   372
      Left            =   1200
      TabIndex        =   2
      Top             =   3000
      Width           =   972
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   375
      Index           =   0
      Left            =   720
      TabIndex        =   1
      Top             =   120
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   661
      _Version        =   393216
      LargeChange     =   1
      Min             =   1
      SelStart        =   1
      TickStyle       =   3
      Value           =   1
      TextPosition    =   1
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   375
      Index           =   1
      Left            =   720
      TabIndex        =   8
      Top             =   600
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   661
      _Version        =   393216
      LargeChange     =   1
      Min             =   1
      SelStart        =   1
      TickStyle       =   3
      Value           =   1
      TextPosition    =   1
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   375
      Index           =   2
      Left            =   720
      TabIndex        =   9
      Top             =   1080
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   661
      _Version        =   393216
      LargeChange     =   1
      Min             =   1
      SelStart        =   1
      TickStyle       =   3
      Value           =   1
      TextPosition    =   1
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   375
      Index           =   3
      Left            =   720
      TabIndex        =   10
      Top             =   1560
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   661
      _Version        =   393216
      LargeChange     =   1
      Min             =   1
      SelStart        =   1
      TickStyle       =   3
      Value           =   1
      TextPosition    =   1
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   375
      Index           =   4
      Left            =   720
      TabIndex        =   11
      Top             =   2040
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   661
      _Version        =   393216
      LargeChange     =   1
      Min             =   1
      SelStart        =   1
      TickStyle       =   3
      Value           =   1
      TextPosition    =   1
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   375
      Index           =   5
      Left            =   720
      TabIndex        =   18
      Top             =   2520
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   661
      _Version        =   393216
      LargeChange     =   1
      Min             =   1
      SelStart        =   1
      TickStyle       =   3
      Value           =   1
      TextPosition    =   1
   End
   Begin VB.Label Label1 
      Caption         =   "Slow"
      Height          =   255
      Index           =   11
      Left            =   3840
      TabIndex        =   20
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Fast"
      Height          =   255
      Index           =   10
      Left            =   120
      TabIndex        =   19
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Forky"
      Height          =   255
      Index           =   9
      Left            =   3840
      TabIndex        =   17
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Pointy"
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   16
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Bushy"
      Height          =   255
      Index           =   7
      Left            =   3840
      TabIndex        =   15
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Sparse"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   14
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Fat"
      Height          =   255
      Index           =   5
      Left            =   3840
      TabIndex        =   13
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Skinny"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   12
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Big"
      Height          =   255
      Index           =   3
      Left            =   3840
      TabIndex        =   7
      Top             =   720
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Small"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Many"
      Height          =   255
      Index           =   1
      Left            =   3840
      TabIndex        =   5
      Top             =   240
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Screen Saver by Paul Bahlawan March 2004"
      ForeColor       =   &H80000009&
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   3480
      Width           =   4455
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Few"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   495
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' EXPLODING FLOWER Screen Saver (configure form)
''' By Paul Bahlawan
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

Private Sub cmdCancel_Click()
Unload Me
End Sub

'Save settings in registry
Private Sub cmdOK_Click()
SaveSetting "Exploding_Flower", "Config", "Qty", Slider1(0).Value
SaveSetting "Exploding_Flower", "Config", "Size", Slider1(1).Value
SaveSetting "Exploding_Flower", "Config", "Width", Slider1(2).Value
SaveSetting "Exploding_Flower", "Config", "Petals", Slider1(3).Value
SaveSetting "Exploding_Flower", "Config", "Forked", Slider1(4).Value
SaveSetting "Exploding_Flower", "Config", "Speed", Slider1(5).Value
Unload Me
End Sub

Private Sub Form_Load()
Label2.Caption = Label2.Caption & " (v" & App.Major & "." & App.Minor & ")"
Slider1(0).Value = userQty
Slider1(1).Value = userSize
Slider1(2).Value = userWidth
Slider1(3).Value = userPetals
Slider1(4).Value = userForked
Slider1(5).Value = userSpeed
End Sub

