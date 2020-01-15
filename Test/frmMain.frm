VERSION 5.00
Object = "{74D505CD-0C0F-4DAA-9399-AC1915878804}#1.0#0"; "TSSTab.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00FEF6EA&
   Caption         =   "²âÊÔ"
   ClientHeight    =   4995
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8880
   LinkTopic       =   "Form1"
   ScaleHeight     =   4995
   ScaleWidth      =   8880
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin TSSTab.SSTab SSTab1 
      Height          =   4575
      Left            =   120
      TabIndex        =   0
      Top             =   270
      Width           =   8595
      _ExtentX        =   15161
      _ExtentY        =   8070
      TabBackColor    =   16709354
      TabPicture(0)   =   "frmMain.frx":0000
      Tab(0).ControlCount=   1
      Tab(0).Control(1)=   $"frmMain.frx":169C2
      TabPicture(1)   =   "frmMain.frx":169D2
      Tab(1).ControlCount=   6
      Tab(1).Control(1)=   $"frmMain.frx":2D394
      Tab(1).Control(2)=   $"frmMain.frx":2D3A5
      Tab(1).Control(3)=   $"frmMain.frx":2D3B5
      Tab(1).Control(4)=   $"frmMain.frx":2D3C5
      Tab(1).Control(5)=   $"frmMain.frx":2D3D5
      Tab(1).Control(6)=   $"frmMain.frx":2D3E8
      TabPicture(2)   =   "frmMain.frx":2D3F8
      Tab(2).ControlCount=   1
      Tab(2).Control(1)=   $"frmMain.frx":43DBA
      Begin VB.CheckBox Check4 
         Caption         =   "Check4"
         Height          =   405
         Left            =   -73080
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   1440
         Width           =   1155
      End
      Begin VB.FileListBox File1 
         Height          =   2250
         Left            =   -73440
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1200
         Width           =   3615
      End
      Begin TSSTab.SSTab SSTab2 
         Height          =   1935
         Left            =   -74490
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   2130
         Width           =   5025
         _ExtentX        =   8864
         _ExtentY        =   3413
         Tab(0).ControlCount=   1
         Tab(0).Control(1)=   $"frmMain.frx":43DCA
         Tab(1).ControlCount=   1
         Tab(1).Control(1)=   $"frmMain.frx":43DDB
         Tab(2).ControlCount=   1
         Tab(2).Control(1)=   $"frmMain.frx":43DEB
         Begin VB.CommandButton Command1 
            Caption         =   "Command1"
            Height          =   675
            Left            =   -73680
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   780
            Width           =   1215
         End
         Begin VB.CheckBox Check3 
            Caption         =   "Check3"
            Height          =   495
            Left            =   -74400
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   960
            Width           =   1695
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Option1"
            Height          =   495
            Left            =   360
            TabIndex        =   8
            Top             =   1200
            Width           =   1455
         End
      End
      Begin VB.ListBox List1 
         Height          =   1140
         Left            =   -71640
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   840
         Width           =   2265
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Check2"
         Height          =   435
         Left            =   -74490
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   750
         Width           =   1155
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   315
         Left            =   -73080
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   840
         Width           =   1155
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   525
         Left            =   660
         TabIndex        =   1
         Top             =   750
         Width           =   1785
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   345
         Left            =   -74520
         TabIndex        =   4
         Top             =   1440
         Width           =   1155
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    'SSTab1.TabEnabled(1) = False
End Sub

