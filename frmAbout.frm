VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About"
   ClientHeight    =   2730
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4635
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2730
   ScaleWidth      =   4635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton pbOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   315
      Left            =   3540
      TabIndex        =   1
      Top             =   2280
      Width           =   975
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1425
      ItemData        =   "frmAbout.frx":0000
      Left            =   1500
      List            =   "frmAbout.frx":0019
      TabIndex        =   0
      Top             =   600
      Width           =   3075
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1125
      Left            =   120
      Picture         =   "frmAbout.frx":00F9
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1140
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      X1              =   60
      X2              =   4560
      Y1              =   2100
      Y2              =   2100
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000E&
      X1              =   60
      X2              =   4560
      Y1              =   2115
      Y2              =   2115
   End
   Begin VB.Label Label1 
      Caption         =   "Wave Mixer V1.0"
      Height          =   195
      Left            =   1500
      TabIndex        =   3
      Top             =   120
      Width           =   2955
   End
   Begin VB.Label Label2 
      Caption         =   "Developed by Vixit (c) - 2002"
      Height          =   195
      Left            =   1500
      TabIndex        =   2
      Top             =   360
      Width           =   2955
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub pbOk_Click()
    Unload Me
End Sub
