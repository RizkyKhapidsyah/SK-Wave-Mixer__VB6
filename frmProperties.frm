VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProperties 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Sound Properties"
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4980
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   4980
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton pbTest 
      Caption         =   "&Test"
      Height          =   315
      Left            =   60
      TabIndex        =   18
      Top             =   2760
      Width           =   855
   End
   Begin VB.CommandButton pbResetAll 
      Caption         =   "&Reset"
      Height          =   315
      Left            =   1020
      TabIndex        =   1
      Top             =   2760
      Width           =   855
   End
   Begin VB.Frame fraFrequency 
      Caption         =   "Current Frequency"
      ForeColor       =   &H00C00000&
      Height          =   855
      Left            =   60
      TabIndex        =   14
      Top             =   0
      Width           =   4875
      Begin MSComctlLib.Slider sldFrequency 
         Height          =   615
         Left            =   660
         TabIndex        =   5
         Top             =   180
         Width           =   2400
         _ExtentX        =   4233
         _ExtentY        =   1085
         _Version        =   393216
         LargeChange     =   5000
         SmallChange     =   1000
         Max             =   99900
         TickStyle       =   2
         TickFrequency   =   10000
      End
      Begin VB.Label lbl100Hz 
         Caption         =   "100 Hz"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   555
      End
      Begin VB.Label lblFrequency 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   255
         Left            =   3780
         TabIndex        =   17
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lbl100kHz 
         Caption         =   "100 kHz"
         Height          =   195
         Left            =   3060
         TabIndex        =   16
         Top             =   360
         Width           =   675
      End
   End
   Begin VB.Frame fraVolume 
      Caption         =   "Attenuation"
      ForeColor       =   &H00C00000&
      Height          =   855
      Left            =   60
      TabIndex        =   10
      Top             =   900
      Width           =   4875
      Begin MSComctlLib.Slider sldVolume 
         Height          =   615
         Left            =   660
         TabIndex        =   4
         Top             =   180
         Width           =   2400
         _ExtentX        =   4233
         _ExtentY        =   1085
         _Version        =   393216
         Max             =   100
         TickStyle       =   2
         TickFrequency   =   10
      End
      Begin VB.Label lblVolume 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   255
         Left            =   3780
         TabIndex        =   13
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lblMax 
         Caption         =   "-100 dB"
         Height          =   195
         Left            =   3060
         TabIndex        =   12
         Top             =   360
         Width           =   555
      End
      Begin VB.Label lblMin 
         Caption         =   "0 dB"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   435
      End
   End
   Begin VB.Frame fraPan 
      Caption         =   "Balance"
      ForeColor       =   &H00C00000&
      Height          =   855
      Left            =   60
      TabIndex        =   6
      Top             =   1800
      Width           =   4875
      Begin MSComctlLib.Slider sldPan 
         Height          =   555
         Left            =   660
         TabIndex        =   3
         Top             =   180
         Width           =   2400
         _ExtentX        =   4233
         _ExtentY        =   979
         _Version        =   393216
         Min             =   -100
         Max             =   100
         TickStyle       =   2
         TickFrequency   =   20
      End
      Begin VB.Label lblPan 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   255
         Left            =   3780
         TabIndex        =   9
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lblRight 
         Caption         =   "Right"
         Height          =   195
         Left            =   3120
         TabIndex        =   8
         Top             =   360
         Width           =   495
      End
      Begin VB.Label lblLeft 
         Caption         =   "Left"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.CommandButton pbCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   315
      Left            =   4080
      TabIndex        =   2
      Top             =   2760
      Width           =   855
   End
   Begin VB.CommandButton pbOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   315
      Left            =   3120
      TabIndex        =   0
      Top             =   2760
      Width           =   855
   End
End
Attribute VB_Name = "frmProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Me.Caption = "Sound Properties - " & frmWaveMixer.m_sName
    sldPan.Value = frmWaveMixer.m_nPan / 100
    sldVolume.Value = frmWaveMixer.m_nVolume / 100
    sldFrequency.Value = frmWaveMixer.m_nFrequency
    lblFrequency.Caption = frmWaveMixer.ConvertFrequency(sldFrequency.Value)
    lblVolume.Caption = frmWaveMixer.ConvertVolume(sldVolume.Value * 100)
    lblPan.Caption = frmWaveMixer.ConvertPan(sldPan.Value * 100)
End Sub

Private Sub pbCancel_Click()
    Unload Me
End Sub

Private Sub pbOk_Click()
    Call ActivateSettings
    Call Unload(Me)
End Sub

Private Sub ActivateSettings()
    frmWaveMixer.m_nPan = sldPan.Value * 100
    frmWaveMixer.m_nVolume = sldVolume.Value * 100
    frmWaveMixer.m_nFrequency = sldFrequency.Value
End Sub

Private Sub pbResetAll_Click()
    sldPan.Value = frmWaveMixer.m_nPan / 100
    sldVolume.Value = frmWaveMixer.m_nVolume / 100
    sldFrequency.Value = frmWaveMixer.m_nFrequency
    lblFrequency.Caption = frmWaveMixer.ConvertFrequency(sldFrequency.Value)
    lblVolume.Caption = frmWaveMixer.ConvertVolume(sldVolume.Value * 100)
    lblPan.Caption = frmWaveMixer.ConvertPan(sldPan.Value * 100)
End Sub

Private Sub pbTest_Click()
    On Error GoTo ErrHandler
    Dim nCurFrequency As Long
    Dim nCurPan As Long
    Dim nCurVolume As Long
    Dim nID As Long
    
    nID = frmWaveMixer.m_nID
    nCurPan = sldPan.Value * 100
    nCurVolume = sldVolume.Value * 100
    nCurFrequency = sldFrequency.Value
    Call frmWaveMixer.m_objMixer.SetFrequency(nID, nCurFrequency)
    Call frmWaveMixer.m_objMixer.SetPan(nID, nCurPan)
    Call frmWaveMixer.m_objMixer.SetVolume(nID, nCurVolume)
    
    Call frmWaveMixer.m_objMixer.Play(nID, False)
ErrHandler:
End Sub

Private Sub sldFrequency_Scroll()
    lblFrequency.Caption = frmWaveMixer.ConvertFrequency(sldFrequency.Value)
End Sub

Private Sub sldPan_Scroll()
    lblPan.Caption = frmWaveMixer.ConvertPan(sldPan.Value * 100)
End Sub

Private Sub sldVolume_Scroll()
    lblVolume.Caption = frmWaveMixer.ConvertVolume(sldVolume.Value * 100)
End Sub
