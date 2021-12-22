VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAddDir 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Adding Wave files..."
   ClientHeight    =   1305
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1305
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton pbCancel 
      Caption         =   "&Cancel"
      Height          =   315
      Left            =   3660
      TabIndex        =   2
      Top             =   900
      Width           =   915
   End
   Begin MSComctlLib.ProgressBar prgBar 
      Height          =   315
      Left            =   60
      TabIndex        =   0
      Top             =   420
      Width           =   4545
      _ExtentX        =   8017
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label lblFile 
      Caption         =   "Adding:"
      Height          =   195
      Left            =   60
      TabIndex        =   1
      Top             =   120
      Width           =   4515
   End
End
Attribute VB_Name = "frmAddDir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private m_sFileMask As String
Private m_objWaveMixer As WaveMixer
Private m_bExitFlag As Boolean

Private Sub Form_Load()
    prgBar.Min = 0
    prgBar.Max = 1
    prgBar.Value = 0
    lblFile.Caption = "Adding:"
    m_bExitFlag = False
End Sub

Public Sub SetProperties( _
           ByVal sFileMask As String, _
           ByVal objWaveMixer As WaveMixer)
    m_sFileMask = sFileMask
    Set m_objWaveMixer = objWaveMixer
End Sub

Public Function AddWaveFiles() _
                As Boolean
    On Error GoTo ErrHandler
    
    Dim nFileCount As Integer
    Dim sFileName As String
    Dim nIndex As Integer
    Dim nCountFrom As Integer
    Dim nCountTo As Integer
    
    'Add files
    nFileCount = GetFileCount(m_sFileMask)
    nCountFrom = m_objWaveMixer.LoadedWaves
    If nFileCount > 0 Then
        prgBar.Min = 0
        prgBar.Max = nFileCount
        prgBar.Value = 0
        
        sFileName = Dir(m_sFileMask)
        While sFileName <> ""
            lblFile.Caption = "Adding: " & Trim(sFileName)
            Me.Refresh
            Call AddFile(sFileName, m_objWaveMixer)
            sFileName = Dir
            nIndex = nIndex + 1
            DoEvents
            If m_bExitFlag Then sFileName = ""
            prgBar.Value = nIndex
        Wend
    End If
    nCountTo = m_objWaveMixer.LoadedWaves
    Call UpdateList(nCountFrom, nCountTo)
    AddWaveFiles = True
    Unload Me
    Exit Function
ErrHandler:
    nCountTo = m_objWaveMixer.LoadedWaves
    Call UpdateList(nCountFrom, nCountTo)
    AddWaveFiles = False
    Unload Me
End Function

Public Sub UpdateList( _
           ByVal nFrom As Integer, _
           ByVal nTo As Integer)
    Dim nIndex As Integer
    For nIndex = nFrom To nTo
        Call frmWaveMixer.AddSoundIndexToList(nIndex)
    Next
End Sub

Private Function AddFile( _
            ByVal sFileName As String, _
            ByRef objWaveMixer As WaveMixer) _
            As Boolean
    On Error GoTo ErrHandler
    Call objWaveMixer.Add(sFileName)
    Call frmWaveMixer.AdjustStatusBar
    AddFile = True
    Exit Function
ErrHandler:
    AddFile = False
End Function

Private Function GetFileCount( _
                 ByVal sFileName As String) _
                 As Integer
    On Error GoTo ErrHandler
    Dim nCount As Integer
    Dim sFile As String
    
    sFile = Dir(sFileName)
    While sFile <> ""
        nCount = nCount + 1
        sFile = Dir
    Wend
ErrHandler:
    GetFileCount = nCount
End Function

Private Sub pbCancel_Click()
    m_bExitFlag = True
End Sub
