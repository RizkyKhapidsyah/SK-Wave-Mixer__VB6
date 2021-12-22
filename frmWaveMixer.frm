VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmWaveMixer 
   Caption         =   "Wave Mixer"
   ClientHeight    =   4800
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   10335
   Icon            =   "frmWaveMixer.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   4800
   ScaleWidth      =   10335
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   600
      Top             =   3960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Open Wave"
      Filter          =   "*.WAV"
   End
   Begin MSComctlLib.ImageList imlSounds 
      Left            =   0
      Top             =   3900
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWaveMixer.frx":000C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   1
      Top             =   4530
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   5874
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   5874
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   5874
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lstSounds 
      Height          =   3855
      Left            =   0
      TabIndex        =   0
      Top             =   60
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   6800
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "imlSounds"
      SmallIcons      =   "imlSounds"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "NAME"
         Text            =   "Name"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Key             =   "ID"
         Text            =   "ID"
         Object.Width           =   706
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Key             =   "SIZE"
         Text            =   "Size"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Key             =   "CHAN"
         Text            =   "Channels"
         Object.Width           =   2028
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Key             =   "DEFFREQ"
         Text            =   "Default Frequency"
         Object.Width           =   2999
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Key             =   "FREQ"
         Text            =   "Current Frequency"
         Object.Width           =   2999
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Key             =   "VOL"
         Text            =   "Attenuation"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   7
         Key             =   "PAN"
         Text            =   "Balance"
         Object.Width           =   2117
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuAddWave 
         Caption         =   "Add Wave..."
      End
      Begin VB.Menu mnuAddDir 
         Caption         =   "Add Directory..."
      End
      Begin VB.Menu mnuNop6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuStopAll 
         Caption         =   "Stop All"
      End
      Begin VB.Menu mnuNop5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClear 
         Caption         =   "&Clear All"
      End
      Begin VB.Menu mnuNop3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuOption 
      Caption         =   "Option"
      Begin VB.Menu mnuPreferences 
         Caption         =   "Preferences"
      End
      Begin VB.Menu mnuNop4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu mnuAdd 
         Caption         =   "Add"
      End
      Begin VB.Menu mnuRemove 
         Caption         =   "Remove"
      End
      Begin VB.Menu mnuNop0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPlay 
         Caption         =   "Play"
      End
      Begin VB.Menu mnuPlayLoop 
         Caption         =   "Play Loop"
      End
      Begin VB.Menu mnuStop 
         Caption         =   "Stop"
      End
      Begin VB.Menu mnuNop1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuReset 
         Caption         =   "Reset Frequency"
      End
      Begin VB.Menu mnuNop2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuProperties 
         Caption         =   "Properties"
      End
   End
End
Attribute VB_Name = "frmWaveMixer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public m_nFrequency As Long
Public m_nPan As Long
Public m_nVolume As Long
Public m_nID As Long
Public m_sName As String
Public m_nLastSortIndex As Integer

Public m_objMixer As New DLSMEDIALib.WaveMixer

Private Sub Form_Load()
    If Not InitMixer Then
        'Unable to init mixer, show error dialog and exit
        Call MsgBox("Unable to initialize the Wave Mixer object." & vbCrLf & _
                    "Please make sure that the 'DLS Media.DLL' and" & vbCrLf & _
                    "DirectX are properly installed on this machine.", _
                    vbOKOnly + vbCritical, _
                    "Critical Error")
        End
    End If
    m_nLastSortIndex = -1
    Call AdjustStatusBar
    
    'Check if a filename was specified from commandline
    If (AddWaveFile(Command) > 0) Then
        Call AddLastSoundToList
    End If
End Sub

Private Function AddWaveFile( _
                 ByVal sFileName As String) _
                 As Long
    On Error GoTo ErrHandler
    Dim nID As Long
    nID = m_objMixer.Add(sFileName)
    AddWaveFile = nID
    Exit Function
ErrHandler:
    AddWaveFile = -1
End Function

Private Function InitMixer() _
                 As Boolean
    On Error GoTo ErrHandler
    Call m_objMixer.Initialize(Me.hWnd)
    InitMixer = True
    Exit Function
ErrHandler:
    InitMixer = False
End Function

Public Function AddLastSoundToList() _
                 As Boolean
    On Error GoTo ErrHandler
    
    'Declare Variables
    Dim nCount As Long
    nCount = m_objMixer.LoadedWaves
    Call AddSoundIndexToList(nCount)
    Call AdjustStatusBar
    AddLastSoundToList = True
    Exit Function
ErrHandler:
    AddLastSoundToList = False
End Function

Public Function AddSoundIndexToList( _
                ByVal nIndex As Integer) _
                As Boolean
    On Error GoTo ErrHandler
    
    'Declare Variables
    Dim nID As Long
    Dim sFileName As String
    Dim nOrgFrequency As Long
    Dim nFrequency As Long
    Dim nPan As Long
    Dim nVolume As Long
    Dim nChannels As Long
    Dim nBits As Long
    Dim nSize As Long
    Dim lstItm As ListItem
    
    Call m_objMixer.GetSoundInfo(nIndex, _
                                 nID, _
                                 sFileName, _
                                 nOrgFrequency, _
                                 nChannels, _
                                 nBits, _
                                 nSize)
    Call m_objMixer.GetFrequency(nID, nFrequency)
    Call m_objMixer.GetPan(nID, nPan)
    Call m_objMixer.GetVolume(nID, nVolume)
    
    'Add this node to listview
    Set lstItm = lstSounds.ListItems.Add(, "K" & nID, sFileName, 1, 1)
    Call lstItm.ListSubItems.Add(, "ID", nID)
    Call lstItm.ListSubItems.Add(, "SIZE", ConvertBytes(nSize))
    Call lstItm.ListSubItems.Add(, "CHAN", ConvertChannels(nChannels))
    Call lstItm.ListSubItems.Add(, "DEFFREQ", ConvertFrequency(nOrgFrequency))
    Call lstItm.ListSubItems.Add(, "FREQ", ConvertFrequency(nFrequency))
    Call lstItm.ListSubItems.Add(, "VOL", ConvertVolume(nVolume))
    Call lstItm.ListSubItems.Add(, "PAN", ConvertPan(nPan))
    
    'Return success
    AddSoundIndexToList = True
    Exit Function
ErrHandler:
    AddSoundIndexToList = False
End Function

Public Sub AdjustStatusBar()
    On Error GoTo ErrHandler
    
    'Adjust statusbar
    StatusBar.Panels(1).Text = m_objMixer.LoadedWaves & " sound(s) loaded"
    StatusBar.Panels(2).Text = ConvertBytes(m_objMixer.UsedMemory) & " in buffer"
    StatusBar.Panels(3).Text = ConvertBytes(m_objMixer.MaxMemory) & " total buffer size"
ErrHandler:
End Sub

Private Sub Form_Resize()
    If Me.Width <= 100 Then Me.Width = 100
    If Me.Height <= 960 Then Me.Height = 960
    
    lstSounds.Left = 0
    lstSounds.Top = 0
    lstSounds.Width = Me.Width - 100
    lstSounds.Height = Me.Height - 960
End Sub

Private Sub lstSounds_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Dim nSortIndex As Integer
    
    nSortIndex = ColumnHeader.Index - 1
    lstSounds.SortKey = nSortIndex
    If m_nLastSortIndex = nSortIndex Then
        If lstSounds.SortOrder = lvwAscending Then
            lstSounds.SortOrder = lvwDescending
        Else
            lstSounds.SortOrder = lvwAscending
        End If
    Else
        lstSounds.SortOrder = lvwAscending
    End If
    lstSounds.Sorted = True
    m_nLastSortIndex = nSortIndex
End Sub

Private Sub lstSounds_DblClick()
    Call mnuProperties_Click
End Sub

Public Function ConvertVolume( _
                ByVal nVolume As Long) _
                As String
    Dim sVolume As String
    nVolume = (nVolume / 100)
    sVolume = nVolume & " dB"
    If nVolume > 0 Then sVolume = "-" & sVolume
    ConvertVolume = sVolume
End Function

Public Function ConvertPan( _
                ByVal nPan As Long) _
                As String
    Dim sBalance As String
    nPan = nPan / 100
    Select Case nPan
        Case 0:         sBalance = "Center"
        Case Is > 0:    sBalance = Abs(nPan) & "% Right"
        Case Is < 0:    sBalance = Abs(nPan) & "% Left"
    End Select
    ConvertPan = sBalance
End Function

Private Function ConvertChannels( _
                 ByVal nChannels As Long) _
                 As String
    Dim sChannels As String
    Select Case nChannels
        Case 1: sChannels = "Mono"
        Case 2: sChannels = "Stereo"
        Case Else: sChannels = Str(nChannels)
    End Select
    ConvertChannels = sChannels
End Function

Private Function ConvertBytes( _
                 ByVal nBytes As Long) _
                 As String
    Dim sBytes As String
    sBytes = FormatNumber(nBytes, 0, vbFalse, vbFalse, vbTrue)
    If sBytes = "" Then sBytes = "0"
    sBytes = sBytes & " bytes"
    ConvertBytes = sBytes
End Function

Public Function ConvertFrequency( _
                ByVal nFrequency As Long) _
                As String
    Dim sFrequency As String
    Dim dFrequency As Double
    dFrequency = nFrequency
    dFrequency = dFrequency / 1000
    sFrequency = FormatNumber(dFrequency, 0, vbFalse, vbFalse, vbTrue)
    ConvertFrequency = sFrequency & " kHz"
End Function

Private Sub lstSounds_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeySpace: Call mnuPlay_Click
        Case vbKeyDelete: Call mnuRemove_Click
        Case vbKeyInsert: Call mnuAdd_Click
        Case vbKeyLeft: Call AdjustFrequency(False)
        Case vbKeyRight: Call AdjustFrequency(True)
        Case 13: Call mnuProperties_Click
    End Select
End Sub

Private Sub lstSounds_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
        PopupMenu mnuPopup
    End If
End Sub

Private Sub mnuAbout_Click()
    Call frmAbout.Show(vbModal, Me)
End Sub

Private Sub mnuAddWave_Click()
    Call mnuAdd_Click
End Sub

Private Sub mnuClear_Click()
    On Error GoTo ErrHandler
    m_objMixer.Clear
    Call AdjustStatusBar
    Call lstSounds.ListItems.Clear
ErrHandler:
End Sub

Private Sub mnuExit_Click()
    On Error GoTo ErrHandler
    Call m_objMixer.Clear
ErrHandler:
    End
End Sub

Private Sub mnuPlay_Click()
    On Error GoTo ErrHandler
    Dim nID As Long
    Dim lstItm As ListItem
    Set lstItm = lstSounds.SelectedItem
    nID = lstItm.ListSubItems("ID").Text
    Call m_objMixer.Play(nID, False)
ErrHandler:
End Sub

Private Sub mnuPlayLoop_Click()
    On Error GoTo ErrHandler
    Dim nID As Long
    Dim lstItm As ListItem
    Set lstItm = lstSounds.SelectedItem
    nID = lstItm.ListSubItems("ID").Text
    Call m_objMixer.Play(nID, True)
ErrHandler:
End Sub

Private Sub mnuPreferences_Click()
    On Error GoTo ErrHandler
    Dim nMaxMemory As Long
    Dim nNewMaxMemory As Long
    Dim sNewMaxMemory As String
    
    nMaxMemory = m_objMixer.MaxMemory
    sNewMaxMemory = InputBox("Specify new buffer size (in bytes)", _
                             "Buffer Size", _
                             nMaxMemory)
    nNewMaxMemory = Val(sNewMaxMemory)
    m_objMixer.MaxMemory = nNewMaxMemory
    Call AdjustStatusBar
    Exit Sub
ErrHandler:
    Call MsgBox("Unable to resize buffer to specified size.", _
                vbOKOnly + vbCritical, _
                "Buffer Size Error")
End Sub

Private Sub mnuProperties_Click()
    On Error GoTo ErrHandler
    Dim nPan As Long
    Dim nVolume As Long
    Dim nFrequency As Long
    Dim sName As String
    Dim nChannels As Long
    Dim nOrgFrequency
    Dim nBits As Long
    
    Dim nID As Long
    Dim lstItm As ListItem
    
    Set lstItm = lstSounds.SelectedItem
    nID = lstItm.ListSubItems("ID").Text
    m_nID = nID
    Call m_objMixer.GetPan(nID, m_nPan)
    Call m_objMixer.GetVolume(nID, m_nVolume)
    Call m_objMixer.GetFrequency(nID, m_nFrequency)
    Call m_objMixer.GetName(nID, m_sName)
    
    Call frmProperties.Show(vbModal, Me)
    
    Call m_objMixer.SetPan(nID, m_nPan)
    Call m_objMixer.SetVolume(nID, m_nVolume)
    Call m_objMixer.SetFrequency(nID, m_nFrequency)
    
    lstItm.ListSubItems("FREQ").Text = ConvertFrequency(m_nFrequency)
    lstItm.ListSubItems("PAN").Text = ConvertPan(m_nPan)
    lstItm.ListSubItems("VOL").Text = ConvertVolume(m_nVolume)
ErrHandler:
End Sub

Private Sub mnuStop_Click()
    On Error GoTo ErrHandler
    Dim nID As Long
    Dim lstItm As ListItem
    Set lstItm = lstSounds.SelectedItem
    nID = lstItm.ListSubItems("ID").Text
    Call m_objMixer.Stop(nID)
ErrHandler:
End Sub

Private Sub mnuRemove_Click()
    On Error GoTo ErrHandler
    Dim nID As Long
    Dim lstItm As ListItem
    Set lstItm = lstSounds.SelectedItem
    nID = lstItm.ListSubItems("ID").Text
    Call m_objMixer.Remove(nID)
    Call lstSounds.ListItems.Remove(lstItm.Index)
    Call AdjustStatusBar
ErrHandler:
End Sub

Private Sub mnuReset_Click()
    On Error GoTo ErrHandler
    Dim nID As Long
    Dim lstItm As ListItem
    Dim nFrequency As Long
    Set lstItm = lstSounds.SelectedItem
    nID = lstItm.ListSubItems("ID").Text
    Call m_objMixer.ResetFrequency(nID)
    Call m_objMixer.GetFrequency(nID, nFrequency)
    lstItm.ListSubItems("FREQ").Text = ConvertFrequency(nFrequency)
ErrHandler:
End Sub

Private Sub mnuAdd_Click()
    On Error GoTo ErrHandler
    Dim sFileName As String
    CommonDialog.DialogTitle = "Add Wave"
    CommonDialog.ShowOpen
    sFileName = CommonDialog.FileName
    If (AddWaveFile(sFileName) > 0) Then
        Call AddLastSoundToList
    End If
ErrHandler:
End Sub

Private Sub mnuAddDir_Click()
    On Error GoTo ErrHandler
    Dim sFileMask As String
    CommonDialog.DialogTitle = "Add Directory"
    CommonDialog.ShowOpen
    sFileMask = CommonDialog.FileName
    sFileMask = GetDirectoryFromFile(sFileMask) & "*.wav"
    Call frmAddDir.SetProperties(sFileMask, m_objMixer)
    Call frmAddDir.Show(vbModeless, Me)
    Call frmAddDir.AddWaveFiles
    Call AdjustStatusBar
ErrHandler:
End Sub

Private Sub AdjustFrequency( _
            ByVal bUp As Boolean)
    On Error GoTo ErrHandler
    Dim nID As Long
    Dim lstItm As ListItem
    Dim nFrequency As Long
    Set lstItm = lstSounds.SelectedItem
    nID = lstItm.ListSubItems("ID").Text
    Call m_objMixer.GetFrequency(nID, nFrequency)
    If (bUp) Then
        nFrequency = nFrequency + 1000
    Else
        nFrequency = nFrequency - 1000
    End If
    If (nFrequency >= 100) And (nFrequency <= 100000) Then
        Call m_objMixer.SetFrequency(nID, nFrequency)
        lstItm.ListSubItems("FREQ").Text = ConvertFrequency(nFrequency)
    End If
ErrHandler:
End Sub

Private Function GetDirectoryFromFile( _
                 ByVal sFileName As String) _
                 As String
    On Error GoTo ErrHandler
    
    'Init Vars
    Dim nIdx As Integer
    Dim sChar As String * 1
    Dim sDirectory As String
    
    sFileName = Trim(sFileName)
    nIdx = Len(sFileName)
    While nIdx > 0
        sChar = Mid$(sFileName, nIdx, 1)
        If (sChar = "/" Or sChar = "\") Then
            sDirectory = Left$(sFileName, nIdx)
            nIdx = 0
        Else
            nIdx = nIdx - 1
        End If
    Wend
    GetDirectoryFromFile = sDirectory
    Exit Function
ErrHandler:
    GetDirectoryFromFile = ""
End Function

Private Sub mnuStopAll_Click()
    On Error GoTo ErrHandler
    Call m_objMixer.StopAll
ErrHandler:
End Sub
