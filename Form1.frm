VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "detect frequencies in signals"
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5235
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   5235
   StartUpPosition =   3  'Windows-Standard
   Begin MSComctlLib.ListView lvwFreq 
      Height          =   2340
      Left            =   150
      TabIndex        =   16
      Top             =   2325
      Width           =   4890
      _ExtentX        =   8625
      _ExtentY        =   4128
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Frequency (Hz)"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Power (dB)"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton cmdDetect 
      Caption         =   "Detect"
      Height          =   315
      Left            =   4050
      TabIndex        =   12
      Top             =   1875
      Width           =   990
   End
   Begin MSComctlLib.ProgressBar prg 
      Height          =   315
      Left            =   150
      TabIndex        =   11
      Top             =   1875
      Width           =   3840
      _ExtentX        =   6773
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Frame Frame1 
      Caption         =   "DTMF Frequencies"
      Height          =   1515
      Left            =   2550
      TabIndex        =   6
      Top             =   150
      Width           =   2490
      Begin VB.TextBox txtWhiteNoise 
         Height          =   285
         Left            =   1350
         TabIndex        =   14
         Text            =   "10"
         Top             =   750
         Width           =   615
      End
      Begin VB.ComboBox cboFreq 
         Height          =   315
         Left            =   1350
         Style           =   2  'Dropdown-Liste
         TabIndex        =   8
         Top             =   300
         Width           =   915
      End
      Begin VB.Label lblPercent 
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   195
         Index           =   1
         Left            =   2025
         TabIndex        =   15
         Top             =   780
         Width           =   120
      End
      Begin VB.Label lblWhiteNoise 
         AutoSize        =   -1  'True
         Caption         =   "Noise:"
         Height          =   195
         Left            =   225
         TabIndex        =   13
         Top             =   780
         Width           =   450
      End
      Begin VB.Label lblFreq1 
         AutoSize        =   -1  'True
         Caption         =   "Sign:"
         Height          =   195
         Left            =   225
         TabIndex        =   7
         Top             =   350
         Width           =   360
      End
   End
   Begin VB.Frame frmSignal 
      Caption         =   "Signal"
      Height          =   1515
      Left            =   150
      TabIndex        =   0
      Top             =   150
      Width           =   2265
      Begin VB.TextBox txtMS 
         Height          =   285
         Left            =   1125
         TabIndex        =   10
         Text            =   "2000"
         Top             =   1050
         Width           =   690
      End
      Begin VB.TextBox txtSamplerate 
         Height          =   285
         Left            =   1125
         TabIndex        =   2
         Text            =   "11025"
         Top             =   300
         Width           =   690
      End
      Begin VB.TextBox txtVolume 
         Height          =   285
         Left            =   1125
         TabIndex        =   1
         Text            =   "100"
         Top             =   675
         Width           =   690
      End
      Begin VB.Label lblMS 
         AutoSize        =   -1  'True
         Caption         =   "ms"
         Height          =   195
         Left            =   1875
         TabIndex        =   19
         Top             =   1080
         Width           =   195
      End
      Begin VB.Label lblLength 
         AutoSize        =   -1  'True
         Caption         =   "Length:"
         Height          =   195
         Left            =   150
         TabIndex        =   9
         Top             =   1080
         Width           =   540
      End
      Begin VB.Label lblSamplerate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Samplerate:"
         Height          =   195
         Left            =   150
         TabIndex        =   5
         Top             =   330
         Width           =   840
      End
      Begin VB.Label lblVolume 
         AutoSize        =   -1  'True
         Caption         =   "Volume:"
         Height          =   195
         Left            =   150
         TabIndex        =   4
         Top             =   705
         Width           =   570
      End
      Begin VB.Label lblPercent 
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   195
         Index           =   0
         Left            =   1875
         TabIndex        =   3
         Top             =   705
         Width           =   120
      End
   End
   Begin VB.Label lblResult 
      Alignment       =   1  'Rechts
      AutoSize        =   -1  'True
      Caption         =   "Result:"
      Height          =   195
      Left            =   4500
      TabIndex        =   18
      Top             =   4725
      Width           =   495
   End
   Begin VB.Label lblTime 
      AutoSize        =   -1  'True
      Caption         =   "Time:"
      Height          =   195
      Left            =   225
      TabIndex        =   17
      Top             =   4725
      Width           =   390
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' DTMF frequency detection with Goertzel,
' a Fourier Transformation for 1 frequency

' source:
' http://www.musicdsp.org/archive.php?classid=0#107

Private Const PI            As Single = 3.14159265358979
Private Const PI2           As Single = PI * 2
Private Const LOG10         As Single = 0.434294481903251
Private Const NODIVZ        As Single = 0.000000000000001

Private DTMF_F1()           As Single
Private DTMF_F2()           As Single
Private DTMF_NUM()          As Single

' Dual Tone Multi-Frequency (DTMF) Tabelle (ITU-T Q.23)
'
' Hz          1209    1336    1477    1633
' 697          1       2       3       A
' 770          4       5       6       B
' 852          7       8       9       C
' 941          *       0       #       D
'
Private Sub FillDTMFTable()
    Dim i   As Integer
    Dim j   As Integer
    Dim k   As Integer

    ReDim DTMF_F1(3) As Single
    ReDim DTMF_F2(3) As Single
    ReDim DTMF_NUM(15, 3) As Single

    DTMF_F1(0) = 697: DTMF_F2(0) = 1209
    DTMF_F1(1) = 770: DTMF_F2(1) = 1336
    DTMF_F1(2) = 852: DTMF_F2(2) = 1477
    DTMF_F1(3) = 941: DTMF_F2(3) = 1633

    ReDim DTMF_NUM(15, 2) As Single

    For i = 0 To UBound(DTMF_F1)
        For j = 0 To UBound(DTMF_F2)
            DTMF_NUM(k, 0) = DTMF_F1(i)
            DTMF_NUM(k, 1) = DTMF_F2(j)
            k = k + 1
        Next
    Next

    DTMF_NUM(0, 2) = Asc("1")
    DTMF_NUM(1, 2) = Asc("2")
    DTMF_NUM(2, 2) = Asc("3")
    DTMF_NUM(3, 2) = Asc("A")
    DTMF_NUM(4, 2) = Asc("4")
    DTMF_NUM(5, 2) = Asc("5")
    DTMF_NUM(6, 2) = Asc("6")
    DTMF_NUM(7, 2) = Asc("B")
    DTMF_NUM(8, 2) = Asc("7")
    DTMF_NUM(9, 2) = Asc("8")
    DTMF_NUM(10, 2) = Asc("9")
    DTMF_NUM(11, 2) = Asc("C")
    DTMF_NUM(12, 2) = Asc("*")
    DTMF_NUM(13, 2) = Asc("0")
    DTMF_NUM(14, 2) = Asc("#")
    DTMF_NUM(15, 2) = Asc("D")
End Sub

Private Sub cmdDetect_Click()
    Dim lngFreq1    As Long, lngFreq2   As Long
    '
    Dim lngSamples  As Long, lngSR      As Long
    Dim sngVol      As Single, sngWN    As Single
    '
    Dim sngSignal() As Single
    Dim sngS1()     As Single, sngS2()  As Single
    '
    Dim i           As Long
    '
    Dim dB          As Single
    '
    Dim max_f1      As Single, max_f2   As Single
    Dim max_dB1     As Single, max_dB2  As Single
    '
    Dim tmr         As Single

    lngFreq1 = DTMF_NUM(cboFreq.ListIndex, 0)
    lngFreq2 = DTMF_NUM(cboFreq.ListIndex, 1)

    lngSR = Val(txtSamplerate.Text)
    lngSamples = lngSR / 1000 * Val(txtMS.Text)

    ' we only work with normalized values [-1;+1],
    ' therefore the amplifiers should also be in
    ' that range.
    sngVol = Val(txtVolume.Text) / 100
    sngWN = Val(txtWhiteNoise.Text) / 100

    ReDim sngSignal(lngSamples - 1) As Single

    ' generate 2 frequencies which represent the DTMF sign
    sngS1 = MakeTone(lngSR, lngFreq1, lngSamples, sngVol)
    sngS2 = MakeTone(lngSR, lngFreq2, lngSamples, sngVol)

    ' add the signals together to 1 signal
    For i = 0 To lngSamples - 1
        sngSignal(i) = (sngS1(i) + sngS2(i)) * 0.5
    Next

    ' add some random noise
    AddWhiteNoise sngSignal, sngWN, False

    lvwFreq.ListItems.Clear

    prg.Max = (UBound(DTMF_F1) + 1) * 2
    prg.value = 0

    ' we only measure the time needed to
    ' do the analysis, not the signal generation
    tmr = Timer

    For i = 0 To UBound(DTMF_F1)
        dB = power(Goertzel(sngSignal, _
                            lngSamples, _
                            DTMF_F1(i), _
                            lngSR))

        prg.value = prg.value + 1

        With lvwFreq.ListItems.Add(Text:=DTMF_F1(i))
            .SubItems(1) = dB
        End With

        If dB > max_dB1 Then
            max_dB1 = dB
            max_f1 = DTMF_F1(i)
        End If
    Next

    For i = 0 To UBound(DTMF_F2)
        dB = power(Goertzel(sngSignal, _
                            lngSamples, _
                            DTMF_F2(i), _
                            lngSR))

        prg.value = prg.value + 1

        With lvwFreq.ListItems.Add(Text:=DTMF_F2(i))
            .SubItems(1) = dB
        End With

        If dB > max_dB2 Then
            max_dB2 = dB
            max_f2 = DTMF_F2(i)
        End If
    Next

    ' mark the 2 most powerful frequencies found
    With lvwFreq.ListItems
        For i = 1 To .Count
            If .Item(i).Text = CStr(max_f1) Or _
               .Item(i).Text = CStr(max_f2) Then
                .Item(i).ForeColor = vbRed
            End If
        Next
    End With

    ' get the sign from the 2 most powerful frequencies found
    For i = 0 To UBound(DTMF_NUM)
        If DTMF_NUM(i, 0) = max_f1 Then
            If DTMF_NUM(i, 1) = max_f2 Then
                lblResult.Caption = "Result: " & Chr$(DTMF_NUM(i, 2))
                Exit For
            End If
        End If
    Next

    lblTime.Caption = "Time: " & Timer - tmr & " secs"
End Sub

Private Sub Form_Load()
    Dim i   As Long

    FillDTMFTable

    For i = 0 To UBound(DTMF_NUM)
        cboFreq.AddItem Chr$(DTMF_NUM(i, 2))
    Next
    cboFreq.ListIndex = 0
End Sub

Function Goertzel( _
    sngData() As Single, _
    ByVal N As Long, _
    ByVal freq As Single, _
    ByVal sampr As Long _
) As Single

    Dim Skn     As Single
    Dim Skn1    As Single
    Dim Skn2    As Single
    Dim c       As Single
    Dim c2      As Single
    Dim i       As Long

    c = PI2 * freq / sampr
    c2 = Cos(c)

    For i = 0 To N - 1
        Skn2 = Skn1
        Skn1 = Skn
        Skn = 2 * c2 * Skn1 - Skn2 + sngData(i)
    Next

    Goertzel = Skn - Exp(-c) * Skn1
End Function

Function power( _
    ByVal value As Single _
) As Single

    power = 20 * Log(Abs(value) + NODIVZ) * LOG10
End Function

Sub AddWhiteNoise( _
    sngSignal() As Single, _
    ByVal volume As Single, _
    Optional ByVal clip As Boolean _
)

    Dim i           As Long

    For i = 0 To UBound(sngSignal)
        sngSignal(i) = sngSignal(i) + Rnd() * volume

        If clip Then
            If sngSignal(i) > 1 Then
                sngSignal(i) = 1
            ElseIf sngSignal(i) < -1 Then
                sngSignal(i) = -1
            End If
        End If
    Next
End Sub

Function MakeTone( _
    ByVal sampr As Long, _
    ByVal freq As Single, _
    ByVal length As Long, _
    Optional gain As Single = 1 _
) As Single()

    Dim sngTone()   As Single
    Dim i           As Long

    ReDim sngTone(length - 1) As Single

    For i = 0 To length - 1
        sngTone(i) = Sin(freq * PI2 * (i / sampr)) * gain
    Next

    MakeTone = sngTone
End Function
