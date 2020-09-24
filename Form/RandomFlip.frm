VERSION 5.00
Begin VB.Form RandomFlip 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Instability Monitor"
   ClientHeight    =   5385
   ClientLeft      =   1440
   ClientTop       =   2865
   ClientWidth     =   7950
   DrawMode        =   8  'Xor Pen
   FillColor       =   &H00E0E0E0&
   Icon            =   "RandomFlip.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   359
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   530
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox w 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4920
      TabIndex        =   46
      Text            =   "0"
      Top             =   5040
      Width           =   975
   End
   Begin VB.TextBox mo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5880
      TabIndex        =   45
      Text            =   "0"
      Top             =   5040
      Width           =   975
   End
   Begin VB.TextBox h 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2640
      TabIndex        =   44
      Text            =   "0"
      Top             =   5040
      Width           =   975
   End
   Begin VB.TextBox d 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3720
      TabIndex        =   43
      Text            =   "0"
      Top             =   5040
      Width           =   975
   End
   Begin VB.TextBox s 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   480
      TabIndex        =   42
      Text            =   "0"
      Top             =   5040
      Width           =   975
   End
   Begin VB.TextBox m 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1560
      TabIndex        =   41
      Text            =   "0"
      Top             =   5040
      Width           =   975
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000009&
      Caption         =   "Total Flips: 0"
      Height          =   3480
      Left            =   120
      TabIndex        =   7
      Top             =   0
      Width           =   7545
      Begin VB.CommandButton Command1 
         BackColor       =   &H80000009&
         Height          =   495
         Left            =   120
         Picture         =   "RandomFlip.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   2880
         Width           =   495
      End
      Begin VB.Label Chi 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   1080
         TabIndex        =   40
         Top             =   3120
         Width           =   45
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "="
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         TabIndex        =   39
         Top             =   3000
         Width           =   375
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Past Year Instability:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3120
         TabIndex        =   31
         Top             =   3120
         Width           =   2520
      End
      Begin VB.Label Yri 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   360
         Left            =   5850
         TabIndex        =   30
         Top             =   3000
         Width           =   1065
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Past Minute Instability:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3120
         TabIndex        =   29
         Top             =   720
         Width           =   2760
      End
      Begin VB.Label Mini 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   360
         Left            =   5850
         TabIndex        =   28
         Top             =   720
         Width           =   1065
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Delta: 0"
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   6
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   120
         Left            =   150
         TabIndex        =   25
         Top             =   585
         Width           =   720
      End
      Begin VB.Line Line1 
         X1              =   3045
         X2              =   3045
         Y1              =   360
         Y2              =   3240
      End
      Begin VB.Label Moni 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   360
         Left            =   5850
         TabIndex        =   22
         Top             =   2640
         Width           =   1065
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Past 30d Instability:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3120
         TabIndex        =   21
         Top             =   2640
         Width           =   2400
      End
      Begin VB.Label Wki 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   360
         Left            =   5850
         TabIndex        =   18
         Top             =   2160
         Width           =   1065
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Past Week Instability:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3120
         TabIndex        =   17
         Top             =   2160
         Width           =   2610
      End
      Begin VB.Label Di 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   360
         Left            =   5850
         TabIndex        =   16
         Top             =   1680
         Width           =   1065
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Past Day Instability:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3120
         TabIndex        =   15
         Top             =   1680
         Width           =   2400
      End
      Begin VB.Label Hri 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   360
         Left            =   5850
         TabIndex        =   14
         Top             =   1200
         Width           =   1065
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Past Hour Instability:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3120
         TabIndex        =   13
         Top             =   1200
         Width           =   2520
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Heads: 0"
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   6
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   120
         Left            =   150
         TabIndex        =   12
         Top             =   225
         Width           =   720
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tails: 0"
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   6
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   120
         Left            =   150
         TabIndex        =   11
         Top             =   405
         Width           =   720
      End
      Begin VB.Image Tails 
         Height          =   2355
         Left            =   360
         Picture         =   "RandomFlip.frx":1194
         Top             =   660
         Visible         =   0   'False
         Width           =   2280
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Instability:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3120
         TabIndex        =   9
         Top             =   240
         Width           =   1920
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   360
         Left            =   5850
         TabIndex        =   8
         Top             =   240
         Width           =   1065
      End
      Begin VB.Image Heads 
         Height          =   2325
         Left            =   360
         Picture         =   "RandomFlip.frx":5878
         Top             =   675
         Width           =   2265
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6960
      Top             =   4800
   End
   Begin VB.ListBox Minutesi 
      Appearance      =   0  'Flat
      Height          =   1200
      Left            =   1560
      TabIndex        =   27
      Top             =   3720
      Width           =   1035
   End
   Begin VB.ListBox Secondsi 
      Appearance      =   0  'Flat
      Height          =   1200
      Left            =   435
      TabIndex        =   26
      Top             =   3720
      Width           =   1035
   End
   Begin VB.ListBox Monthi 
      Appearance      =   0  'Flat
      Height          =   1200
      Left            =   5880
      TabIndex        =   24
      Top             =   3720
      Width           =   1035
   End
   Begin VB.ListBox Weeki 
      Appearance      =   0  'Flat
      Height          =   1200
      Left            =   4800
      TabIndex        =   23
      Top             =   3720
      Width           =   1035
   End
   Begin VB.ListBox Dayi 
      Appearance      =   0  'Flat
      Height          =   1200
      Left            =   3720
      TabIndex        =   20
      Top             =   3720
      Width           =   1035
   End
   Begin VB.ListBox Houri 
      Appearance      =   0  'Flat
      Height          =   1200
      Left            =   2640
      TabIndex        =   19
      Top             =   3720
      Width           =   1035
   End
   Begin VB.PictureBox HyPiano 
      Height          =   372
      Left            =   7080
      Picture         =   "RandomFlip.frx":9AF5
      ScaleHeight     =   315
      ScaleWidth      =   915
      TabIndex        =   5
      Top             =   3000
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.Timer QuitTimer 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4680
      Top             =   3720
   End
   Begin VB.PictureBox Scope 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   2856
      Left            =   7680
      ScaleHeight     =   238
      ScaleMode       =   0  'User
      ScaleWidth      =   521.377
      TabIndex        =   1
      Top             =   840
      Visible         =   0   'False
      Width           =   6672
      Begin VB.ListBox List1 
         Height          =   1425
         Left            =   4785
         Sorted          =   -1  'True
         TabIndex        =   6
         Top             =   1440
         Width           =   1485
      End
   End
   Begin VB.CommandButton StopButton 
      BackColor       =   &H0000C000&
      Caption         =   "&Stop"
      Enabled         =   0   'False
      Height          =   336
      Left            =   3165
      TabIndex        =   2
      Top             =   6015
      Visible         =   0   'False
      Width           =   984
   End
   Begin VB.ComboBox DevicesBox 
      Height          =   315
      Left            =   7800
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1320
      Width           =   3108
   End
   Begin VB.PictureBox ScopeBuff 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C000C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000002&
      Height          =   336
      Left            =   5640
      ScaleHeight     =   22
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   78
      TabIndex        =   3
      Top             =   3840
      Visible         =   0   'False
      Width           =   1176
   End
   Begin VB.CommandButton StartButton 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Start"
      Height          =   336
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   600
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Months"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   6000
      TabIndex        =   37
      Top             =   3480
      Width           =   735
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Weeks"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   4920
      TabIndex        =   36
      Top             =   3480
      Width           =   855
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Days"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   3840
      TabIndex        =   35
      Top             =   3480
      Width           =   735
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Hours"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   2640
      TabIndex        =   34
      Top             =   3480
      Width           =   855
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Minutes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   1680
      TabIndex        =   33
      Top             =   3480
      Width           =   735
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Seconds"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   32
      Top             =   3480
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "22,050 Hz"
      ForeColor       =   &H80000004&
      Height          =   372
      Left            =   5640
      TabIndex        =   4
      Top             =   0
      Width           =   972
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuNew 
         Caption         =   "New"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "Open"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save"
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save As"
      End
      Begin VB.Menu dygjdgyj 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
      Begin VB.Menu mnuStart 
         Caption         =   "Start"
      End
      Begin VB.Menu mnuPause 
         Caption         =   "Pause/Stop"
      End
      Begin VB.Menu mnuReset 
         Caption         =   "Reset"
      End
      Begin VB.Menu dfgsdfgd 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSimulate 
         Caption         =   "Simulate"
      End
   End
   Begin VB.Menu view 
      Caption         =   "View"
      Begin VB.Menu mnuExpanded 
         Caption         =   "Collapsed View"
      End
      Begin VB.Menu mnuInstability 
         Caption         =   "Instability Stats"
      End
      Begin VB.Menu mnuSequential 
         Caption         =   "Sequential Data"
      End
      Begin VB.Menu mnuWe 
         Caption         =   "-"
      End
      Begin VB.Menu mnuChisq 
         Caption         =   "Chi Square Table"
      End
   End
   Begin VB.Menu mnuData 
      Caption         =   "Data"
      Begin VB.Menu mnuGraphic 
         Caption         =   "Graphic Data"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
      Begin VB.Menu mnuWebpage 
         Caption         =   "Webpage"
      End
   End
End
Attribute VB_Name = "RandomFlip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Active As Boolean
Private DevHandle As Long 'Handle of the open audio device
Private Visualizing As Boolean
Private Divisor As Long
Private ScopeHeight As Long 'Saves time because hitting up a Long is faster
                            'than a property.
Private Type WaveFormatEx
    FormatTag As Integer
    Channels As Integer
    SamplesPerSec As Long
    AvgBytesPerSec As Long
    BlockAlign As Integer
    BitsPerSample As Integer
    ExtraDataSize As Integer
End Type
Private Type WaveHdr
    lpData As Long
    dwBufferLength As Long
    dwBytesRecorded As Long
    dwUser As Long
    dwFlags As Long
    dwLoops As Long
    lpNext As Long 'wavehdr_tag
    Reserved As Long
End Type
Private Type WaveInCaps
    ManufacturerID As Integer      'wMid
    ProductID As Integer       'wPid
    DriverVersion As Long       'MMVERSIONS vDriverVersion
    ProductName(1 To 32) As Byte 'szPname[MAXPNAMELEN]
    Formats As Long
    Channels As Integer
    Reserved As Integer
End Type
Private Const WAVE_INVALIDFORMAT = &H0&                 '/* invalid format */
Private Const WAVE_FORMAT_1M08 = &H1&                   '/* 11.025 kHz, Mono,   8-bit
Private Const WAVE_FORMAT_1S08 = &H2&                   '/* 11.025 kHz, Stereo, 8-bit
Private Const WAVE_FORMAT_1M16 = &H4&                   '/* 11.025 kHz, Mono,   16-bit
Private Const WAVE_FORMAT_1S16 = &H8&                   '/* 11.025 kHz, Stereo, 16-bit
Private Const WAVE_FORMAT_2M08 = &H10&                  '/* 22.05  kHz, Mono,   8-bit
Private Const WAVE_FORMAT_2S08 = &H20&                  '/* 22.05  kHz, Stereo, 8-bit
Private Const WAVE_FORMAT_2M16 = &H40&                  '/* 22.05  kHz, Mono,   16-bit
Private Const WAVE_FORMAT_2S16 = &H80&                  '/* 22.05  kHz, Stereo, 16-bit
Private Const WAVE_FORMAT_4M08 = &H100&                 '/* 44.1   kHz, Mono,   8-bit
Private Const WAVE_FORMAT_4S08 = &H200&                 '/* 44.1   kHz, Stereo, 8-bit
Private Const WAVE_FORMAT_4M16 = &H400&                 '/* 44.1   kHz, Mono,   16-bit
Private Const WAVE_FORMAT_4S16 = &H800&                 '/* 44.1   kHz, Stereo, 16-bit
Private Const WAVE_FORMAT_PCM = 1
Private Const WHDR_DONE = &H1&              '/* done bit */
Private Const WHDR_PREPARED = &H2&          '/* set if this header has been prepared */
Private Const WHDR_BEGINLOOP = &H4&         '/* loop start block */
Private Const WHDR_ENDLOOP = &H8&           '/* loop end block */
Private Const WHDR_INQUEUE = &H10&          '/* reserved for driver */
Private Const WIM_OPEN = &H3BE
Private Const WIM_CLOSE = &H3BF
Private Const WIM_DATA = &H3C0
Private Declare Function waveInAddBuffer Lib "winmm" (ByVal InputDeviceHandle As Long, ByVal WaveHdrPointer As Long, ByVal WaveHdrStructSize As Long) As Long
Private Declare Function waveInPrepareHeader Lib "winmm" (ByVal InputDeviceHandle As Long, ByVal WaveHdrPointer As Long, ByVal WaveHdrStructSize As Long) As Long
Private Declare Function waveInUnprepareHeader Lib "winmm" (ByVal InputDeviceHandle As Long, ByVal WaveHdrPointer As Long, ByVal WaveHdrStructSize As Long) As Long
Private Declare Function waveInGetNumDevs Lib "winmm" () As Long
Private Declare Function waveInGetDevCaps Lib "winmm" Alias "waveInGetDevCapsA" (ByVal uDeviceID As Long, ByVal WaveInCapsPointer As Long, ByVal WaveInCapsStructSize As Long) As Long
Private Declare Function waveInOpen Lib "winmm" (WaveDeviceInputHandle As Long, ByVal WhichDevice As Long, ByVal WaveFormatExPointer As Long, ByVal CallBack As Long, ByVal CallBackInstance As Long, ByVal FLAGS As Long) As Long
Private Declare Function waveInClose Lib "winmm" (ByVal WaveDeviceInputHandle As Long) As Long
Private Declare Function waveInStart Lib "winmm" (ByVal WaveDeviceInputHandle As Long) As Long
Private Declare Function waveInReset Lib "winmm" (ByVal WaveDeviceInputHandle As Long) As Long
Private Declare Function waveInStop Lib "winmm" (ByVal WaveDeviceInputHandle As Long) As Long
Dim maxvol As Long, Hz As Long, oscila As Long
Dim HzColor As Long, xMax As Integer, HzTip As Long
Dim Instability As Integer: Dim TotalFlips As Long: Dim Head As Long: Dim Tail As Long
Dim Timing As Long: Dim Timing1 As Long: Dim Timing2 As Long
Dim Headz As Long: Dim Tailz As Long
Dim Simulate As Boolean
Sub InitDevices()
    'Fill the DevicesBox box with all the compatible audio input devices
    'Bail if there are none.
    Dim Caps As WaveInCaps, Which As Long
    DevicesBox.Clear
    For Which = 0 To waveInGetNumDevs - 1
        Call waveInGetDevCaps(Which, VarPtr(Caps), Len(Caps))
        If Caps.Formats And WAVE_FORMAT_4M16 Then '16-bit mono devices
            Call DevicesBox.AddItem(StrConv(Caps.ProductName, vbUnicode), Which)
        End If
    Next
    If DevicesBox.ListCount = 0 Then
        MsgBox "You have no audio input devices!", vbCritical, "Ack!"
        End 'Ewww!  End!  Bad me!
    End If
    DevicesBox.ListIndex = 0
End Sub


Private Sub Command1_Click()
Load ChiSquared
ChiSquared.Show
End Sub

Private Sub Form_Load()
On Error Resume Next
Me.Height = 4290
Me.Width = 3180
Frame1.Width = 191
Dim PresentInstability As String
Active = False
Open App.Path & "\PresentInstability" For Input As #1
    Line Input #1, PresentInstability
Close #1
Instability = Val(PresentInstability)
Open App.Path & "\TotalFlips" For Input As #1
    Line Input #1, PresentInstability
Close #1
TotalFlips = Val(PresentInstability)
Open App.Path & "\Head" For Input As #1
    Line Input #1, PresentInstability
Close #1
Head = Val(PresentInstability)
Label7.Caption = "Heads: " & Head
Open App.Path & "\Tail" For Input As #1
    Line Input #1, PresentInstability
Close #1
Tail = Val(PresentInstability)
Label5.Caption = "Tails: " & Tail
    'HyPiano.Picture = LoadPicture(App.Path & "\hypiano.bmp")
    Call InitDevices 'Fill the DevicesBox
    Call DoReverse   'Pre-calculate these
    'Set the double buffer to match the display
    ScopeBuff.Width = Scope.ScaleWidth
    ScopeBuff.Height = Scope.ScaleHeight
    ScopeBuff.BackColor = Scope.BackColor
    ScopeHeight = Scope.Height
    mpiano_Click
    
End Sub
Private Sub Form_Unload(Cancel As Integer)
        Call DoStop
        Cancel = 1
    QuitTimer.Enabled = False
    Timer1.Enabled = False
    Unload Me
    Set RandomFlip = Nothing
    End
End Sub

Private Sub Label1_Change()
On Error Resume Next
Dim Hz As Single, Note As Integer, i As Long
Hz = Val(Label1.Caption)
Note = Int(12 * Log(Hz / 440) / Log(2))
Note = Note + 69
'MsgBox Note
'For I = 12 To 84
    'If Note >= I And Note < I + 1 Then
        'frmMain.SendMIDIOut (NOTE_ON + 1), Note, 120, 0, MIDIOUT_QUEUE
        'Delay (200)
    'End If
'Next
End Sub

Private Sub mnuAbout_Click()
Load frmAbout
frmAbout.Show
End Sub

Private Sub mnuChisq_Click()
Command1_Click
End Sub

Private Sub mnuExit_Click()
Dim intsave As Integer
    intsave = MsgBox("Do you want to Quit?", _
                     vbYesNoCancel + vbExclamation)
    Select Case intsave
      Case vbYes
        Unload Me
    End Select


End Sub

Private Sub mnuExpanded_Click()
    Me.Height = 4290
    Me.Width = 3180
    Frame1.Width = 191
    mnuInstability.Checked = False
    mnuSequential.Checked = False
End Sub

Private Sub mnuGraphic_Click()
Load Chart
Chart.Show
End Sub

Private Sub mnuInstability_Click()
If mnuInstability.Checked = True Then
    mnuInstability.Checked = False
    Me.Width = 3180
    Frame1.Width = 191
Else
    mnuInstability.Checked = True
    Me.Width = 7425
    Frame1.Width = 476
End If
    mnuSequential.Checked = False

End Sub

Private Sub mnuPause_Click()
    Timer1.Enabled = False

End Sub

Private Sub mnuReset_Click()
On Error Resume Next
Dim intsave As Integer
    intsave = MsgBox("Do you want to Reset all Values and Files?", _
                     vbYesNoCancel + vbExclamation)
    Select Case intsave
      Case vbYes
        Open App.Path & "\Instability" For Output As #1: Close #1
        Open App.Path & "\PresentInstability" For Output As #1: Close #1
        Open App.Path & "\TotalFlips" For Output As #1: Close #1
        Kill App.Path & "\15Minutes"
        Open App.Path & "\Outdata" For Output As #1: Close #1
        Open App.Path & "\Nostradamoose.dat" For Output As #1: Close #1
        Open App.Path & "\Nostradamoose.tim" For Output As #1: Close #1
        Open App.Path & "\CompleteRecord" For Output As #1: Close #1
        Instability = 0
        Label3.Caption = 0
        Label14.Caption = "Delta: 0"
        Label3.Caption = "-"
        Mini.Caption = "-"
        Hri.Caption = "-"
        Di.Caption = "-"
        Moni.Caption = "-"
        Yri.Caption = "-"
        Frame1.Caption = "Total Flips = 0"
        Label5.Caption = "Tails: 0"
        Tail = 0
        Label7.Caption = "Heads: 0"
        Head = 0
        TotalFlips = 0
    End Select


End Sub

Private Sub mnuSequential_Click()
If mnuSequential.Checked = True Then
    mnuSequential.Checked = False
    Me.Height = 4290
    Me.Width = 3180
    Frame1.Width = 191
Else
    mnuSequential.Checked = True
    Me.Height = 6165
    Me.Width = 7425
    Frame1.Width = 476
End If
    mnuInstability.Checked = False
End Sub

Private Sub mnuSimulate_Click()
mnuPause_Click
If mnuSimulate.Checked = True Then
    mnuSimulate.Checked = False
    Simulate = False
    mnuReset_Click
Else
    mnuSimulate.Checked = True
    Simulate = True
    mnuReset_Click
End If
mnuStart_Click
End Sub

Private Sub mnuStart_Click()
Timer1.Enabled = True
End Sub

Private Sub mpiano_Click()
'If mpiano.Checked = False Then mpiano.Checked = True Else mpiano.Checked = False
End Sub
Private Sub QuitTimer_Timer()
    Unload Me
End Sub
Private Sub Scope_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Y >= ScaleHeight / 3 And Y < ScaleHeight / 3 * 2 Then Scope.ToolTipText = "516.84 to 5,125.33Hz": Exit Sub
  If Y >= ScaleHeight / 3 * 2 Then Scope.ToolTipText = "43.07 to 473.77 Hz": Exit Sub
  Scope.ToolTipText = "5,168 to 22K Hz"
End Sub
Private Sub StartButton_Click()

    Static WaveFormat As WaveFormatEx
    With WaveFormat
        .FormatTag = WAVE_FORMAT_PCM
        .Channels = 1
        .SamplesPerSec = 44100
        .BitsPerSample = 16
        .BlockAlign = (.Channels * .BitsPerSample) \ 8
        .AvgBytesPerSec = .BlockAlign * .SamplesPerSec
        .ExtraDataSize = 0
    End With
    maxvol = waveInOpen(DevHandle, DevicesBox.ListIndex, VarPtr(WaveFormat), 0, 0, 0)
    If DevHandle = 0 Then
        Call MsgBox("Wave input device didn't open!", vbExclamation, "Ack!")
        Exit Sub
    End If
    Call waveInStart(DevHandle)
    StopButton.Caption = "&Stop"
    StopButton.Enabled = True
    StartButton.Enabled = False
    DevicesBox.Enabled = False
    Call Visualize
End Sub
Private Sub StopButton_Click()
Close #1
    Call DoStop
    If StopButton.Caption = "&Stop" Then
    StopButton.Caption = "Exit"
    Exit Sub
    End If
    QuitTimer = True
End Sub
Private Sub DoStop()
    Call waveInReset(DevHandle)
    Call waveInClose(DevHandle)
    DevHandle = 0
    'StopButton.Enabled = False
    StartButton.Enabled = True
    DevicesBox.Enabled = True
End Sub
Private Sub Visualize()
On Error Resume Next
Dim MyValue As Variant: Dim i As Long: Dim TempVal As Single: Dim ValVal As Long
Dim HeadFlag As Integer: Dim TailFlag As Integer: Dim Minutes15 As Integer
Dim Chisq As Single
    
    '                    Original code to get the data and process the FFT
      Static X As Long
      Static Wave As WaveHdr
      Static InData(0 To NumSamples - 1) As Integer
      Static OutData(0 To NumSamples - 1) As Single
TopHere:
If Simulate = True Then GoTo Here1
              Wave.lpData = VarPtr(InData(0))
              Wave.dwBufferLength = NumSamples
              Wave.dwFlags = 0
              Call waveInPrepareHeader(DevHandle, VarPtr(Wave), Len(Wave))
              Call waveInAddBuffer(DevHandle, VarPtr(Wave), Len(Wave))
              Call waveInUnprepareHeader(DevHandle, VarPtr(Wave), Len(Wave))
            Do
                'Just wait for the blocks to be done or the device to close
            Loop Until ((Wave.dwFlags And WHDR_DONE) = WHDR_DONE) Or DevHandle = 0
            
            Call waveInUnprepareHeader(DevHandle, VarPtr(Wave), Len(Wave))
            Call FFTAudio(InData, OutData)
        
              
              Dim c As Double, LowMidHig
            For X = 1 To 511
               ScopeBuff.DrawWidth = 1
               ScopeBuff.PSet (X, ScopeHeight / 3 - (InData(X) / 500)), oscila ' oscilloscope
            If Abs(OutData(X)) > maxvol Then    'And Active = True Then
               Active = False
               maxvol = Abs(OutData(X))
               Hz = Int(44100 * X) / 1024
               Label1.Caption = Hz
               Active = False
               List1.AddItem Hz     'maxvol / 100
               HzColor = vbRed + Hz
               LowMidHig = ScopeHeight
               xMax = X
               Exit For
            End If
            Next
            X = xMax
                maxvol = 0
                'MsgBox OutData(X)
                If OutData(X) = 0 Then GoTo TopHere
Here1:
            If Simulate = False Then
                 Randomize Abs(Int(OutData(X)))
                 Timer1.Interval = 3000
            Else
                Randomize
                Timer1.Interval = 1
                OutData(X) = Timer
            End If
                 MyValue = Int((2 * Rnd) + 1)
                  If MyValue = 2 Then
                      Heads.Visible = True
                      Tails.Visible = False
                      Me.Caption = "Head - Seed: " & Int(Abs(OutData(X)))
                      Instability = Instability + 1
                      Head = Head + 1
                      Headz = Headz + 1
                      Open App.Path & "\Head" For Output As #1
                          Print #1, Head
                      Close #1
                      Open App.Path & "\CompleteRecord" For Append As #1
                            Print #1, "Heads : " & Now
                      Close #1
                      Label7.Caption = "Heads: " & Head
                      HeadFlag = 100
                      TailFlag = 0
                  Else
                      Heads.Visible = False
                      Tails.Visible = True
                      ValVal = Int(Abs(OutData(X)))
                      If ValVal = 0 Then Visualize
                      Me.Caption = "Tail - Seed: " & Int(Abs(OutData(X)))
                      Open App.Path & "\Outdata" For Append As #1: Print #1, ValVal: Close #1
                      Instability = Instability - 1
                      Tail = Tail + 1
                      Tailz = Tailz + 1
                      Open App.Path & "\Tail" For Output As #1
                          Print #1, Tail
                      Close #1
                      Open App.Path & "\CompleteRecord" For Append As #1
                            Print #1, "Tails : " & Now
                      Close #1
                      
                      Label5.Caption = "Tails: " & Tail
                      TailFlag = -100
                      HeadFlag = 0
                  End If
                  Open App.Path & "\Instability" For Append As #1
                      Print #1, Instability
                  Close #1
                  Open App.Path & "\PresentInstability" For Output As #1
                      Print #1, Instability
                  Close #1
                  Label3.Caption = Abs((Int((Head / Tail) * 1000) / 10) - 100) & " %" 'Instability
                  Label14.Caption = "Delta: " & Instability
                  If TailFlag = 0 Then
                    Secondsi.AddItem HeadFlag   'Abs((Int((Headz / Tailz) * 1000) / 10) - 100)
                  Else
                    Secondsi.AddItem TailFlag   'Abs((Int((Headz / Tailz) * 1000) / 10) - 100)
                  End If
                  TotalFlips = TotalFlips + 1
                  Open App.Path & "\TotalFlips" For Output As #1
                      Print #1, TotalFlips
                  Close #1
                  Chisq = (((Head - TotalFlips / 2) ^ 2) / (TotalFlips / 2)) + (((Tail - TotalFlips / 2) ^ 2) / (TotalFlips / 2))
                  Chi.Caption = Chisq
Frame1.Caption = "Total Flips = " & TotalFlips
Timing = Timing + 1
s.Text = Timing
TempVal = 0
If Timing = 60 Then
Headz = 0
Tailz = 0
    Timing = 0
    For i = 0 To Secondsi.ListCount - 1
        TempVal = TempVal + Val(Secondsi.List(i)) * 10
    Next
    TempVal = Abs(Int((TempVal / i)) / 10)
    Minutesi.AddItem TempVal
    m.Text = Minutesi.ListCount
    Mini.Caption = TempVal & " %"
    Open App.Path & "\Nostradamoose.tim" For Append As #2
    Open App.Path & "\Nostradamoose.dat" For Append As #1
        Print #1, TempVal
        Print #2, Format(Now, "ddmmyyhhmmss")
    Close #1
    Close #2
    Secondsi.Clear
    Minutes15 = Minutesi.ListCount - 1
    If (Minutes15) / 15 = Int((Minutes15) / 15) Then _
        Open App.Path & "\15Minutes" For Append As #1: Print #1, TempVal: Close #1
        
    If Minutesi.ListCount - 1 = 60 Then
        TempVal = 0
        For i = 0 To Minutesi.ListCount - 1
            TempVal = TempVal + Val(Minutesi.List(i)) * 10
        Next
        TempVal = Int((TempVal / i)) / 10
        Houri.AddItem TempVal
        h.Text = Houri.ListCount
        Hri.Caption = TempVal & " %"
        Minutesi.Clear
    End If
    If Houri.ListCount - 1 = 24 Then
        TempVal = 0
        For i = 0 To Houri.ListCount - 1
            TempVal = TempVal + Val(Houri.List(i)) * 10
        Next
        TempVal = Int((TempVal / i)) / 10
        Dayi.AddItem TempVal
        d.Text = Dayi.ListCount
        Di.Caption = TempVal & " %"
        Houri.Clear
    End If
    If Dayi.ListCount - 1 = 7 Then
        TempVal = 0
        For i = 0 To Dayi.ListCount - 1
            TempVal = TempVal + Val(Dayi.List(i)) * 10
        Next
        TempVal = Int((TempVal / i)) / 10
        Weeki.AddItem TempVal
        w.Text = Weeki.ListCount
        Wki.Caption = TempVal & " %"
        Dayi.Clear
    End If
    If Weeki.ListCount - 1 = 4 Then
        TempVal = 0
        For i = 0 To Weeki.ListCount - 1
            TempVal = TempVal + Val(Weeki.List(i)) * 10
        Next
        TempVal = Int((TempVal / i)) / 10
        Monthi.AddItem TempVal
        mo.Text = Monthi.ListCount
        Moni.Caption = TempVal & " %"
        Weeki.Clear
    End If
    If Monthi.ListCount - 1 = 12 Then
        TempVal = 0
        For i = 0 To Monthi.ListCount - 1
            TempVal = TempVal + Val(Monthi.List(i)) * 10
        Next
        TempVal = Int((TempVal / i)) / 10
        Yri.Caption = TempVal
        Monthi.Clear
    End If
End If
Call DoStop
Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
Timer1.Enabled = False
StartButton_Click

End Sub
