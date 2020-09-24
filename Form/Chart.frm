VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form Chart 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Nocturnal Emissions"
   ClientHeight    =   6480
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7485
   DrawMode        =   10  'Mask Pen
   FillColor       =   &H00E0E0E0&
   FillStyle       =   6  'Cross
   Icon            =   "Chart.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6480
   ScaleWidth      =   7485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin RandomFlipper.AutoResize Resize 
      Left            =   6960
      Tag             =   "NO"
      Top             =   360
      _ExtentX        =   714
      _ExtentY        =   714
      AspectRatioValue=   0
   End
   Begin VB.ComboBox Combo2 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   315
      ItemData        =   "Chart.frx":08CA
      Left            =   9480
      List            =   "Chart.frx":08CC
      Sorted          =   -1  'True
      TabIndex        =   20
      Tag             =   "no"
      Text            =   "Actigraphic Data"
      Top             =   3480
      Width           =   2235
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      DrawMode        =   9  'Not Mask Pen
      FillColor       =   &H00FFFF80&
      FillStyle       =   3  'Vertical Line
      ForeColor       =   &H80000008&
      Height          =   1680
      Left            =   45
      ScaleHeight     =   1680
      ScaleWidth      =   7440
      TabIndex        =   12
      Top             =   4785
      Width           =   7440
      Begin VB.OptionButton graphType1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "2D Bar"
         Height          =   1245
         Index           =   1
         Left            =   270
         Picture         =   "Chart.frx":08CE
         Style           =   1  'Graphical
         TabIndex        =   19
         Tag             =   "grCol2d"
         Top             =   375
         Value           =   -1  'True
         Width           =   990
      End
      Begin VB.OptionButton graphType1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "2D Line"
         Height          =   1245
         Index           =   2
         Left            =   1245
         Picture         =   "Chart.frx":2F08
         Style           =   1  'Graphical
         TabIndex        =   18
         Tag             =   "grLine2D"
         Top             =   375
         Width           =   990
      End
      Begin VB.OptionButton graphType1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "2D Area"
         Height          =   1245
         Index           =   3
         Left            =   2250
         Picture         =   "Chart.frx":525A
         Style           =   1  'Graphical
         TabIndex        =   17
         Tag             =   "grArea2D"
         Top             =   375
         Width           =   990
      End
      Begin VB.OptionButton graphType1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "2D Stack"
         Height          =   1245
         Index           =   4
         Left            =   3270
         Picture         =   "Chart.frx":76E4
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   375
         Width           =   990
      End
      Begin VB.OptionButton graphType1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "3D Bar"
         Height          =   1245
         Index           =   5
         Left            =   4245
         Picture         =   "Chart.frx":9A16
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   375
         Width           =   990
      End
      Begin VB.OptionButton graphType1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "3D Area"
         Height          =   1245
         Index           =   7
         Left            =   6255
         Picture         =   "Chart.frx":BCB8
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   375
         Width           =   990
      End
      Begin VB.OptionButton graphType1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "3D Line"
         Height          =   1245
         Index           =   8
         Left            =   5265
         Picture         =   "Chart.frx":C269
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   375
         Width           =   990
      End
      Begin VB.Shape Shape3 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FFFF80&
         FillStyle       =   3  'Vertical Line
         Height          =   1695
         Left            =   -120
         Top             =   120
         Width           =   7815
      End
   End
   Begin VB.DirListBox Dir1 
      Height          =   315
      Left            =   8385
      TabIndex        =   11
      Top             =   4410
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.CommandButton cmdAddLegend 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Legend"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   6555
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4770
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.FileListBox File2 
      Height          =   285
      Left            =   8340
      Pattern         =   "*.avi"
      TabIndex        =   8
      Top             =   4035
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   8310
      Pattern         =   "*.act"
      TabIndex        =   7
      Top             =   3720
      Visible         =   0   'False
      Width           =   1035
   End
   Begin MSChart20Lib.MSChart MSChart1 
      Height          =   5655
      Left            =   -360
      OleObjectBlob   =   "Chart.frx":C716
      TabIndex        =   6
      Top             =   -240
      Width           =   7935
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   5640
      TabIndex        =   5
      Top             =   4920
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Data"
      Height          =   270
      Left            =   10560
      TabIndex        =   0
      Top             =   5760
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.ComboBox Combo4 
      Height          =   315
      Left            =   4560
      TabIndex        =   9
      Text            =   "Combo4"
      Top             =   4845
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   7320
      TabIndex        =   21
      Text            =   "Combo1"
      Top             =   4200
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.Shape Shape2 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFF80&
      FillStyle       =   3  'Vertical Line
      Height          =   1635
      Left            =   0
      Top             =   0
      Width           =   7710
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFF80&
      FillStyle       =   3  'Vertical Line
      Height          =   1635
      Left            =   -90
      Top             =   5190
      Width           =   7710
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "1998"
      Height          =   255
      Left            =   4680
      TabIndex        =   4
      Top             =   4920
      Width           =   855
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "1997"
      Height          =   255
      Left            =   4680
      TabIndex        =   3
      Top             =   4560
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "1996"
      Height          =   255
      Left            =   4680
      TabIndex        =   2
      Top             =   4200
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "1995"
      Height          =   255
      Left            =   4680
      TabIndex        =   1
      Top             =   3840
      Width           =   855
   End
End
Attribute VB_Name = "Chart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim Actig(40000) As String, ActigT(40000) As String, i As Long, Stuff As String, j As Long
Dim tglLegend As Boolean, tglTitle As Boolean
Dim ActigV(40000) As String, ActigA(40000) As String
Private Sub cmdAddLegend_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

  tglLegend = Not tglLegend
  
  'Call AddLegend(MSChart1, tglLegend)

End Sub

Private Sub Combo1_Click()
On Error Resume Next

    MSChart1.chartType = Combo1.ListIndex
    MSChart1.Refresh
End Sub

Private Sub Command2_Click()
On Error Resume Next
i = 1
'ReDim Actig(100000)
Open App.Path & "\Actigraphics\ThisOne" For Input As #10
Do While Not EOF(10)
    Line Input #10, Stuff
    Actig(i) = Trim(Left(Stuff, InStr(Stuff, " ")))
    ActigT(i) = Trim(Replace(Mid(Stuff, InStr(Stuff, " "), Len(Stuff)), "---", ""))
    i = i + 1
Loop
Close #10
'MSChart1.TitleText = "Nocturnal Emissions" 'The chart title
MSChart1.RowCount = i   'UBound(Actig)
MSChart1.RowLabelCount = Int(i / 10)
For j = 1 To i ' UBound(Actig)
    MSChart1.Row = j
    MSChart1.Data = Val(Actig(j))
    MSChart1.RowLabel = ActigT(j)   '"1995" 'The name of the first row
Next
Exit Sub
End Sub

Private Sub Form_Activate()
Dim j As Long, i As Long
i = 0
If Dir(App.Path & "\15Minutes") = "" Then
    MsgBox "15 minutes of data have not been sampled as of yet!"
    Unload Me
    Exit Sub
End If

Open App.Path & "\15Minutes" For Input As #1
    Do While Not EOF(1)
        Line Input #1, Actig(i)
        i = i + 1
    Loop
Close #1
MSChart1.RowCount = i   'UBound(Actig)
For j = 1 To i - 1 ' UBound(Actig)
    MSChart1.Column = 1
    MSChart1.Row = j
    MSChart1.Data = Val(Actig(j))
    'MSChart1.RowLabel = Left(Right(ActigT(j), 6), 2) & ":" _
                       & Mid((Right(ActigT(j), 6)), 3, 2) & ":" _
                       & Mid((Right(ActigT(j), 6)), 5, 2)
Next
MSChart1.Row = 1
'MSChart1.RowLabel = SStart
'MSChart1.RowLabelCount = 5

End Sub

Private Sub Form_Load()
Combo1.AddItem "3d Bar Graph"
Combo1.AddItem "2d Bar Graph"
Combo1.AddItem "3d Line Graph"
Combo1.AddItem "2d Line Graph"
Combo1.AddItem "3d Area Graph"
Combo1.AddItem "2d Step Graph"
Combo1.AddItem "3d Step Graph"
Dir1.Path = App.Path

End Sub

Private Sub Form_Unload(Cancel As Integer)

Set Chart = Nothing
End Sub

Private Sub graphType1_Click(Index As Integer)
Exit Sub
On Error Resume Next
  Dim graphInt As Integer, chartStr As String
  Me.Enabled = False
  With MSChart1
  
  chartStr = graphType1(Index).Caption
  
  Select Case chartStr
 
      Case Is = "2D Area"
      
         .chartType = VtChChartType2dArea
         '.Stacking = True
      Case Is = "3D Area"
      
         .chartType = 4
         '.Stacking = True
         
      Case Is = "2D Bar"
      
         .chartType = 1 'VtChChartType2dBar
         '.Stacking = False
  
     Case Is = "3D Bar"
  
       .chartType = 0   'VtChSeriesType3dBar
       '.Plot.Projection = VtProjectionTypeOblique
       '.Stacking = True
    
     Case Is = "2D Stack"
  
         .chartType = 10
         .Stacking = True
      
     Case Is = "2D Line"
     
       .chartType = 3       'VtChChartType2dLine
       '.Stacking = False
     Case Is = "3D Line"
     
       .chartType = 2       'VtChChartType2dLine
       '.Stacking = False
     
          
  End Select
  End With
    MSChart1.SetFocus
End Sub

Private Sub graphType1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
  Dim graphInt As Integer, chartStr As String
  Me.Enabled = False
  With MSChart1
  
  chartStr = graphType1(Index).Caption
  
  Select Case chartStr
 
      Case Is = "2D Area"
      
         .chartType = VtChChartType2dArea
         '.Stacking = True
      Case Is = "3D Area"
      
         .chartType = 4
         '.Stacking = True
         
      Case Is = "2D Bar"
      
         .chartType = 1 'VtChChartType2dBar
         '.Stacking = False
  
     Case Is = "3D Bar"
  
       .chartType = 0   'VtChSeriesType3dBar
       '.Plot.Projection = VtProjectionTypeOblique
       '.Stacking = True
    
     Case Is = "2D Stack"
  
         .chartType = 10
         .Stacking = True
      
     Case Is = "2D Line"
     
       .chartType = 3       'VtChChartType2dLine
       '.Stacking = False
     Case Is = "3D Line"
     
       .chartType = 2       'VtChChartType2dLine
       '.Stacking = False
     
          
  End Select
  End With
    MSChart1.SetFocus

End Sub

Private Sub graphType1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
graphType1(Index).Value = False
Me.Enabled = True
End Sub

Private Sub MSChart1_PointSelected(Series As Integer, DataPoint As Integer, MouseFlags As Integer, Cancel As Integer)
On Error Resume Next
Exit Sub
Dim msg, Style, Title, Help, Ctxt, Response, MyString, ii As Long
Help = "DEMO.HLP"   ' Define Help file.
Ctxt = 1000   ' Define topic
    MyString = Left(Right(ActigT(DataPoint), 6), 2) & ":" _
                       & Mid((Right(ActigT(DataPoint), 6)), 3, 2) & ":" _
                       & Mid((Right(ActigT(DataPoint), 6)), 5, 2)

Me.Caption = "Nocturnal Emissions: " & MyString
If DataPoint - 30 > 0 And DataPoint + 30 <= UBound(Actig) Then
For ii = DataPoint - 30 To DataPoint + 30
    If Dir(App.Path & "\Actigraphics\" & ActigT(DataPoint) & ".avi") <> "" Then
        msg = "Audio/Video Activity was detected within" _
        & vbCrLf & "60 seconds of this Data Point. " & vbCrLf _
        & "Do you want to review it?"   ' Define message.
        Style = vbYesNo + vbCritical + vbDefaultButton2   ' Define buttons.
        Title = "Motion Video Option"   ' Define title.
        Response = MsgBox(msg, Style, Title, Help, Ctxt)
        If Response = vbYes Then   ' User chose Yes.
           'MyVideoFile = App.Path & "\Actigraphics\" & ActigT(DataPoint) & ".avi"
        End If
    End If
Next
Else
    If Dir(App.Path & "\Actigraphics\" & ActigT(DataPoint) & ".avi") <> "" Then
        msg = "Audio/Video Activity was detected at this Data Point. " & vbCrLf _
        & "Do you want to review it?"   ' Define message.
        Style = vbYesNo + vbCritical + vbDefaultButton2   ' Define buttons.
        Title = "Motion Video Option"   ' Define title.
        Response = MsgBox(msg, Style, Title, Help, Ctxt)
        If Response = vbYes Then   ' User chose Yes.
           'MyVideoFile = App.Path & "\Actigraphics\" & ActigT(DataPoint) & ".avi"
           'Load VideoCx
        End If
    End If
End If
End Sub

Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

'If Button = vbLeftButton And ReleaseButtons = True Then
  'ReleaseCapture
  'SendMessage Picture2.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
'End If

End Sub

Private Sub Text6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

'If Button = vbLeftButton And ReleaseButtons = True Then
  'ReleaseCapture
  'SendMessage Text6.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
'End If

End Sub
