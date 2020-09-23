VERSION 5.00
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lunar Calculator 2.1"
   ClientHeight    =   6240
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4125
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MouseIcon       =   "frmMain.frx":030A
   ScaleHeight     =   6240
   ScaleWidth      =   4125
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Now"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1290
      TabIndex        =   5
      Top             =   5520
      Width           =   1500
   End
   Begin VB.CommandButton Command2 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3750
      TabIndex        =   4
      Top             =   5520
      Width           =   345
   End
   Begin VB.CommandButton Command1 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   15
      TabIndex        =   3
      Top             =   5520
      Width           =   315
   End
   Begin MSACAL.Calendar Calendar1 
      Height          =   2415
      Left            =   0
      TabIndex        =   2
      Top             =   3060
      Width           =   4125
      _Version        =   524288
      _ExtentX        =   7276
      _ExtentY        =   4260
      _StockProps     =   1
      BackColor       =   -2147483633
      Year            =   2006
      Month           =   11
      Day             =   1
      DayLength       =   1
      MonthLength     =   2
      DayFontColor    =   0
      FirstDay        =   1
      GridCellEffect  =   1
      GridFontColor   =   10485760
      GridLinesColor  =   -2147483632
      ShowDateSelectors=   -1  'True
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   -1  'True
      ShowTitle       =   -1  'True
      ShowVerticalGrid=   -1  'True
      TitleFontColor  =   10485760
      ValueIsNull     =   0   'False
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   3675
      Top             =   0
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   5865
      Width           =   4125
      _ExtentX        =   7276
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   7223
            MinWidth        =   5292
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
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
      Height          =   1500
      Left            =   1282
      Picture         =   "frmMain.frx":0614
      ScaleHeight     =   1500
      ScaleWidth      =   1500
      TabIndex        =   1
      Top             =   180
      Width           =   1500
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   1290
      Left            =   420
      MouseIcon       =   "frmMain.frx":13FF
      TabIndex        =   7
      ToolTipText     =   "Click to select between next and previous"
      Top             =   1710
      Width           =   3660
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   450
      TabIndex        =   6
      Top             =   -30
      Width           =   3225
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Rewritten and reformulated to use resource file by Thomas A. Swift Nov. 2 2006
'Special thanks to chosen_reject's original submission at
'http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=58642&lngWId=1
'Another special thanks goes to G.W. at
'http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=29000&lngWId=1
'for the MoonPhase.bas module
Dim DispDate As Date
Dim DispPrev As Boolean
Function FixPath(sPath As String) As String
    If Right(sPath, 1) = "\" Then
        FixPath = sPath
    Else
        FixPath = sPath & "\"
    End If
End Function
Private Sub Command1_Click()
    Picture1.SetFocus
    DispDate = DateAdd("m", -1, DispDate)
    Calendar1.Value = DispDate
    ShowMoon
End Sub
Private Sub Command2_Click()
    Picture1.SetFocus
    DispDate = DateAdd("m", 1, DispDate)
    Calendar1.Value = DispDate
    ShowMoon
End Sub
Private Sub Command3_Click()
    Picture1.SetFocus
    DispDate = Now
    Calendar1.Value = DispDate
    ShowMoon
End Sub
Private Sub Calendar1_Click()
    If DateDiff("y", Now, Calendar1.Value) <> 0 Or DateDiff("yyyy", Now, Calendar1.Value) <> 0 Then
    DispDate = Calendar1.Value
    DispDate = DateAdd("h", 24, DispDate)
    'DispDate = DateAdd("n", Minute(Now), DispDate)
    'DispDate = DateAdd("s", Second(Now), DispDate)
    Else
    DispDate = Now
    End If
    ShowMoon
End Sub
Private Sub Form_Load()
    DispDate = Now
    Calendar1.Value = DispDate
    ShowMoon
End Sub
Private Sub ShowMoon()
    Dim iResFile As Integer
    Dim iMoonFile As Integer
    StatusBar1.Panels(1).Text = "Night of " & WeekdayName(Weekday(DispDate), False) & " " & Format(DispDate, "mmmm d yyyy hh:mm")
    If DispPrev = False Then
        Label2.Caption = "Illumination: " & Illum(DispDate) & vbCrLf & "Angle: " & Angle(DispDate) & vbCrLf & "Next Full Moon: " & NextFullMoon(DispDate) & vbCrLf & "Next New Moon: " & NextNewMoon(DispDate) & vbCrLf & "Next First Quarter: " & NextFirstQuarter(DispDate) & vbCrLf & "Next Last Quarter: " & NextLastQuarter(DispDate)
    Else
        Label2.Caption = "Illumination: " & Illum(DispDate) & vbCrLf & "Angle: " & Angle(DispDate) & vbCrLf & "Prev Full Moon: " & PreviousFullMoon(DateAdd("d", -1, DispDate)) & vbCrLf & "Prev New Moon: " & PreviousNewMoon(DateAdd("d", -1, DispDate)) & vbCrLf & "Prev First Quarter: " & PreviousFirstQuarter(DateAdd("d", -1, DispDate)) & vbCrLf & "Prev Last Quarter: " & PreviousLastQuarter(DateAdd("d", -1, DispDate))
    End If
    iResFile = (Angle(DispDate) / 2) + 101
    If iResFile > 280 Then iResFile = 101
    Label1.Caption = "Moon " & GenerateCycleText(DispDate, Angle(DispDate) / 2) '& " Rise: " & JulianDaysToUT(MoonPhaseByLunation(Lunation(DispDate) + 1, 1))
    Picture1.Picture = LoadResPicture(iResFile, vbResBitmap)
End Sub
Public Sub LoadDataIntoFile(DataName As Integer, FileName As String)
    Dim myArray() As Byte
    Dim myFile As Long
    If Dir(FileName) = "" Then
        myArray = LoadResData(DataName, "CUSTOM")
        myFile = FreeFile
        Open FileName For Binary Access Write As #myFile
        Put #myFile, , myArray
        Close #myFile
    End If
End Sub
Private Sub Label2_Click()
    If DispPrev = True Then
        DispPrev = False
    Else
        DispPrev = True
    End If
    ShowMoon
End Sub
Private Sub Timer1_Timer()
    'Debug.Print DateDiff("y", Now, Calendar1.Value) & "-" & DateDiff("yyyy", Now, Calendar1.Value)
    If DateDiff("y", Now, Calendar1.Value) <> 0 Or DateDiff("yyyy", Now, Calendar1.Value) <> 0 Then
       DispDate = Calendar1.Value
       DispDate = DateAdd("h", 24, DispDate)
       'DispDate = DateAdd("n", Minute(Now), DispDate)
       'DispDate = DateAdd("s", Second(Now), DispDate)
    Else
    DispDate = Now
    End If
    If HOUR(Now) = 0 And Minute(Now) = 0 And Second(Now) = 0 Then Command3_Click
    ShowMoon
End Sub
Private Function GenerateCycleText(MyDate As Date, Phase As Integer) As String
    Dim MoonPrec As String
    If Phase < 90 Then
        If Format$(Illum(MyDate), "###") = "" Then
            MoonPrec = "100%"
            GenerateCycleText = MoonDescription(MyDate) & " " & MoonPrec & " of New"
        Else
            MoonPrec = Format$(Illum(MyDate), "###") + "%"
            GenerateCycleText = MoonDescription(MyDate) & " " & MoonPrec & " of Full"
        End If
        
    Else
        Phase = (100 - Illum(MyDate))
        If Format$(GetPercentage(Phase, 100), "###") = "" Then
            MoonPrec = "100%"
            GenerateCycleText = MoonDescription(MyDate) & " " & MoonPrec & " of Full"
        Else
            MoonPrec = Format$(GetPercentage(Phase, 100), "###") + "%"
            GenerateCycleText = MoonDescription(MyDate) & " " & MoonPrec & " of New"
        End If
        
    End If
End Function
Private Function GetPercentage(ByVal SoFar, ByVal Total) As Single 'Or Long whatever you want
    On Error GoTo UhOh 'If there's an error, goto the errorhandler
    'This is the simplest formula known to
    'mathamatics:
    If SoFar > Total Then GoTo UhOh 'If SoFar is higher than
    'the total then return -1 which means there was an error
    If SoFar = 0 Then
        GetPercentage = 0
        Exit Function
    End If
    SoFar = SoFar * 100 'Multiply by 100
    Total = SoFar / Total 'Then devide by the total to Get your percentage
    GetPercentage = Total 'Return the percentage
    Exit Function 'We don't want To activate the errorhandler and return 0 do we?
UhOh:
    GetPercentage = -1
End Function

