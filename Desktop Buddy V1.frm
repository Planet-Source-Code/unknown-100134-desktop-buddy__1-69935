VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{6FBA474E-43AC-11CE-9A0E-00AA0062BB4C}#1.0#0"; "SYSINFO.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Desktop Buddy"
   ClientHeight    =   5775
   ClientLeft      =   12075
   ClientTop       =   5760
   ClientWidth     =   7170
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   7170
   Begin VB.Frame Frame1 
      Caption         =   "WARNING! Computer actions"
      Height          =   1095
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   7215
      Begin VB.CommandButton Command4 
         Caption         =   "Restart"
         Height          =   615
         Left            =   4680
         TabIndex        =   14
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Shut Down"
         Height          =   615
         Left            =   2760
         TabIndex        =   13
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Log Off"
         Height          =   615
         Left            =   840
         TabIndex        =   12
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame fraBattery 
      Caption         =   "Battery Life"
      Height          =   1455
      Left            =   840
      TabIndex        =   0
      Top             =   3840
      Width           =   4215
      Begin MSComctlLib.ProgressBar prgBattery 
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.Label lblStatusString 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF00FF&
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   3975
      End
      Begin VB.Label lblStatus 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   3975
      End
      Begin VB.Label lbl50 
         Alignment       =   2  'Center
         Caption         =   "50 %"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   3975
      End
      Begin VB.Label lblFull 
         Alignment       =   1  'Right Justify
         Caption         =   "100 %"
         Height          =   255
         Left            =   3600
         TabIndex        =   3
         Top             =   480
         Width           =   495
      End
      Begin VB.Label lblEmpty 
         Caption         =   "Empty"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   495
      End
   End
   Begin SysInfoLib.SysInfo sysBattery 
      Left            =   2400
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Timer tmrBattery 
      Interval        =   5000
      Left            =   1920
      Top             =   0
   End
   Begin VB.CommandButton Command1 
      Caption         =   "About Software"
      Height          =   375
      Left            =   3360
      TabIndex        =   10
      Top             =   5400
      Width           =   1575
   End
   Begin VB.Timer Timer1 
      Left            =   720
      Top             =   2760
   End
   Begin VB.CommandButton cmdTimeLeft 
      Caption         =   "Time left on Battery"
      Height          =   375
      Left            =   360
      TabIndex        =   8
      Top             =   5400
      Width           =   1575
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Close Program"
      Height          =   375
      Left            =   5640
      TabIndex        =   6
      Top             =   5400
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Desktop Buddy"
      BeginProperty Font 
         Name            =   "GlooGun"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   735
      Left            =   3720
      TabIndex        =   15
      Top             =   1080
      Width           =   3495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1680
      TabIndex        =   9
      Top             =   2160
      Width           =   4455
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
MsgBox " Desktop Buddy 2008 V1     © Rhys Towey 2008"
End Sub

Private Sub Command2_Click()
'Log Off Windows XP
Shell "shutdown -l -f -t 0"

End Sub

Private Sub Command3_Click()
'Shutdown Windows XP
Shell "shutdown -s -f -t 0"
End Sub

Private Sub Command4_Click()
'Restart Windows XP
Shell "shutdown -r -f -t 0"
End Sub

Private Sub Form_Load()
MsgBox " Welcome to Desktop Buddy "
Timer1.Interval = 100
     Label1().Caption = ""
End Sub

Private Sub Option1_Click()
frmMain.Show
frmMain.BackColor = 12582912
End Sub

Private Sub Option2_Click()
frmMain.Show
frmMain.BackColor = &HFF&
End Sub

Private Sub Timer1_Timer()
Label1().Caption = Time
End Sub

Private Sub tmrBattery_Timer()
'   calls the functions to update the display.
    Progress
    PercentageLeft
    ACStatus
    
'   The following two lines change the window caption and the status
    frmMain.Caption = "Battery Remaining: " & prgBattery.Value & "%"
    lblStatus.Caption = prgBattery.Value & "% battery power left."
End Sub

Private Function Progress()
'   This function updates the progress bar.
'   Tests first for a value, then sets the progress bar value to
'   that value and enables the progress bar.  255 is status unknown
    If sysBattery.BatteryLifePercent <> 255 Then
        prgBattery.Value = sysBattery.BatteryLifePercent
        prgBattery.Enabled = True
    Else
        prgBattery.Value = 0
        prgBattery.Enabled = False
    End If
End Function
Private Function BatteryStatus()
'   Checks the current status for the battery.  this sets the caption
'   for the  lable and the color of the text based on battery state.
    Select Case sysBattery.BatteryStatus
        Case 1 ' battery is OK
            lblStatusString.Caption = "Battery OK"
            lblStatusString.ForeColor = &H8000&
        Case 2 ' Battery getting low
            lblStatusString.Caption = "Battery Low"
            lblStatusString.ForeColor = &HFF00FF
        Case 4 ' Um, you need to plug in
            lblStatusString.Caption = "Battery Critical"
            lblStatusString.ForeColor = &HFF&
        Case 8 ' Battery charging
            lblStatusString.Caption = "Battery Charging"
            lblStatusString.ForeColor = &HFF0000
        Case 128, 255 ' Cannot get status
            lblStatusString.Caption = "No Battery Status"
            lblStatusString.ForeColor = &H0&
    End Select
End Function

Private Function ACStatus()
'   This checks to see if the laptop is on AC or battery power and
'   update caption and text color as necessary.
    Select Case sysBattery.ACStatus
        Case 0 ' Signifies that unit is on battery power
            BatteryStatus
        Case 1 ' Signifies that unit is on AC power
            lblStatusString.Caption = "Unit is on AC Power"
            lblStatusString.ForeColor = &H8000&
        Case 255 ' Cannot get status.
            lblStatusString.Caption = "AC Status Unknown"
            lblStatusString.ForeColor = &HFF&
    End Select
End Function

Private Function PercentageLeft()
'   This sets the color of the caption text based on the amount
'   of charge left on the battery.
    If prgBattery.Value < 75 And prgBattery.Value > 50 Then
        lblStatus.ForeColor = &H8000&
    ElseIf prgBattery.Value < 50 And prgBattery.Value > 25 Then
        lblStatus.ForeColor = &HFF00FF
    ElseIf prgBattery.Value < 25 Then
        lblStatus.ForeColor = &HFF&
    Else
        lblStatus.ForeColor = &HFF0000
    End If
End Function

Private Sub cmdTimeLeft_Click()
'   this gets the time left (hrs/mins) on the battery.
'   I have not been able to see this work, I am taking
'   for granted that it does.
    If sysBattery.BatteryLifeTime <> &HFFFFFFFF Then
        Dim TimeLeft As String
        Dim TimeTotal As String
        Dim temp
        temp = TimeSerial(0, 0, sysBattery.BatteryLifeTime)
        TimeLeft = Format(temp, "h:mm")
        temp = TimeSerial(0, 0, sysBattery.BatteryFullTime)
        TimeTotal = Format(temp, "h:mm")
        MsgBox TimeLeft & " time left from " & TimeTotal & " total."
    Else
        MsgBox "Cannot determine the time remaining on the battery.", _
            vbInformation, "Time Left on Battery"
    End If
End Sub

Private Sub cmdExit_Click()
MsgBox " You are now leaving Desktop Buddy and thank you for using this program "
'   Ends program...
    End
End Sub

