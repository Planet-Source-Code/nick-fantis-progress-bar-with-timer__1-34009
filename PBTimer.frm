VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPBTimer 
   Caption         =   "Progress Bar Timer v. 1.00 by: Nick Fantis"
   ClientHeight    =   2100
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7005
   LinkTopic       =   "Form1"
   ScaleHeight     =   2100
   ScaleWidth      =   7005
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAbout 
      Caption         =   "&About"
      Height          =   495
      Left            =   3000
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   5280
      TabIndex        =   3
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "&Stop"
      Height          =   495
      Left            =   3000
      TabIndex        =   2
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   120
      Top             =   120
   End
   Begin VB.CommandButton cmdRun 
      Caption         =   "&Run"
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   1440
      Width           =   1215
   End
   Begin MSComctlLib.ProgressBar pb1 
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   840
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   661
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   1
      Scrolling       =   1
   End
End
Attribute VB_Name = "frmPBTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'|--------------------------------------|
'|  Progress Bar Timer v. 1.00          |
'|  by: Nick Fantis                     |
'|  phantis86@hotmail.com               |
'|                                      |
'|  This code demonstrates how to hook  |
'|  a timer up to a progress bar        |
'|  accurately.  It allows you to stop  |
'|  and resume from where you left off. |
'|  It also shows the percent complete. |
'|  I haven't worked with progress bars |
'|  that much so there may be errors.   |
'|  If you find any please let me know. |
'|                                      |
'|  Feel free to use any of this code.  |
'|              Enjoy! :)               |
'|--------------------------------------|

Option Explicit

'Timer Interval Variable
Private TmrInterval As Integer

'Progress bar Max Value Variable
Private PBMaxVal As Integer

'Progress bar Increment Variable
Private PBIncrement As Integer

'See timer1 remarks for info on this variable
Private PBDifference As Integer

'Used to stop and resume progress bar
Private Running As Boolean


Private Sub cmdAbout_Click()

    frmAbout.Show
    
End Sub

Private Sub cmdExit_Click()

    Unload Me
    End
    
End Sub

Private Sub cmdRun_Click()

    On Error GoTo ErrorHandle
    
    pb1.Value = 0
    
    PBMaxVal = InputBox("Type in the maximum value of the progess bar.")
    pb1.Max = PBMaxVal
    
    PBIncrement = InputBox("Type in the increment of the progress bar.")
    
    TmrInterval = InputBox("Type in the timer interval.")
    Timer1.Interval = TmrInterval
    Running = True
    Timer1.Enabled = True

Exit Sub

ErrorHandle:
    If Err.Number = 13 Then
        MsgBox ("The input box can only contain a number")
        Exit Sub
    Else
        MsgBox ("An unexpected error has occured.")
        Exit Sub
    End If
    
End Sub

Private Sub cmdStop_Click()
    
    'This code allows you to stop and continue the progressbar.
    
    If pb1.Value = 0 Then Exit Sub
    
    Select Case Running
        Case Is = True
            Timer1.Enabled = False
            Running = False
            cmdStop.Caption = "&Continue"
        Case Is = False
            Timer1.Enabled = True
            Running = True
            cmdStop.Caption = "&Stop"
    End Select
    
End Sub

Private Sub Form_Load()

    Running = False
    
End Sub

Private Sub Timer1_Timer()
    
    If pb1.Value < pb1.Max Then
        PBDifference = pb1.Max - pb1.Value
        
        'The following If statement checks to see if the maximum value minus
        '   value and current value is less than the increment.
        'If the difference is greater it moves to Else.
        'If it is not then this line of code is acitvated.
        'It makes sure we do not recieve a run-time error and makes sure
        '   the progress bar fills completely.
        If PBDifference <= PBIncrement Then
            pb1.Value = pb1.Value + PBDifference
            frmPBTimer.Caption = "Progress: " & Format(pb1.Value / PBMaxVal, "##.00%")
        Else
            pb1.Value = pb1.Value + PBIncrement
            frmPBTimer.Caption = "Progress: " & Format(pb1.Value / PBMaxVal, "##.00%")
        End If
    End If
    
End Sub
