VERSION 5.00
Begin VB.Form frmshutdown 
   BackColor       =   &H00000000&
   Caption         =   "Timed ShutDown..."
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   Icon            =   "shutdown.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtsec 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   2520
      TabIndex        =   9
      Top             =   840
      Width           =   375
   End
   Begin VB.TextBox txtmin 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   1920
      TabIndex        =   8
      Top             =   840
      Width           =   375
   End
   Begin VB.CommandButton cmdreset 
      Caption         =   "RESET!!"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   5
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Timer tmrdone 
      Interval        =   1000
      Left            =   4080
      Top             =   840
   End
   Begin VB.CommandButton cmdstart 
      Caption         =   "START!!"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   4
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox txthr 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   1320
      TabIndex        =   3
      Top             =   840
      Width           =   375
   End
   Begin VB.TextBox txtnow 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   1320
      TabIndex        =   1
      Top             =   240
      Width           =   1575
   End
   Begin VB.Timer tmrshut 
      Interval        =   1000
      Left            =   4080
      Top             =   1440
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2400
      TabIndex        =   10
      Top             =   840
      Width           =   135
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1800
      TabIndex        =   7
      Top             =   840
      Width           =   135
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Copyright of Raynard. All rights reserved ®"
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   840
      TabIndex        =   6
      Top             =   2640
      Width           =   3135
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Target Time: "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Time Now  : "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "frmshutdown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim now As String
Dim target As String

Private Sub cmdreset_Click()
tmrdone.Enabled = False
txthr.Text = " "
txtmin.Text = " "
txtsec.Text = " "
End Sub

Private Sub cmdstart_Click()
If txtmin.Text >= 60 Or txthr.Text >= 24 Or txtsec.Text >= 60 Then
    MsgBox " Please check your Time to be filled in again.."
    txthr.Text = " "
    txtmin.Text = " "
    txtsec.Text = " "
    cmdstart.Enabled = True
    tmrdone.Enabled = False
Else
    tmrdone.Enabled = True
End If
End Sub

Private Sub Form_Load()
txtnow.Text = Time$
tmrdone.Enabled = False
tmrshut.Enabled = True
End Sub

Private Sub tmrdone_Timer()
cmdstart.Enabled = False
cmdreset.Enabled = False
Dim after As String
Dim e As ShutError

after = txthr.Text & ":" & txtmin.Text & ":" & txtsec.Text
target = Format(after, "hh:mm:ss")
If now = target Then
    e = ShutDown()
    If e = sePrivileges Then
        MsgBox "Insufficient rights to shut down system"
    ElseIf e = seShutdown Then
        MsgBox "Error in ExitWindowsEx"
    End If

    Unload Me
End If
End Sub

Private Sub tmrshut_Timer()
txtnow.Text = " "
now = Format(Time$, "hh:mm:ss")
txtnow.Text = now
End Sub
