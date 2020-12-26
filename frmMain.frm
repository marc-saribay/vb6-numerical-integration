VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Numerical Integration Calculator"
   ClientHeight    =   3645
   ClientLeft      =   2805
   ClientTop       =   2610
   ClientWidth     =   7455
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   7455
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdMethod 
      Caption         =   "&Simpson's Rule"
      Height          =   375
      Index           =   4
      Left            =   5280
      TabIndex        =   11
      Top             =   1845
      Width           =   1935
   End
   Begin VB.CommandButton cmdMethod 
      Caption         =   "&Trapezoidal Rule"
      Height          =   375
      Index           =   3
      Left            =   5280
      TabIndex        =   10
      Top             =   1365
      Width           =   1935
   End
   Begin VB.CommandButton cmdMethod 
      Caption         =   "&Midpoint Method"
      Height          =   375
      Index           =   2
      Left            =   5280
      TabIndex        =   9
      Top             =   885
      Width           =   1935
   End
   Begin VB.CommandButton cmdMethod 
      Caption         =   "&Rectangular Method"
      Height          =   375
      Index           =   1
      Left            =   5280
      TabIndex        =   8
      Top             =   405
      Width           =   1935
   End
   Begin VB.Frame fraOutput 
      Height          =   855
      Left            =   120
      TabIndex        =   12
      Top             =   2400
      Width           =   7215
      Begin VB.Label lblArea 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1740
         TabIndex        =   14
         Top             =   300
         Width           =   5175
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Estimated area = "
         Height          =   195
         Left            =   270
         TabIndex        =   13
         Top             =   345
         Width           =   1290
      End
   End
   Begin VB.TextBox txtValue 
      Height          =   300
      Index           =   3
      Left            =   1680
      TabIndex        =   7
      ToolTipText     =   "Determines the number of sub-intervals"
      Top             =   1920
      Width           =   3375
   End
   Begin VB.TextBox txtValue 
      Height          =   300
      Index           =   2
      Left            =   1680
      TabIndex        =   5
      ToolTipText     =   "Determines the upper bound"
      Top             =   1440
      Width           =   3375
   End
   Begin VB.TextBox txtValue 
      Height          =   300
      Index           =   1
      Left            =   1680
      TabIndex        =   3
      ToolTipText     =   "Determines the lower bound"
      Top             =   960
      Width           =   3375
   End
   Begin VB.TextBox txtFunction 
      Height          =   300
      Left            =   240
      TabIndex        =   1
      ToolTipText     =   "Use programming concept in writing equations for f(x)"
      Top             =   480
      Width           =   4815
   End
   Begin MSScriptControlCtl.ScriptControl ScriptControl1 
      Left            =   6705
      Top             =   2415
      _ExtentX        =   1005
      _ExtentY        =   1005
      AllowUI         =   -1  'True
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "©2001 by Marc Christian Saribay"
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   4920
      TabIndex        =   16
      ToolTipText     =   "Numerical Integration™ (September 30, 2001)"
      Top             =   3360
      Width           =   2385
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Numerical Integration Calculator™"
      Height          =   195
      Left            =   4875
      TabIndex        =   15
      Top             =   75
      Width           =   2460
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Enter the &f(x)"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Enter value for &n"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Enter value for &b"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Enter value for &a"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   1215
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdMethod_Click(Index As Integer)
  IntegrationMethod (Index)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyEscape Then
    Unload Me
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Unload Me
End Sub

Private Sub txtValue_GotFocus(Index As Integer)
  txtValue(Index).SelStart = 0
  txtValue(Index).SelLength = Len(txtValue(Index).Text)
End Sub
