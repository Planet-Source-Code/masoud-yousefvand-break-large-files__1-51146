VERSION 5.00
Begin VB.Form FrmWizard1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " File splitter wizard"
   ClientHeight    =   4644
   ClientLeft      =   36
   ClientTop       =   324
   ClientWidth     =   6924
   Icon            =   "FrmWizard1.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4644
   ScaleWidth      =   6924
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit Wizard"
      Height          =   372
      Left            =   5172
      TabIndex        =   1
      Top             =   4080
      Width           =   1572
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "&Next     >>"
      Height          =   372
      Left            =   2880
      TabIndex        =   0
      Top             =   4080
      Width           =   1572
   End
   Begin VB.Label lblGreeting 
      Caption         =   "Not Registered."
      Height          =   252
      Left            =   830
      TabIndex        =   7
      Top             =   3468
      Width           =   5172
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Welcome to 'File Splitter' demo version"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   192
      Left            =   1752
      TabIndex        =   6
      Top             =   192
      Width           =   3264
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "This wizard helps you to split a file into desired pieces with custom sizes."
      Height          =   192
      Left            =   830
      TabIndex        =   5
      Top             =   816
      Width           =   5100
   End
   Begin VB.Label Label3 
      Caption         =   $"FrmWizard1.frx":030A
      Height          =   612
      Left            =   830
      TabIndex        =   4
      Top             =   1296
      Width           =   5052
   End
   Begin VB.Label Label4 
      Caption         =   "This application will also assemble the pieces that it have made and make your original file for you."
      Height          =   492
      Left            =   830
      TabIndex        =   3
      Top             =   2160
      Width           =   5172
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "If you encountered any problem please contact : yousefvand@yahoo.com."
      Height          =   192
      Left            =   830
      TabIndex        =   2
      Top             =   2988
      Width           =   5292
   End
End
Attribute VB_Name = "FrmWizard1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()

    ExitWizard

End Sub

Private Sub cmdNext_Click()
    
    FrmWizard2.Show
    FrmWizard2.cmdNext.SetFocus
    Unload Me
    
End Sub

Private Sub Command2_Click()

    ExitWizard

End Sub

Private Sub Form_Load()

    lblGreeting.Caption = GreetingMessage

End Sub
