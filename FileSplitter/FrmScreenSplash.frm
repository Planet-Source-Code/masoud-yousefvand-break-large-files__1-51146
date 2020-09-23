VERSION 5.00
Begin VB.Form FrmScreenSplash 
   BorderStyle     =   0  'None
   Caption         =   " File splitter wizard"
   ClientHeight    =   3570
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5025
   Icon            =   "FrmScreenSplash.frx":0000
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   Picture         =   "FrmScreenSplash.frx":030A
   ScaleHeight     =   238
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrSpalsh 
      Interval        =   5000
      Left            =   3120
      Top             =   120
   End
End
Attribute VB_Name = "FrmScreenSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

    Me.Width = 300 * Screen.TwipsPerPixelX
    Me.Height = 200 * Screen.TwipsPerPixelY

End Sub

Private Sub tmrSpalsh_Timer()

    Unload Me
    FrmWizard1.Show

End Sub
