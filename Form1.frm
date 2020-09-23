VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "FLASH.OCX"
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "DON'T press the red button"
   ClientHeight    =   6345
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   10560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "I give up!"
      Height          =   375
      Left            =   8640
      TabIndex        =   2
      Top             =   5160
      Width           =   1335
   End
   Begin VB.CommandButton cmdCredits 
      Caption         =   "Credits"
      Height          =   375
      Left            =   8640
      TabIndex        =   1
      Top             =   4680
      Width           =   1335
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash ShockwaveFlash1 
      Height          =   3135
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   3855
      _cx             =   6800
      _cy             =   5530
      FlashVars       =   ""
      Movie           =   ""
      Src             =   ""
      WMode           =   "Window"
      Play            =   0   'False
      Loop            =   -1  'True
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   0   'False
      Base            =   ""
      AllowScriptAccess=   "always"
      Scale           =   "ShowAll"
      DeviceFont      =   0   'False
      EmbedMovie      =   0   'False
      BGColor         =   ""
      SWRemote        =   ""
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCredits_Click()
   MsgBox "I found this Shockwave movie at http://www.zipcity.nl/" & vbCrLf & _
          "Full credit goes to the maker(s) of this pearl." & vbCrLf & _
          "My only contribution was to embedd it in this program." & vbCrLf & _
          " ... Paul Turcksin", vbOKOnly, "Big red button"
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub Form_Load()
' full screen
   Me.Move 0, 0, Screen.Width, Screen.Height
' position command buttons
   cmdExit.Move Me.ScaleWidth - 1500, Me.ScaleHeight - 500
   cmdCredits.Move Me.ScaleWidth - 1500, Me.ScaleHeight - 1000
' Shockwave movie
   ShockwaveFlash1.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
   ShockwaveFlash1.Movie = App.Path & "\bigredbutton.swf"
   ShockwaveFlash1.Playing = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set Form1 = Nothing
End Sub
