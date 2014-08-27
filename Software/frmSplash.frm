VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   5295
   ClientLeft      =   210
   ClientTop       =   1365
   ClientWidth     =   7455
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   7455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   30
      Left            =   5640
      Top             =   4080
   End
   Begin VB.ListBox lst 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   990
      ItemData        =   "frmSplash.frx":000C
      Left            =   720
      List            =   "frmSplash.frx":000E
      TabIndex        =   3
      Top             =   3720
      Width           =   3615
   End
   Begin VB.Line Line4 
      BorderWidth     =   5
      X1              =   7440
      X2              =   7440
      Y1              =   0
      Y2              =   5280
   End
   Begin VB.Line Line3 
      BorderWidth     =   5
      X1              =   0
      X2              =   7440
      Y1              =   5280
      Y2              =   5280
   End
   Begin VB.Line Line2 
      BorderWidth     =   5
      X1              =   7560
      X2              =   0
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line1 
      BorderWidth     =   5
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   5280
   End
   Begin VB.Label lblProductName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Project HID"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   32.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   360
      TabIndex        =   2
      Top             =   1800
      Width           =   3480
   End
   Begin VB.Label lblPlatform 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Platform : Win32"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1200
      TabIndex        =   1
      Top             =   2760
      Width           =   2505
   End
   Begin VB.Label lblVersion 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version : 1.0.0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2040
      TabIndex        =   0
      Top             =   3240
      Width           =   1620
   End
   Begin VB.Image Image1 
      Height          =   5295
      Left            =   240
      Picture         =   "frmSplash.frx":0010
      Top             =   0
      Width           =   7035
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub

Private Sub Frame1_Click()
    Unload Me
End Sub

Private Sub Timer1_Timer()
Static state As Integer
Select Case state
Case 0:
lst.AddItem "Loading Mainform"
Load MainForm
state = state + 1
lst.Selected(lst.ListCount - 1) = True

Case 1:
lst.AddItem "Connecting to HID"
state = state + 1
lst.Selected(lst.ListCount - 1) = True

Case 2:
lst.AddItem "Searching For Device"
state = state + 1
lst.Selected(lst.ListCount - 1) = True

Case 3:
lst.AddItem "VID:1240 PID :63"
state = state + 1
lst.Selected(lst.ListCount - 1) = True

Case 4:
If hidIsAvailable(VendorID, ProductID) Then
lst.AddItem "Device Connected"
state = state + 3
lst.Selected(lst.ListCount - 1) = True
Else
lst.AddItem "Device Unavailable"
state = state + 1
End If
lst.Selected(lst.ListCount - 1) = True
Case 5:

lst.AddItem "Waiting for Device ......"
state = state + 1
lst.Selected(lst.ListCount - 1) = True

Case 6:
If hidIsAvailable(VendorID, ProductID) Then
    lst.AddItem "Device Found"
    lst.Selected(lst.ListCount - 1) = True
    state = state + 1
Else
    Msg = MsgBox("Connect USB Device(VID :1240 PID :63) and press OK", vbOKCancel + vbInformation, "Device Request")
    If Msg = vbCancel Then
        Unload MainForm
        Unload Me
    End If
End If

Case 7:
lst.AddItem "Synchronizing"
state = state + 1
lst.Selected(lst.ListCount - 1) = True

Case 8:
lst.AddItem "Launching Main Application"
state = state + 1
lst.Selected(lst.ListCount - 1) = True
Case 9:
MainForm.Show
Unload Me

End Select
End Sub
