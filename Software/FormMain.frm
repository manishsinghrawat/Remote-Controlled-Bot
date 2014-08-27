VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form MainForm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PIC Control Panel"
   ClientHeight    =   6240
   ClientLeft      =   5490
   ClientTop       =   4080
   ClientWidth     =   8415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   8415
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "GO"
      Height          =   615
      Left            =   2760
      TabIndex        =   5
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Timer Timer2 
      Interval        =   50
      Left            =   1560
      Top             =   3960
   End
   Begin VB.Timer readbutton 
      Interval        =   50
      Left            =   1320
      Top             =   5400
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   2295
      Index           =   0
      Left            =   600
      TabIndex        =   1
      Top             =   360
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   4048
      _Version        =   393216
      Appearance      =   0
      Orientation     =   1
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   2295
      Index           =   1
      Left            =   1200
      TabIndex        =   4
      Top             =   360
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   4048
      _Version        =   393216
      Appearance      =   0
      Min             =   90
      Max             =   145
      Orientation     =   1
      Scrolling       =   1
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Gripper Signal"
      Height          =   375
      Left            =   1080
      TabIndex        =   17
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Gripper Feedback"
      Height          =   375
      Left            =   240
      TabIndex        =   16
      Top             =   3240
      Width           =   735
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Elevation"
      Height          =   255
      Left            =   3000
      TabIndex        =   15
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Motor Speed"
      Height          =   255
      Left            =   3000
      TabIndex        =   14
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Gripper Enabled"
      Height          =   255
      Left            =   1680
      TabIndex        =   13
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Collision"
      Height          =   255
      Left            =   1680
      TabIndex        =   12
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Object Status"
      Height          =   255
      Left            =   1680
      TabIndex        =   11
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   375
      Left            =   3120
      TabIndex        =   10
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   255
      Left            =   3120
      TabIndex        =   9
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   375
      Left            =   3120
      TabIndex        =   8
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Label5"
      Height          =   375
      Left            =   4560
      TabIndex        =   7
      Top             =   960
      Width           =   615
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Label4"
      Height          =   375
      Left            =   2160
      TabIndex        =   6
      Top             =   960
      Width           =   615
   End
   Begin VB.Image Image2 
      Height          =   750
      Left            =   6240
      Picture         =   "FormMain.frx":0000
      Top             =   3360
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image Image1 
      Height          =   1770
      Left            =   3600
      Picture         =   "FormMain.frx":08E6
      Top             =   3360
      Visible         =   0   'False
      Width           =   1770
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      Height          =   375
      Left            =   4560
      TabIndex        =   3
      Top             =   480
      Width           =   3495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label lbls 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   2760
      Width           =   615
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Const BufferInSize = 64
Private Const BufferOutSize = 64
Dim BufferIn(0 To BufferInSize) As Byte
Dim BufferOut(0 To BufferOutSize) As Byte

Dim bsy As Boolean
Dim bs1 As Boolean
Dim bs2 As Boolean
Dim bs3 As Boolean

Dim lft As Integer
Dim rgt As Integer

Dim grp As Integer
Dim up As Boolean
Dim dwn As Boolean
Dim opn As Boolean
Dim kls As Boolean

Dim kk As Boolean
Dim ii As Boolean
Dim jj As Boolean
Dim ll As Boolean

Dim acc As Integer
Dim accl As Double
Dim vel As Integer
Dim ang As Integer
Dim enb As Boolean

Dim command As Byte

Private Sub Command2_Click()
readbutton.Enabled = Not readbutton.Enabled
End Sub



Private Sub Command5_Click()
BufferOut(0) = 0
BufferOut(1) = &H84
     hidWriteEx VendorID, ProductID, BufferOut(0)
End Sub

Private Sub Command1_KeyDown(KeyCode As Integer, Shift As Integer)
If bs3 = False Then
If KeyCode = 71 Then
bs3 = True
enb = Not enb
End If
End If




If bs2 = False Then
If KeyCode = 85 Then
bs2 = True
opn = True
ElseIf KeyCode = 79 Then
bs2 = True
kls = True
End If
End If


If bs1 = False Then
If KeyCode = 89 Then
bs1 = True
up = True
ElseIf KeyCode = 72 Then
bs1 = True
dwn = True
End If
End If

If bsy = False Then
If KeyCode = 75 Then
kk = True
bsy = True
lft = -12
rgt = -12
ElseIf KeyCode = 73 Then
ii = True
bsy = True
lft = 12
rgt = 12
ElseIf KeyCode = 74 Then
jj = True
bsy = True
lft = -12
rgt = 12
ElseIf KeyCode = 76 Then
ll = True
bsy = True
lft = 12
rgt = -12
End If
End If

End Sub

Private Sub Command1_KeyUp(KeyCode As Integer, Shift As Integer)
If bs3 = True Then
If KeyCode = 71 Then
bs3 = False
End If
End If


If bs2 = True Then
If KeyCode = 85 And opn Then
bs2 = False
opn = False
ElseIf KeyCode = 79 And kls Then
bs2 = False
kls = False
End If
End If

If bs1 = True Then
If KeyCode = 89 And up Then
bs1 = False
up = False
ElseIf KeyCode = 72 And dwn Then
bs1 = False
dwn = False
End If
End If

If bsy = True Then
If KeyCode = 75 And kk Then
bsy = False
lft = 0
rgt = 0
kk = False
ElseIf KeyCode = 73 And ii Then
bsy = False
lft = 0
rgt = 0
ii = False
ElseIf KeyCode = 74 And jj Then
bsy = False
lft = 0
rgt = 0
jj = False
ElseIf KeyCode = 76 And ll Then
bsy = False
lft = 0
rgt = 0
ll = False
End If
End If

End Sub

Private Sub Form_Load()
   ConnectToHID (Me.hwnd)
bsy = False
acc = 3
bs1 = False
bs2 = False
grp = 145
accl = 3
End Sub

Private Sub Form_Unload(Cancel As Integer)
   DisconnectFromHID
End Sub

Public Sub OnPlugged(ByVal pHandle As Long)
   If hidGetVendorID(pHandle) = VendorID And hidGetProductID(pHandle) = ProductID Then
      ' ** YOUR CODE HERE **
   End If
End Sub

Public Sub OnUnplugged(ByVal pHandle As Long)
   If hidGetVendorID(pHandle) = VendorID And hidGetProductID(pHandle) = ProductID Then

   End If
End Sub

Public Sub OnChanged()
   Dim DeviceHandle As Long
   
   ' get the handle of the device we are interested in, then set
   ' its read notify flag to true - this ensures you get a read
   ' notification message when there is some data to read...
   DeviceHandle = hidGetHandle(VendorID, ProductID)
   hidSetReadNotify DeviceHandle, True
End Sub

'*****************************************************************
' on read event...
'*****************************************************************
Public Sub OnRead(ByVal pHandle As Long)
On Error Resume Next
   If hidRead(pHandle, BufferIn(0)) Then
        i = 1
       If BufferIn(1) = &H88 Then
      Label1.Caption = CInt(BufferIn(2))
      End If
        If BufferIn(1) = &H90 Then
      Label6.Caption = CInt(BufferIn(2))
      End If
      If BufferIn(1) = &H91 Then
      Label7.Caption = CInt(BufferIn(2))
      End If
      If BufferIn(1) = &H82 Then

      ProgressBar1(0).Value = ((CInt(BufferIn(3)) * 256 + CInt(BufferIn(2))) / 1024) * 95
      lbls.Caption = (CInt(BufferIn(3)) * 256 + CInt(BufferIn(2))) / 5

      End If
   End If
End Sub

Public Sub WriteData()
   hidWriteEx VendorID, ProductID, BufferOut(0)
End Sub
 
 
Private Sub in_Click(Index As Integer)

End Sub

Private Sub ReadButton_Timer()

   BufferOut(0) = 0
  BufferOut(1) = &H82
  BufferOut(2) = i
   hidWriteEx VendorID, ProductID, BufferOut(0)
  
    BufferOut(1) = &H84

    If dwn Then
        BufferOut(2) = 1
    Else
        BufferOut(2) = 0
    End If
    
    hidWriteEx VendorID, ProductID, BufferOut(0)
  
  BufferOut(1) = &H87

    If up Then
        BufferOut(2) = 1
    Else
        BufferOut(2) = 0
    End If
    
    hidWriteEx VendorID, ProductID, BufferOut(0)
  
  
      BufferOut(1) = &H83

   BufferOut(2) = grp
   hidWriteEx VendorID, ProductID, BufferOut(0)
  
  
      BufferOut(1) = &H88

   
   hidWriteEx VendorID, ProductID, BufferOut(0)
  
      BufferOut(1) = &H90

   
   hidWriteEx VendorID, ProductID, BufferOut(0)
    BufferOut(1) = &H91

   If enb Then
   BufferOut(2) = 1
   Else
   BufferOut(2) = 0
   End If
   
   hidWriteEx VendorID, ProductID, BufferOut(0)
  
  
End Sub



Private Sub Timer2_Timer()
If bs2 = True Then
If opn = True Then

If grp > 93 Then
grp = grp - accl
End If

ElseIf kls = True Then

If grp < 142 Then
grp = grp + accl
End If
End If

End If

If bsy = True Then
If kk = True Then
If lft > -95 Then
lft = lft - acc
End If
If rgt > -95 Then
rgt = rgt - acc
End If
ElseIf ii = True Then
If lft < 95 Then
lft = lft + acc
End If
If rgt < 95 Then
rgt = rgt + acc
End If
ElseIf jj = True Then
If lft > -95 Then
lft = lft - acc
End If
If rgt < 95 Then
rgt = rgt + acc
End If
ElseIf ll = True Then
If lft < 95 Then
lft = lft + acc
End If
If rgt > -95 Then
rgt = rgt - acc
End If
End If
End If

Label2.Caption = lft
Label3.Caption = rgt
Label4.Caption = up
Label5.Caption = dwn
ProgressBar1(1).Value = grp
   
   BufferOut(0) = 0
   BufferOut(1) = &H85
   BufferOut(2) = lft + 100
   hidWriteEx VendorID, ProductID, BufferOut(0)
   BufferOut(1) = &H86
   BufferOut(2) = rgt + 100
   hidWriteEx VendorID, ProductID, BufferOut(0)
End Sub


Function floor(num As Double) As Byte
If CInt(num) <= num Then
    floor = CInt(num)
Else
    floor = CInt(num) - 1
End If
End Function
