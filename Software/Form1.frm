VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PIC USB CONSOLE"
   ClientHeight    =   4410
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7425
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   7425
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List 
      Height          =   3570
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   7215
   End
   Begin VB.TextBox data 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5775
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "Send"
      Height          =   375
      Left            =   6000
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Const BufferInSize = 64
Private Const BufferOutSize = 64
Dim BufferIn(0 To BufferInSize) As Byte
Dim BufferOut(0 To BufferOutSize) As Byte


Private Sub Command1_Click()
BufferOut(1) = &H90
For i = 1 To Len(data.Text)
    BufferOut(i + 1) = CByte(Asc(Mid(data.Text, i, 1)))
Next

hidWriteEx VendorID, ProductID, BufferOut(0)
data.Text = vbNullString
End Sub

Private Sub data_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
Command1_Click

End If

End Sub

Private Sub Form_Load()
 ConnectToHID (Me.hwnd)
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
i = 1
   If hidRead(pHandle, BufferIn(0)) Then

     
        
            While ((Not BufferIn(i) = 0) And i < 64)
                txt = txt + Chr(BufferIn(i))
                i = i + 1
            Wend
        List.AddItem (txt)
        If List.ListCount > 20 Then List.RemoveItem (0)
        List.Selected(List.ListCount - 1) = True
    
    End If
End Sub

Public Sub WriteData()
   hidWriteEx VendorID, ProductID, BufferOut(0)
End Sub
 


Function floor(num As Double) As Byte
If CInt(num) <= num Then
    floor = CInt(num)
Else
    floor = CInt(num) - 1
End If
End Function

Private Sub txt_Change()
txt.SelStart = Len(txt.Text)
End Sub

