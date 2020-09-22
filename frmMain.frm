VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Get External IP"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7230
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000F&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   7230
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtLocalHost 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   3240
      Width           =   3015
   End
   Begin VB.TextBox txtDomain 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   2640
      Width           =   3015
   End
   Begin VB.TextBox txtExtIP 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   2040
      Width           =   3015
   End
   Begin VB.TextBox txtLocalIP 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   1440
      Width           =   3015
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   4335
      Width           =   7230
      _ExtentX        =   12753
      _ExtentY        =   450
      Style           =   1
      SimpleText      =   "Status : Idle."
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   8040
      Top             =   3600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Socket 
      Left            =   7920
      Top             =   3000
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet Inet 
      Left            =   7800
      Top             =   1920
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Label lblCopy 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Copy"
      Height          =   375
      Left            =   2640
      MouseIcon       =   "frmMain.frx":0ECA
      MousePointer    =   99  'Custom
      TabIndex        =   13
      Top             =   3825
      Width           =   1935
   End
   Begin VB.Image CopyUp 
      Height          =   480
      Left            =   2640
      Picture         =   "frmMain.frx":11D4
      Top             =   3720
      Width           =   1950
   End
   Begin VB.Image CopyDown 
      Height          =   480
      Left            =   2640
      Picture         =   "frmMain.frx":36A2
      Top             =   3720
      Width           =   1950
   End
   Begin VB.Label lblSave 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Save Info"
      Height          =   375
      Left            =   240
      MouseIcon       =   "frmMain.frx":5B1E
      MousePointer    =   99  'Custom
      TabIndex        =   12
      Top             =   3825
      Width           =   1935
   End
   Begin VB.Image SaveUp 
      Height          =   480
      Left            =   240
      Picture         =   "frmMain.frx":5E28
      Top             =   3720
      Width           =   1950
   End
   Begin VB.Image SaveDown 
      Height          =   480
      Left            =   240
      Picture         =   "frmMain.frx":82F6
      Top             =   3720
      Width           =   1950
   End
   Begin VB.Label lblGetRefresh 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Get / Refresh"
      Height          =   375
      Left            =   5040
      MouseIcon       =   "frmMain.frx":A772
      MousePointer    =   99  'Custom
      TabIndex        =   11
      Top             =   3825
      Width           =   1935
   End
   Begin VB.Image GetUp 
      Height          =   480
      Left            =   5040
      Picture         =   "frmMain.frx":AA7C
      Top             =   3720
      Width           =   1950
   End
   Begin VB.Image GetDown 
      Height          =   480
      Left            =   5040
      Picture         =   "frmMain.frx":CF4A
      Top             =   3720
      Width           =   1950
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Local Hostname / Computer Name :"
      Height          =   210
      Left            =   360
      TabIndex        =   6
      Top             =   3240
      Width           =   3360
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Host / Domain Name :"
      Height          =   210
      Left            =   360
      TabIndex        =   5
      Top             =   2640
      Width           =   2070
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "External IP Address :"
      Height          =   210
      Left            =   360
      TabIndex        =   4
      Top             =   2040
      Width           =   1995
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Local IP Address :"
      Height          =   210
      Left            =   360
      TabIndex        =   3
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Retrieved Information"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   2250
      TabIndex        =   2
      Top             =   1005
      Width           =   2715
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   6960
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Shape Shape1 
      Height          =   2655
      Left            =   240
      Top             =   960
      Width           =   6735
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   240
      Picture         =   "frmMain.frx":F3C6
      Stretch         =   -1  'True
      Top             =   960
      Width           =   6735
   End
   Begin VB.Image imgICO 
      Height          =   480
      Left            =   120
      Picture         =   "frmMain.frx":11BB3
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmMain.frx":11FF5
      Height          =   735
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   6255
   End
   Begin VB.Image imgBG 
      Height          =   4815
      Left            =   0
      Picture         =   "frmMain.frx":1209C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7695
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub lblCopy_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call DoButton(Button, CopyUp, CopyDown) 'use the handy DoButton() routine to simulate the press of a button
End Sub

Private Sub lblCopy_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call DoButton(Button, CopyDown, CopyUp)
Dim sDat As String: sDat = Empty 'we're gonna store all the retrieved info into a variable, and then copy it to the clipboard
'gather all of the information from the text boxes
sDat = "Local IP Address : " & txtLocalIP.Text & vbNewLine
sDat = sDat & "External IP Address : " & txtExtIP.Text & vbNewLine
sDat = sDat & "Host / Domain Name : " & txtDomain.Text & vbNewLine
sDat = sDat & "Local Hostname / Computer Name : " & txtLocalHost.Text
'clear the clipboard
Clipboard.Clear
'put the info into the clipboard
Clipboard.SetText sDat
sDat = Empty 'clean up the mess
End Sub

Private Sub lblGetRefresh_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call DoButton(Button, GetUp, GetDown)
End Sub

Private Sub lblGetRefresh_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call DoButton(Button, GetDown, GetUp)
Dim sTmp As String: sTmp = Empty
sTmp = Inet.OpenURL("http://www.ShowMyIP.com/xml", icString)
'get the source code from "http://www.ShowMyIP.com/xml" and store it into the sTmp variable
StatusBar.SimpleText = "Status : Gathering Information . . ."
Do
DoEvents 'start a loop
Loop Until Not Inet.StillExecuting 'keep doing 'nothing' until the inet control has gathered all of the source code
txtLocalIP.Text = Socket.LocalIP 'use the winsock's LocalIP method to show the local ip address
txtLocalHost.Text = Socket.LocalHostName 'use the winsock's LocalHostName method to show the local hostname
txtExtIP.Text = GetIP(sTmp) 'parse the IP from the source code, and place it into the appropriate textbox
txtDomain.Text = GetHost(sTmp) 'same thing for host
sTmp = Empty 'clean up
StatusBar.SimpleText = "Status : Information Gathered."
End Sub

Private Sub lblSave_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call DoButton(Button, SaveUp, SaveDown)
End Sub

Private Sub lblSave_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'does basically the same thing as the Copy routine
'except it just saves the info to a file using the WriteToFile() routine
Call DoButton(Button, SaveDown, SaveUp)
Dim sTmp As String, sDat As String: sTmp = Empty: sDat = Empty
With CD
.DialogTitle = "Save Output File"
.Filter = "All Supported Files|*.txt;*.usr;*.log;*.doc;*.rtf"
.ShowSave
If Len(.FileName) = 0 Then 'user didn't select a file, must have hit cancel, so exit the routine
Exit Sub
End If
sDat = "Local IP Address : " & txtLocalIP.Text & vbNewLine
sDat = sDat & "External IP Address : " & txtExtIP.Text & vbNewLine
sDat = sDat & "Host / Domain Name : " & txtDomain.Text & vbNewLine
sDat = sDat & "Local Hostname / Computer Name : " & txtLocalHost.Text
Call WriteToFile(.FileName, sDat)
sDat = Empty
End With
End Sub
