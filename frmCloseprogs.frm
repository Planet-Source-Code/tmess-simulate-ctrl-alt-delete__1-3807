VERSION 5.00
Begin VB.Form frmCloseProgram 
   Caption         =   "Close Program"
   ClientHeight    =   5160
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6390
   Icon            =   "frmCloseprogs.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   6390
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frmPath 
      Caption         =   "Display Path"
      Height          =   615
      Left            =   4200
      TabIndex        =   5
      Top             =   360
      Width           =   2055
      Begin VB.OptionButton optNo 
         Caption         =   "No"
         Height          =   255
         Left            =   1080
         TabIndex        =   7
         Top             =   240
         Width           =   615
      End
      Begin VB.OptionButton optYes 
         BackColor       =   &H8000000A&
         Caption         =   "Yes"
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.ListBox lstItems 
      Height          =   3180
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   6135
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   495
      Left            =   4800
      TabIndex        =   2
      Top             =   4440
      Width           =   1355
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "End Task"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2520
      TabIndex        =   1
      Top             =   4440
      Width           =   1355
   End
   Begin VB.CommandButton cmdCloseProgram 
      Caption         =   "List Active Programs"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Label lblActiveProg 
      Caption         =   "Active Programs"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   1815
   End
End
Attribute VB_Name = "frmCloseProgram"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Integer
Dim NumOfProcess As Long
Dim objActiveProcess As clsActiveProcess

Private Sub Form_Load()

    Set objActiveProcess = New clsActiveProcess
    Me.optYes = True
           
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set objActiveProcess = Nothing
    
End Sub

Private Sub cmdCloseProgram_Click()
On Error Resume Next

    lstItems.Clear
    
    NumOfProcess = objActiveProcess.GetActiveProcess
    
        For i = 0 To NumOfProcess
            If optYes = True Then
             lstItems.AddItem objActiveProcess.szExeFile(i)
             lstItems.ItemData(i) = objActiveProcess.th32ProcessID(i)
               Else
             fEnumWindows objActiveProcess.th32ProcessID(i)
             If i = 0 Then IsResond = "Responding"
             lstItems.AddItem stripPath(objActiveProcess.szExeFile(i)) & " " & IsResond
             lstItems.ItemData(i) = objActiveProcess.th32ProcessID(i)
            End If
        Next
        
  
        If optYes Then
            lstItems.RemoveItem 0
        End If
        
    lstItems.Selected(0) = True
    Me.cmdStop.Enabled = True
    
End Sub

Private Sub cmdStop_Click()
Dim lPid     As Long
Dim lProcess As Long
Dim lReturn  As Long

    lPid = lstItems.ItemData(lstItems.ListIndex)

' Terminate the application unconditionally.
    lProcess = OpenProcess(PROCESS_ALL_ACCESS, 0&, lPid)
    lReturn = TerminateProcess(lProcess, 0&)
    
    lstItems.RemoveItem lstItems.ListIndex
    lstItems.Selected(0) = True
    
End Sub

Private Sub cmdQuit_Click()

    Unload Me
    
End Sub

Private Function stripPath(path As String) As String
'strip the path from the string

Dim holdval As Integer, holdString As String

    holdval = InStr(1, path, Chr(0)) - 1
       
    holdString = StrReverse(Left(path, holdval))
    holdString = Left(holdString, (InStr(1, holdString, "\") - 1))
    holdString = StrReverse(holdString)
    
    stripPath = holdString
    
    
End Function


