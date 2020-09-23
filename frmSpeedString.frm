VERSION 5.00
Begin VB.Form frmSpeedString 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SpeedString Demo"
   ClientHeight    =   3885
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5025
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3885
   ScaleWidth      =   5025
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "Bench"
      Height          =   735
      Left            =   120
      TabIndex        =   10
      Top             =   3000
      Width           =   4815
      Begin VB.CommandButton Command1 
         Caption         =   "&Run"
         Height          =   375
         Left            =   3360
         TabIndex        =   14
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   840
         TabIndex        =   12
         Text            =   "10000"
         Top             =   300
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "times."
         Height          =   255
         Left            =   1680
         TabIndex        =   13
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Execute"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "SpeedString concatenation"
      Height          =   1335
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   4815
      Begin VB.PictureBox picBar2 
         Height          =   255
         Left            =   120
         ScaleHeight     =   195
         ScaleWidth      =   4515
         TabIndex        =   6
         Top             =   960
         Width           =   4575
         Begin VB.PictureBox bar2 
            BackColor       =   &H00C00000&
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   0
            ScaleHeight     =   255
            ScaleWidth      =   135
            TabIndex        =   7
            Top             =   0
            Visible         =   0   'False
            Width           =   135
         End
      End
      Begin VB.Label Label4 
         Caption         =   "Syntax: SpeedString.Append ""Some text string"", 16"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   4575
      End
      Begin VB.Label lblResult2 
         Caption         =   "Execute time: 0 ms"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   4575
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Visual Basic string concatenation"
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4815
      Begin VB.PictureBox picBar1 
         Height          =   255
         Left            =   120
         ScaleHeight     =   195
         ScaleWidth      =   4515
         TabIndex        =   2
         Top             =   960
         Width           =   4575
         Begin VB.PictureBox bar1 
            BackColor       =   &H00C00000&
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   0
            ScaleHeight     =   255
            ScaleWidth      =   135
            TabIndex        =   3
            Top             =   0
            Visible         =   0   'False
            Width           =   135
         End
      End
      Begin VB.Label lblResult1 
         Caption         =   "Execute time: 0 ms"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   4575
      End
      Begin VB.Label Label1 
         Caption         =   "Syntax: strData = strData && ""Some text string"""
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   4575
      End
   End
End
Attribute VB_Name = "frmSpeedString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Sub Command1_Click()

    Dim Count As Long, Times As Long
    
    Dim timeStart As Long
    Dim stringValue As String
    Dim stringLength As Long

    Dim sBuffer As String
    Dim SpeedString As New cSpeedString

    Dim result1 As Long, result2 As Long

    Command1.Enabled = False
        
    Times = Val(Text1.Text)
        
    stringValue = "Some text string"
    stringLength = Len(stringValue)
    
    bar1.Visible = False
    bar2.Visible = False
    
    Me.Caption = "VB contatenation ..."
    DoEvents
    
    ' Do the VB Concatenation
    timeStart = GetTickCount
    For Count = 1 To Times
        sBuffer = sBuffer & stringValue
    Next
    result1 = GetTickCount - timeStart
    
    Me.Caption = "SpeedString contatenation ..."
    DoEvents
    
    ' Do the SpeedString concatenation
    timeStart = GetTickCount
    For Count = 1 To Times
        SpeedString.Append stringValue, stringLength
    Next
    result2 = GetTickCount - timeStart
    
    bar1.Width = picBar1.ScaleWidth * 0.99
    bar2.Width = picBar1.Width * (result2 / result1)
    
    bar1.Visible = True
    bar2.Visible = True
    
    lblResult1.Caption = "Execute time: " & CStr(result1) & " ms"
    lblResult2.Caption = "Execute time: " & CStr(result2) & " ms"
        
    Me.Caption = "SpeedString Demo"
    DoEvents
        
    Command1.Enabled = True

End Sub
