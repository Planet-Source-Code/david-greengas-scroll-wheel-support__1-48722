VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmScroll 
   Caption         =   "Form1"
   ClientHeight    =   5280
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7035
   LinkTopic       =   "Form1"
   ScaleHeight     =   5280
   ScaleWidth      =   7035
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   1455
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "frmScroll.frx":0000
      Top             =   120
      Width           =   3735
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   3015
      Left            =   240
      TabIndex        =   0
      Top             =   2040
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   5318
      _Version        =   393216
   End
End
Attribute VB_Name = "frmScroll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
  Dim lRow As Long
  Dim lCol As Long
  
  TrackMouseWheel Text1.hWnd
  TrackMouseWheel MSFlexGrid1.hWnd
  
  MSFlexGrid1.FormatString = "col 1| col 2| col 3 | col 4"
  With MSFlexGrid1
    .Redraw = False
    For lRow = 1 To 100
      .AddItem vbNullString
      For lCol = 0 To 3
        .TextMatrix(lRow, lCol) = lRow & ", " & lCol
      Next
    Next
    .Redraw = True
  End With
End Sub

Private Sub Form_Resize()
  'Dim rRect As RECT
  
  'Call GetClientRect(Me.hwnd, rRect)
  'Debug.Print "Form Rect:", rRect.Left, rRect.Top, rRect.Right, rRect.Bottom

End Sub

Private Sub Form_Unload(Cancel As Integer)
 UnTrackMouseWheel Text1.hWnd
 UnTrackMouseWheel MSFlexGrid1.hWnd
End Sub
