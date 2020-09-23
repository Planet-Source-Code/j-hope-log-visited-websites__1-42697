VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "IE Captions"
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrScroll 
      Interval        =   1500
      Left            =   120
      Top             =   480
   End
   Begin VB.Timer tmrRecord 
      Interval        =   1000
      Left            =   120
      Top             =   120
   End
   Begin VB.ListBox lstURLS 
      Height          =   2205
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4455
   End
   Begin VB.Label lblCaption 
      Alignment       =   2  'Center
      Caption         =   "By Jeremy Hope"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   2520
      Width           =   4455
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Dim OldCaption As String


Private Sub tmrRecord_Timer()
Dim lIEHandle As Long
Dim Name As String * 256

lIEHandle = FindWindow("IEFrame", vbNullString)

If lIEHandle <> 0 Then
    GetWindowText lIEHandle, Name, 256
        If Name <> OldCaption Then
            lstURLS.AddItem Name
            OldCaption = Name
        End If
        
End If



End Sub

