VERSION 5.00
Object = "*\A..\vbalDTab6.vbp"
Begin VB.Form frmDocument 
   Caption         =   "Form1"
   ClientHeight    =   3435
   ClientLeft      =   6015
   ClientTop       =   5625
   ClientWidth     =   4770
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3435
   ScaleWidth      =   4770
   Begin vbalDTab6X.vbalDTabControlX TabLeft 
      Align           =   3  'Align Left
      Height          =   3435
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   6059
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty SelectedFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Pinnable        =   -1  'True
      Begin VB.PictureBox picEvents 
         Height          =   2295
         Left            =   240
         ScaleHeight     =   2235
         ScaleWidth      =   1755
         TabIndex        =   2
         Top             =   600
         Visible         =   0   'False
         Width           =   1815
         Begin VB.ListBox lstEvents 
            Height          =   2460
            IntegralHeight  =   0   'False
            Left            =   0
            TabIndex        =   3
            Top             =   0
            Width           =   1635
         End
      End
   End
   Begin VB.TextBox txtDocument 
      Height          =   3435
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   60
      Width           =   4335
   End
End
Attribute VB_Name = "frmDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub LogEvent(ByVal sMsg As String)
   lstEvents.AddItem sMsg
   lstEvents.ListIndex = lstEvents.NewIndex
End Sub

Private Sub Form_Load()
Dim tabX As cTab
   TabLeft.ImageList = mfrmMain.ilsIcons
   Set tabX = TabLeft.Tabs.Add("EVENTS", , "Events")
   tabX.Panel = Me.picEvents
   
   
End Sub

Private Sub Form_Resize()
   On Error Resume Next
   If Not (StrComp(TabLeft.Tag, "CLOSE")) = 0 Then
      txtDocument.Left = TabLeft.Left + TabLeft.Width + 2 * Screen.TwipsPerPixelX
   End If
   txtDocument.Move txtDocument.Left, txtDocument.Top, Me.ScaleWidth - txtDocument.Left - 2 * Screen.TwipsPerPixelX, Me.ScaleHeight - txtDocument.Top - 2 * Screen.TwipsPerPixelY
End Sub

Private Sub picEvents_Resize()
   lstEvents.Move 2 * Screen.TwipsPerPixelX, 2 * Screen.TwipsPerPixelY, picEvents.ScaleWidth - 2 * Screen.TwipsPerPixelX, picEvents.ScaleHeight - 2 * Screen.TwipsPerPixelY
End Sub

Private Sub tabLeft_Pinned()
   LogEvent "Pinned"
   Form_Resize
End Sub

Private Sub tabLeft_Resize()
   LogEvent "Resize"
End Sub

Private Sub tabLeft_TabBarClick(ByVal iButton As MouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Single, ByVal y As Single)
   LogEvent "TabBarClick"
End Sub

Private Sub tabLeft_TabClick(theTab As cTab, ByVal iButton As MouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Single, ByVal y As Single)
   LogEvent "TabClick"
End Sub

Private Sub tabLeft_TabClose(theTab As cTab, bCancel As Boolean)
   LogEvent "TabClose"
   TabLeft.Tag = "Close"
   TabLeft.Visible = False
   txtDocument.Left = 2 * Screen.TwipsPerPixelX
   Form_Resize
End Sub

Private Sub tabLeft_TabDoubleClick(theTab As cTab)
   LogEvent "TabDoubleClick"
End Sub

Private Sub tabLeft_TabSelected(theTab As cTab)
   LogEvent "TabSelected"
End Sub

Private Sub tabLeft_UnPinned()
   LogEvent "Unpinned"
   Form_Resize
End Sub
