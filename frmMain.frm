VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00DE8D21&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ebay Monitor"
   ClientHeight    =   5130
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   8985
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   8985
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Left            =   4200
      Top             =   1620
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Left            =   3120
      Top             =   1440
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Left            =   2220
      Top             =   1470
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Left            =   1470
      Top             =   1500
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   630
      Top             =   1470
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8955
      _ExtentX        =   15796
      _ExtentY        =   661
      _Version        =   393216
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Item"
      TabPicture(0)   =   "frmMain.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "Blank Item"
      TabPicture(1)   =   "frmMain.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Blank Item"
      TabPicture(2)   =   "frmMain.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      TabCaption(3)   =   "Blank Item"
      TabPicture(3)   =   "frmMain.frx":0060
      Tab(3).ControlEnabled=   0   'False
      Tab(3).ControlCount=   0
      TabCaption(4)   =   "Blank Item"
      TabPicture(4)   =   "frmMain.frx":007C
      Tab(4).ControlEnabled=   0   'False
      Tab(4).ControlCount=   0
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser2 
      Height          =   4755
      Left            =   0
      TabIndex        =   2
      Top             =   330
      Width           =   8955
      ExtentX         =   15796
      ExtentY         =   8387
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser3 
      Height          =   4755
      Left            =   0
      TabIndex        =   3
      Top             =   330
      Width           =   8955
      ExtentX         =   15796
      ExtentY         =   8387
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser4 
      Height          =   4755
      Left            =   0
      TabIndex        =   4
      Top             =   330
      Width           =   8955
      ExtentX         =   15796
      ExtentY         =   8387
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser5 
      Height          =   4755
      Left            =   0
      TabIndex        =   5
      Top             =   330
      Width           =   8955
      ExtentX         =   15796
      ExtentY         =   8387
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   4755
      Left            =   0
      TabIndex        =   1
      Top             =   330
      Width           =   8955
      ExtentX         =   15796
      ExtentY         =   8387
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   4755
      Left            =   30
      TabIndex        =   6
      Top             =   330
      Width           =   8925
   End
   Begin VB.Menu mnuRightClick 
      Caption         =   "&File"
      Begin VB.Menu mnuStop 
         Caption         =   "&Stop Watching"
      End
      Begin VB.Menu mnuWatch 
         Caption         =   "&Watch Item..."
      End
      Begin VB.Menu a 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "&Close"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then 'Right Click
        Me.PopupMenu mnuRightClick
    End If
End Sub

Private Sub mnuClose_Click()
    Timer1.Enabled = False
    Timer2.Enabled = False
    Timer3.Enabled = False
    Timer4.Enabled = False
    Timer5.Enabled = False
    Unload Me
End Sub

Private Sub mnuStop_Click()
    'Determine which item to stop watching
    Select Case SSTab1.Tab
        Case 0 'First tab
            WebBrowser1.Navigate App.Path & "\none.html"
            Timer1.Enabled = False
            SSTab1.TabCaption(0) = "Blank Item"
        Case 1
            WebBrowser2.Navigate App.Path & "\none.html"
            Timer2.Enabled = False
            SSTab1.TabCaption(1) = "Blank Item"
        Case 2
            WebBrowser3.Navigate App.Path & "\none.html"
            Timer3.Enabled = False
            SSTab1.TabCaption(2) = "Blank Item"
        Case 3
            WebBrowser4.Navigate App.Path & "\none.html"
            Timer4.Enabled = False
            SSTab1.TabCaption(3) = "Blank Item"
        Case 4
            WebBrowser5.Navigate App.Path & "\none.html"
            Timer5.Enabled = False
            SSTab1.TabCaption(4) = "Blank Item"
    End Select
End Sub

Private Sub mnuWatch_Click()
    frmOpen.Show
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
        Select Case SSTab1.Tab
            Case 0 'First Tab
                WebBrowser1.Visible = True
                WebBrowser2.Visible = False
                WebBrowser3.Visible = False
                WebBrowser4.Visible = False
                WebBrowser5.Visible = False
            Case 1 'Second Tab
                WebBrowser1.Visible = False
                WebBrowser2.Visible = True
                WebBrowser3.Visible = False
                WebBrowser4.Visible = False
                WebBrowser5.Visible = False
            Case 2 'Third Tab
                WebBrowser1.Visible = False
                WebBrowser2.Visible = False
                WebBrowser3.Visible = True
                WebBrowser4.Visible = False
                WebBrowser5.Visible = False
            Case 3 'Fourth Tab
                WebBrowser1.Visible = False
                WebBrowser2.Visible = False
                WebBrowser3.Visible = False
                WebBrowser4.Visible = True
                WebBrowser5.Visible = False
            Case 4 'Last Tab
                WebBrowser1.Visible = False
                WebBrowser2.Visible = False
                WebBrowser3.Visible = False
                WebBrowser4.Visible = False
                WebBrowser5.Visible = True
        End Select
End Sub

Private Sub Timer1_Timer()
    WebBrowser1.Refresh
End Sub

Private Sub Timer2_Timer()
    WebBrowser2.Refresh
End Sub

Private Sub Timer3_Timer()
    WebBrowser3.Refresh
End Sub

Private Sub Timer4_Timer()
    WebBrowser4.Refresh
End Sub

Private Sub Timer5_Timer()
    WebBrowser5.Refresh
End Sub
