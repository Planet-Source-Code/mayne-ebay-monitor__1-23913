VERSION 5.00
Begin VB.Form frmOpen 
   BackColor       =   &H00DE8D21&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "EBAY Monitor -- Item Number:"
   ClientHeight    =   1305
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   2925
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1305
   ScaleWidth      =   2925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00DE8D21&
      Caption         =   " Ebay Item Number && Refresh Rate: "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1305
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2925
      Begin VB.CommandButton cmdOK 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Open Item"
         Height          =   345
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   870
         Width           =   1455
      End
      Begin VB.ComboBox cmbRefresh 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmOpen.frx":0000
         Left            =   1470
         List            =   "frmOpen.frx":0010
         TabIndex        =   4
         Text            =   "1 Min"
         Top             =   510
         Width           =   915
      End
      Begin VB.TextBox txtItem 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1350
         MaxLength       =   10
         TabIndex        =   2
         Top             =   270
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Refresh Rate:"
         Height          =   195
         Left            =   420
         TabIndex        =   3
         Top             =   540
         Width           =   1035
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Item Number:"
         Height          =   255
         Left            =   360
         TabIndex        =   1
         Top             =   270
         Width           =   945
      End
   End
End
Attribute VB_Name = "frmOpen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
    Dim nURL As String, nRefresh, nCount As Integer, I As Integer
    'Set URL
    nURL = "http://cgi.ebay.com/aw-cgi/eBayISAPI.dll?ViewItem&item=" & txtItem.Text
    'Set Refresh Rate
    Select Case cmbRefresh.Text
        Case "1 Min"
            nRefresh = 60000
        Case "45 Sec"
            nRefresh = 45000
        Case "30 Sec"
            nRefresh = 30000
        Case "15 Sec"
            nRefresh = 15000
    End Select
    If frmMain.SSTab1.TabCaption(0) = "Item" Then
        'The first tab is blank so use first webbrowser
        frmMain.WebBrowser1.Navigate nURL
        frmMain.SSTab1.TabCaption(0) = "Item " & txtItem.Text
        frmMain.SSTab1.Tab = 0
        frmMain.Timer1.Interval = nRefresh
        frmMain.Timer1.Enabled = True
        frmMain.WebBrowser1.Visible = True
        frmMain.WebBrowser2.Visible = False
        frmMain.WebBrowser3.Visible = False
        frmMain.WebBrowser4.Visible = False
        frmMain.WebBrowser5.Visible = False
    Else
        'The first tab is being used so use the next available tab
        For I = 0 To 4
            If frmMain.SSTab1.TabCaption(I) = "Blank Item" Then
                nCount = I
                If I <> 4 Then I = 4
            End If
        Next I
        Select Case nCount
            Case 1 'Second tab is the next available
                frmMain.WebBrowser2.Navigate nURL
                frmMain.Timer2.Interval = nRefresh
                frmMain.Timer2.Enabled = True
                frmMain.SSTab1.TabCaption(1) = "Item " & txtItem.Text
                frmMain.SSTab1.Tab = 1
                frmMain.WebBrowser1.Visible = False
                frmMain.WebBrowser2.Visible = True
                frmMain.WebBrowser3.Visible = False
                frmMain.WebBrowser4.Visible = False
                frmMain.WebBrowser5.Visible = False
            Case 2 'Third tab is the next available
                frmMain.WebBrowser3.Navigate nURL
                frmMain.Timer3.Interval = nRefresh
                frmMain.Timer3.Enabled = True
                frmMain.SSTab1.TabCaption(2) = "Item " & txtItem.Text
                frmMain.SSTab1.Tab = 2
                frmMain.WebBrowser1.Visible = False
                frmMain.WebBrowser2.Visible = False
                frmMain.WebBrowser3.Visible = True
                frmMain.WebBrowser4.Visible = False
                frmMain.WebBrowser5.Visible = False
            Case 3 'Fourth tab is the next available
                frmMain.WebBrowser4.Navigate nURL
                frmMain.Timer4.Interval = nRefresh
                frmMain.Timer4.Enabled = True
                frmMain.SSTab1.TabCaption(3) = "Item " & txtItem.Text
                frmMain.SSTab1.Tab = 3
                frmMain.WebBrowser1.Visible = False
                frmMain.WebBrowser2.Visible = False
                frmMain.WebBrowser3.Visible = False
                frmMain.WebBrowser4.Visible = True
                frmMain.WebBrowser5.Visible = False
            Case 4 'Fifth tab is the next available
                frmMain.WebBrowser5.Navigate nURL
                frmMain.Timer5.Interval = nRefresh
                frmMain.Timer5.Enabled = True
                frmMain.SSTab1.TabCaption(4) = "Item " & txtItem.Text
                frmMain.SSTab1.Tab = 4
                frmMain.WebBrowser1.Visible = False
                frmMain.WebBrowser2.Visible = False
                frmMain.WebBrowser3.Visible = False
                frmMain.WebBrowser4.Visible = False
                frmMain.WebBrowser5.Visible = True
            Case Else
                MsgBox "There was an unexpected error." & vbCrLf & "There are probably no available watch windows.", vbExclamation + vbOKOnly, "Ebay Monitor"
        End Select
    End If
    Unload Me
    frmMain.Show
End Sub
