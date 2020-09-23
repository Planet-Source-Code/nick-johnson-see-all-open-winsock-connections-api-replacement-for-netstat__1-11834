VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "Netstat"
   ClientHeight    =   3240
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   4680
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3240
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtUpdate 
      Height          =   285
      Left            =   1800
      TabIndex        =   1
      Text            =   "1000"
      Top             =   120
      Width           =   1215
   End
   Begin MSComctlLib.ImageList imlMain 
      Left            =   2040
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0442
            Key             =   "Closed"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0556
            Key             =   "Listening"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":09AA
            Key             =   "SYN Sent"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0CC6
            Key             =   "SYN Recieved"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0FE2
            Key             =   "Established"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1436
            Key             =   "FIN Wait 1"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":188A
            Key             =   "FIN Wait 2"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1CDE
            Key             =   "Close Wait"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2132
            Key             =   "Closing"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2586
            Key             =   "Last ACK"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":29DA
            Key             =   "Time Wait"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2E2E
            Key             =   "Other"
         EndProperty
      EndProperty
   End
   Begin VB.Timer tmrRefresh 
      Interval        =   1000
      Left            =   1800
      Top             =   1320
   End
   Begin MSComctlLib.ListView lvMain 
      Height          =   2775
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   4895
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   393217
      Icons           =   "imlMain"
      SmallIcons      =   "imlMain"
      ColHdrIcons     =   "imlMain"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "Local"
         Text            =   "Local Port"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "Remote"
         Text            =   "Remote Address"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "Host"
         Text            =   "Remote Hostname"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Update Frequency (ms):"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   120
      Width           =   1815
   End
   Begin VB.Menu mnuContext 
      Caption         =   "&Context"
      Visible         =   0   'False
      Begin VB.Menu mnuContextKill 
         Caption         =   "&Kill"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This code is copyright 2000 Nick Johnson.
'This code may be reused and modified for non-commercial
'purposes only as long as credit is given to the author
'in the programmes about box and it's documentation.
'If you use this code, please email me at:
'arachnid@mad.scientist.com and let me know what you think
'and what you are doing with it.

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode <> vbAppWindows Then
        MsgBox "Netstat code is copyright 2000, Nick Johnson. This code may be reused for non commercial purposes on condition that credit is given to the author in the programmes about box and documentation.", vbInformation + vbOKOnly
    End If
End Sub

Private Sub Form_Resize()
    Dim a As Integer
    lvMain.Width = lvMain.Parent.Width - 100
    lvMain.Height = lvMain.Parent.Height - 850
    
    For a = 2 To lvMain.ColumnHeaders.Count
        lvMain.ColumnHeaders(a).Width = (frmMain.Width - 100) / (lvMain.ColumnHeaders.Count - 1) - 600
    Next a
End Sub

Private Sub lvMain_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        If ipsMain.RowData(lvMain.SelectedItem.Tag).State = TCP_STATE_ESTAB Then
            mnuContextKill.Enabled = True
        Else
            mnuContextKill.Enabled = False
        End If
        frmMain.PopupMenu mnuContext
    End If
End Sub

Private Sub mnuContextKill_Click()
    ipsMain.RowData(lvMain.SelectedItem.Tag).Kill
End Sub

Private Sub tmrRefresh_Timer()
    Dim a As Integer
    Dim intLVPtr As Integer
    
    ipsMain.getTCPConnections
    
    'Update routine - if the existing entry is the same as this one, leave it, otherwise overwrite it.
    intlvpointer = 0
    For a = 0 To ipsMain.RowCount - 1
        If ipsMain.RowData(a).State <> TCP_STATE_LISTEN Then
            intLVPtr = intLVPtr + 1
            'If we are past the bounds of the current array, add a new line
            If intLVPtr > lvMain.ListItems.Count Then
                lvMain.ListItems.Add , , ipsMain.RowData(a).LocalPort, , ipsMain.RowData(a).StateText
                lvMain.ListItems(intLVPtr).ToolTipText = ipsMain.RowData(a).StateText
                lvMain.ListItems(lvMain.ListItems.Count).ListSubItems.Add , , ipsMain.RowData(a).RemoteIPString & ":" & ipsMain.RowData(a).RemotePort
                lvMain.ListItems(lvMain.ListItems.Count).ListSubItems.Add , , "Retrieving..."
                lvMain.Refresh
                lvMain.ListItems(lvMain.ListItems.Count).ListSubItems(2).Text = iphDNS.AddressToName(ipsMain.RowData(a).RemoteIPString)
                lvMain.ListItems(lvMain.ListItems.Count).Tag = a
            Else
                'We are still in the bounds. If the current
                'entry equals the one to insert, just change
                'the icon. Otherwise, overwrite it.
                If lvMain.ListItems(intLVPtr).Text = ipsMain.RowData(a).LocalPort And lvMain.ListItems(intLVPtr).ListSubItems(1).Text = ipsMain.RowData(a).RemoteIPString & ":" & ipsMain.RowData(a).RemotePort And lvMain.ListItems(intLVPtr).Tag = a Then
                    'lvMain.ListItems(intLVPtr).SmallIcon = ipsMain.RowData(a).StateText
                    If lvMain.ListItems(intLVPtr).SmallIcon <> ipsMain.RowData(a).StateText Then
                        lvMain.ListItems(intLVPtr).SmallIcon = ipsMain.RowData(a).StateText
                        lvMain.ListItems(intLVPtr).ToolTipText = ipsMain.RowData(a).StateText
                    End If
                Else
                    'Different, overwrite it.
                    lvMain.ListItems(intLVPtr).Text = ipsMain.RowData(a).LocalPort
                    lvMain.ListItems(intLVPtr).ListSubItems(1).Text = ipsMain.RowData(a).RemoteIPString & ":" & ipsMain.RowData(a).RemotePort
                    lvMain.ListItems(lvMain.ListItems.Count).ListSubItems(2).Text = "Retrieving..."
                    lvMain.Refresh
                    lvMain.ListItems(lvMain.ListItems.Count).ListSubItems(2).Text = iphDNS.AddressToName(ipsMain.RowData(a).RemoteIPString)
                    lvMain.ListItems(intLVPtr).Tag = a
                    lvMain.ListItems(intLVPtr).SmallIcon = ipsMain.RowData(a).StateText
                    lvMain.ListItems(intLVPtr).ToolTipText = ipsMain.RowData(a).StateText
                End If
            End If
        End If
    Next a
    
    'If there are more listitem entries than connections, kill the extra ones.
    For a = lvMain.ListItems.Count To intLVPtr + 1 Step -1
        lvMain.ListItems.Remove a
    Next a
End Sub

Private Sub txtUpdate_Change()
    tmrRefresh.Interval = Val(txtUpdate.Text)
End Sub
