VERSION 5.00
Begin VB.UserControl OutlookPanel 
   Alignable       =   -1  'True
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   5250
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3240
   ScaleHeight     =   5250
   ScaleWidth      =   3240
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   705
      Index           =   2
      Left            =   5520
      Picture         =   "OutlookPanel.ctx":0000
      ScaleHeight     =   705
      ScaleWidth      =   2250
      TabIndex        =   6
      Top             =   1560
      Width           =   2250
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   705
      Index           =   1
      Left            =   5520
      Picture         =   "OutlookPanel.ctx":533E
      ScaleHeight     =   705
      ScaleWidth      =   2250
      TabIndex        =   5
      Top             =   840
      Width           =   2250
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   705
      Index           =   0
      Left            =   5520
      Picture         =   "OutlookPanel.ctx":A67C
      ScaleHeight     =   705
      ScaleWidth      =   2250
      TabIndex        =   4
      Top             =   120
      Width           =   2250
   End
   Begin VB.PictureBox picOuter 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2175
      Left            =   0
      ScaleHeight     =   2175
      ScaleWidth      =   2655
      TabIndex        =   2
      Top             =   0
      Width           =   2655
      Begin VB.PictureBox picInner 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1785
         Left            =   0
         ScaleHeight     =   1785
         ScaleWidth      =   2535
         TabIndex        =   3
         Top             =   0
         Width           =   2535
         Begin VB.PictureBox mItem 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   705
            Index           =   0
            Left            =   0
            Picture         =   "OutlookPanel.ctx":F9BA
            ScaleHeight     =   705
            ScaleWidth      =   2250
            TabIndex        =   7
            Top             =   0
            Visible         =   0   'False
            Width           =   2250
            Begin VB.Label mItemCaption 
               BackStyle       =   0  'Transparent
               Caption         =   "Caption"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   0
               Left            =   240
               TabIndex        =   9
               Top             =   30
               Width           =   1935
            End
            Begin VB.Label mItemText 
               BackStyle       =   0  'Transparent
               Caption         =   "Info"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Index           =   0
               Left            =   240
               TabIndex        =   8
               Top             =   240
               Width           =   1935
               WordWrap        =   -1  'True
            End
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No Items to Display"
            Height          =   195
            Left            =   600
            TabIndex        =   10
            Top             =   120
            Width           =   1365
         End
      End
   End
   Begin VB.VScrollBar VBar 
      Height          =   2175
      Left            =   2760
      TabIndex        =   1
      Top             =   0
      Width           =   255
   End
   Begin VB.HScrollBar HBar 
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   2280
      Width           =   2655
   End
   Begin VB.Menu mnufile 
      Caption         =   "mnyfile"
      Begin VB.Menu mnusort 
         Caption         =   "Filter by"
         Begin VB.Menu mnualertr 
            Caption         =   "Red"
         End
         Begin VB.Menu mnuinfo 
            Caption         =   "Blue"
         End
         Begin VB.Menu mnusuggest 
            Caption         =   "Green"
         End
         Begin VB.Menu mnurefresh 
            Caption         =   "None"
         End
      End
      Begin VB.Menu mnuline1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuDelete 
         Caption         =   "Delete"
      End
   End
End
Attribute VB_Name = "OutlookPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True



Option Explicit

Private Pan As Collection



Public Selected As String
Sub AddRecord(key As String, oCaption As String, oText As String, oType As mtype)

Dim hObj As ClsPanels
   
Set hObj = New ClsPanels
hObj.key = key
hObj.mCaption = oCaption
hObj.mInfo = oText
hObj.mColor = oType

Pan.Add hObj


mItemCaption(0).Caption = Pan.Count

End Sub


Private Function getindexfromkey(key As String) As Integer

Dim pobj                                As ClsPanels
Dim i As Integer
For i = 1 To Pan.Count
    Set pobj = Pan(i)

    If pobj.key = key Then

        getindexfromkey = i

    End If

Next i

End Function
Private Sub AddItems(Index As Integer)
Dim pobj As ClsPanels

Set pobj = Pan.Item(Index)

    ' Create the control.
    Index = mItem.Count
    Load mItem(Index)
    Load mItemCaption(Index)
    Load mItemText(Index)
   
    Set mItem(Index).Container = picInner
    Set mItemCaption(Index).Container = mItem(Index)
    Set mItemText(Index).Container = mItem(Index)

    ' Position the control.
    If Index = 1 Then
    mItem(Index).Top = 30
    
    Else
   mItem(Index).Top = mItem(Index - 1).Top + _
       mItem(Index - 1).Height + 30
    End If
    ' Change Color
    mItem(Index).Picture = pic(pobj.mColor).Picture
    'set text
    mItemCaption(Index).Caption = pobj.mCaption
    mItemText(Index).Caption = pobj.mInfo
    
    
    
    ' Size picInner to hold the control.
        picInner.Move 0, 0, _
        mItem(Index).Left + mItem(Index).Width + 120, _
        mItem(Index).Top + mItem(Index).Height + 120

    ' Display the control.
    mItem(Index).Visible = True
    Label1.Visible = False
    mItemCaption(Index).Visible = True
    mItemText(Index).Visible = True
    mItem(Index).Tag = pobj.key
    ' Rearrange the scroll bars.
   

    ArrangeScrollBars
End Sub


' Arrange the scroll bars.
Private Sub ArrangeScrollBars()
Dim have_wid As Single
Dim have_hgt As Single
Dim need_wid As Single
Dim need_hgt As Single
Dim need_hbar As Boolean
Dim need_vbar As Boolean
picOuter.Width = mItem(0).Width
picInner.Width = mItem(0).Width
UserControl.Width = picOuter.Width + VBar.Width + 25
    ' Don't bother if we're minimized.
   ' If WindowState = vbMinimized Then Exit Sub

    ' See how much room we need and
    ' how much room we have.
    need_wid = picInner.Width + (picOuter.Width - picOuter.ScaleWidth)
    need_hgt = picInner.Height + (picOuter.Height - picOuter.ScaleHeight)
    have_wid = ScaleWidth
    have_hgt = ScaleHeight

    ' See which scroll bars we need.
    need_hbar = (need_wid > have_wid)
    If need_hbar Then have_hgt = have_hgt - HBar.Height

    need_vbar = (need_hgt > have_hgt)
    If need_vbar Then
        ' This takes away a little width so we
        ' might need the horizontal scroll bar now.
        have_wid = have_wid - VBar.Width
        If Not need_hbar Then
            need_hbar = (need_wid > have_wid)
            If need_hbar Then have_hgt = have_hgt - HBar.Height
        End If
    End If

    ' Position the outer PictureBox leaving room
    ' for the scroll bars.
    picOuter.Move 0, 0, have_wid, have_hgt

    ' Position or hide the scroll bars.
    If need_hbar Then
        HBar.Move 0, have_hgt, have_wid
        HBar.Min = 0
        HBar.Max = picOuter.ScaleWidth - picInner.Width
        HBar.LargeChange = picOuter.ScaleWidth
        HBar.SmallChange = picOuter.ScaleWidth / 5
        HBar.Visible = True
    Else
        HBar.Visible = False
    End If

    If need_vbar Then
        VBar.Move have_wid, 0, VBar.Width, have_hgt
        VBar.Min = 0
        VBar.Max = picOuter.ScaleHeight - picInner.Height
        VBar.LargeChange = picOuter.ScaleHeight
        VBar.SmallChange = picOuter.ScaleHeight / 5
        VBar.Visible = True
    Else
        VBar.Visible = False
    End If
End Sub

Sub Populate()
Dim i As Integer
On Error Resume Next
'
'First unload all panels
'
For i = 1 To Pan.Count
       Unload mItemCaption(i)
       Unload mItemText(i)
       Unload mItem(i)
       
Next i
    
    
Dim pobj As ClsPanels




    For i = 1 To Pan.Count
        AddItems i
    Next i

End Sub

Sub Filter(FilterType As mtype)
Dim i As Integer
On Error Resume Next
'
'First unload all panels
'
For i = 1 To Pan.Count
       mItemCaption(i).Visible = False
       mItemText(i).Visible = False
       mItem(i).Visible = False
Next i
    
If Pan.Count = 0 Then
    Label1.Visible = False
Else
    Label1.Visible = True
End If
    
Dim pobj As ClsPanels
Dim oldindex As Integer
oldindex = 0


    For i = 1 To Pan.Count
    Set pobj = Pan.Item(i)
    DoEvents
       If FilterType = pobj.mColor Then
        mItemCaption(i).Visible = True
        mItemText(i).Visible = True
        mItem(i).Visible = True
        
        
        If oldindex = 0 Then
        mItem(i).Top = 30

        Else
            mItem(i).Top = mItem(oldindex).Top + _
            mItem(oldindex).Height + 30
        End If
        oldindex = i
            
       End If
       
    Next i
ArrangeScrollBars
End Sub
Sub RemoveRecord(key As String)
Dim X As Integer
X = getindexfromkey(key)
If X > 0 Then
    Dim i As Integer
        For i = 1 To Pan.Count
               Unload mItemCaption(i)
               Unload mItemText(i)
               Unload mItem(i)
        Next i
    Pan.Remove X
    Populate
End If


End Sub

Sub Clear()

Dim i As Integer

For i = 1 To Pan.Count
       Unload mItemCaption(i)
       Unload mItemText(i)
       Unload mItem(i)
       
Next i

For i = 1 To Pan.Count
Pan.Remove Pan.Count
Next i
Label1.Visible = True


End Sub

Private Sub HBar_Change()
picInner.Left = HBar.Value
End Sub


Private Sub HBar_Scroll()
picInner.Left = HBar.Value
End Sub


Private Sub Image1_Click()
Populate
End Sub

Private Sub mItem_Click(Index As Integer)
Selected = mItem(Index).Tag
End Sub

Private Sub mItem_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim pobj As ClsPanels
Dim i As Integer
    mnuinfo.Visible = False
    mnualertr.Visible = False
    mnusuggest.Visible = False


If Button = vbRightButton Then
 Selected = mItem(Index).Tag
 
  For i = 1 To Pan.Count
    Set pobj = Pan.Item(i)
    DoEvents
        If pobj.mColor = info Then
            mnuinfo.Visible = True
        ElseIf pobj.mColor = Alert Then
            mnualertr.Visible = True
        ElseIf pobj.mColor = Suggest Then
            mnusuggest.Visible = True
        End If
  Next i
 
 
 UserControl.PopupMenu mnufile
 
End If

End Sub


Private Sub mItemCaption_Click(Index As Integer)
Selected = mItem(Index).Tag

End Sub


Private Sub mItemCaption_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim pobj As ClsPanels
Dim i As Integer
    mnuinfo.Visible = False
    mnualertr.Visible = False
    mnusuggest.Visible = False


If Button = vbRightButton Then
 Selected = mItem(Index).Tag
 
  For i = 1 To Pan.Count
    Set pobj = Pan.Item(i)
    DoEvents
        If pobj.mColor = info Then
            mnuinfo.Visible = True
        ElseIf pobj.mColor = Alert Then
            mnualertr.Visible = True
        ElseIf pobj.mColor = Suggest Then
            mnusuggest.Visible = True
        End If
  Next i
 
 
 UserControl.PopupMenu mnufile
 
End If

End Sub


Private Sub mItemText_Click(Index As Integer)
Selected = mItem(Index).Tag

End Sub


Private Sub mItemText_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim pobj As ClsPanels
Dim i As Integer
    mnuinfo.Visible = False
    mnualertr.Visible = False
    mnusuggest.Visible = False


If Button = vbRightButton Then
 Selected = mItem(Index).Tag
 
  For i = 1 To Pan.Count
    Set pobj = Pan.Item(i)
    DoEvents
        If pobj.mColor = info Then
            mnuinfo.Visible = True
        ElseIf pobj.mColor = Alert Then
            mnualertr.Visible = True
        ElseIf pobj.mColor = Suggest Then
            mnusuggest.Visible = True
        End If
  Next i
 
 
 UserControl.PopupMenu mnufile
 
End If
End Sub


Private Sub mnualertr_Click()
Filter Alert
End Sub

Private Sub MnuDelete_Click()
RemoveRecord Selected
End Sub

Private Sub mnuinfo_Click()
Filter info
End Sub

Private Sub mnurefresh_Click()
Populate
End Sub

Private Sub mnusuggest_Click()
Filter Suggest
End Sub

Private Sub UserControl_Initialize()


Set Pan = New Collection




End Sub

Private Sub UserControl_Resize()
ArrangeScrollBars
End Sub

Private Sub VBar_Change()
picInner.Top = VBar.Value
End Sub


Private Sub VBar_Scroll()
picInner.Top = VBar.Value
End Sub


