VERSION 5.00
Object = "{19C5C5DE-CF81-4D89-AF58-7AB0B4CE3293}#1.0#0"; "Opanel.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5430
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8010
   LinkTopic       =   "Form1"
   ScaleHeight     =   5430
   ScaleWidth      =   8010
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   240
      TabIndex        =   4
      Text            =   "none"
      Top             =   4680
      Width           =   1815
   End
   Begin VB.CommandButton Command4 
      Caption         =   "removeall"
      Height          =   735
      Left            =   360
      TabIndex        =   3
      Top             =   3360
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Remove"
      Height          =   735
      Left            =   1200
      TabIndex        =   2
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add"
      Height          =   615
      Left            =   480
      TabIndex        =   1
      Top             =   960
      Width           =   1095
   End
   Begin Opanel.OutlookPanel OutlookPanel1 
      Align           =   4  'Align Right
      Height          =   5430
      Left            =   5475
      TabIndex        =   0
      Top             =   0
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   9578
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   5760
      TabIndex        =   5
      Top             =   480
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Combo1_Click()
If Combo1.Text = "Alert" Then
    OutlookPanel1.Filter Alert
ElseIf Combo1.Text = "Info" Then
    OutlookPanel1.Filter info
ElseIf Combo1.Text = "none" Then
    OutlookPanel1.Populate
ElseIf Combo1.Text = "Suggest" Then
    OutlookPanel1.Filter Suggest
End If
End Sub

Private Sub Command1_Click()
OutlookPanel1.Clear


For i = 1 To 5

Me.OutlookPanel1.AddRecord "Test" & i, "caption " & i, "test string in the info" & i, info
Me.OutlookPanel1.AddRecord "Test" & i, "caption " & i, "test string in the info" & i, Alert
'Me.OutlookPanel1.AddRecord "Test" & i, "caption " & i, "test string in the info" & i, Suggest

Next i


OutlookPanel1.Populate


Combo1.AddItem "Info"
Combo1.AddItem "Suggest"
Combo1.AddItem "Alert"
Combo1.AddItem "none"


End Sub

Private Sub Command2_Click()
Me.OutlookPanel1.RemoveRecord OutlookPanel1.Selected

End Sub




Private Sub Command4_Click()
OutlookPanel1.Clear
End Sub

