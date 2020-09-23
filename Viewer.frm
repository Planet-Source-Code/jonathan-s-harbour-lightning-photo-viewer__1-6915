VERSION 5.00
Begin VB.Form frmViewer 
   Caption         =   "Lightning Photo Viewer"
   ClientHeight    =   7590
   ClientLeft      =   1875
   ClientTop       =   1605
   ClientWidth     =   9750
   Icon            =   "Viewer.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7590
   ScaleWidth      =   9750
   Begin VB.CheckBox chkLoop 
      Caption         =   "Continuous Loop"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   180
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   7155
      Width           =   2085
   End
   Begin VB.ListBox List2 
      Height          =   255
      Left            =   90
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   15
      Top             =   810
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Timer Timer1 
      Left            =   675
      Top             =   810
   End
   Begin VB.VScrollBar vsSpeed 
      Height          =   465
      Left            =   2070
      Max             =   1
      Min             =   10
      TabIndex        =   12
      Top             =   6615
      Value           =   1
      Width           =   195
   End
   Begin VB.ListBox List1 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2700
      Left            =   45
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   11
      Top             =   540
      Width           =   2265
   End
   Begin VB.CommandButton cmdSlideshow 
      Caption         =   "&Slideshow"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   180
      TabIndex        =   8
      Top             =   6615
      Width           =   1320
   End
   Begin VB.PictureBox picContainer 
      Height          =   6885
      Left            =   2430
      ScaleHeight     =   455
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   479
      TabIndex        =   3
      Top             =   135
      Width           =   7245
      Begin VB.CommandButton cmdEnd 
         Height          =   240
         Left            =   6840
         TabIndex        =   10
         Top             =   6480
         Width           =   240
      End
      Begin VB.PictureBox picSize 
         Height          =   6270
         Left            =   45
         ScaleHeight     =   6210
         ScaleWidth      =   6615
         TabIndex        =   9
         Top             =   45
         Width           =   6675
         Begin VB.Image picMain 
            Height          =   6090
            Left            =   0
            Top             =   0
            Width           =   6495
         End
      End
      Begin VB.PictureBox picHScroll 
         Height          =   330
         Left            =   -45
         ScaleHeight     =   270
         ScaleWidth      =   7065
         TabIndex        =   6
         Top             =   6435
         Width           =   7125
         Begin VB.HScrollBar HScroll1 
            Height          =   285
            Left            =   0
            TabIndex        =   7
            Top             =   0
            Width           =   6765
         End
      End
      Begin VB.PictureBox picVScroll 
         Height          =   6765
         Left            =   6795
         ScaleHeight     =   6705
         ScaleWidth      =   270
         TabIndex        =   4
         Top             =   -45
         Width           =   330
         Begin VB.VScrollBar VScroll1 
            Height          =   6405
            Left            =   0
            TabIndex        =   5
            Top             =   0
            Width           =   285
         End
      End
   End
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   225
      MultiSelect     =   2  'Extended
      Pattern         =   "*.gif;*.jpg;*.bmp;*.wmf;*.ico;*.cur"
      TabIndex        =   2
      Top             =   6705
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.DirListBox Dir1 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2520
      Left            =   45
      TabIndex        =   1
      Top             =   3375
      Width           =   2280
   End
   Begin VB.DriveListBox Drive1 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   45
      TabIndex        =   0
      Top             =   6030
      Width           =   2280
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   960
      Left            =   45
      TabIndex        =   18
      Top             =   6525
      Width           =   2310
   End
   Begin VB.Label lblCurrent 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   45
      TabIndex        =   17
      Top             =   135
      Width           =   2265
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "SPEED:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1530
      TabIndex        =   14
      Top             =   6570
      Width           =   525
   End
   Begin VB.Label lblSpeed 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1575
      TabIndex        =   13
      Top             =   6795
      Width           =   375
   End
End
Attribute VB_Name = "frmViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdEnd_Click()
    VScroll1.Value = VScroll1.Max
    HScroll1.Value = HScroll1.Max
End Sub

Private Sub cmdSlideshow_Click()
    List2.Clear
    For i = 0 To List1.ListCount - 1
        If List1.Selected(i) Then
            List2.AddItem List1.List(i)
        End If
    Next i
    
    Timer1.Interval = Val(lblSpeed.Caption) * 500
    Timer1.Enabled = True
End Sub

Private Sub Fill_List()
    List1.Clear
    For n = 0 To File1.ListCount
        If Len(File1.List(n)) > 0 Then
            List1.AddItem (File1.List(n))
        End If
    Next
End Sub

Private Sub Dir1_Change()
    File1.Path = Dir1.Path
    Fill_List
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
End Sub

Private Sub Display_Picture(fn As String)
    picMain.Picture = LoadPicture(fn)
End Sub

Private Sub List1_Click()
    fn = Dir1.Path + "\" + List1.List(List1.ListIndex)
    lblCurrent.Caption = List1.List(List1.ListIndex)
    Display_Picture (fn)
    
    VScroll1.Min = 0
    VScroll1.Max = picMain.Height \ 2
    VScroll1.SmallChange = 10
    VScroll1.LargeChange = picMain.Picture.Height / 10
    
    HScroll1.Min = 0
    HScroll1.Max = picMain.Width \ 2
    HScroll1.SmallChange = 10
    HScroll1.LargeChange = picMain.Picture.Width / 10
End Sub

Private Sub Form_Load()
    picSize.BorderStyle = 0
    
    picMain.Left = picSize.Left
    picMain.Top = picSize.Top
    picMain.Width = picSize.ScaleWidth
    picMain.Height = picSize.ScaleHeight
End Sub

Private Sub Form_Resize()
    Me.ScaleMode = 3
    picContainer.ScaleMode = 3
    picVScroll.ScaleMode = 3
    picHScroll.ScaleMode = 3
    picSize.ScaleMode = 3

    picContainer.Width = ScaleWidth - 165
    picContainer.Height = ScaleHeight - 15
    
    picVScroll.Left = ScaleWidth - 189
    picVScroll.Height = ScaleHeight - 34
    VScroll1.Height = picVScroll.ScaleHeight
    
    picHScroll.Top = ScaleHeight - 39
    picHScroll.Width = ScaleWidth - 184
    HScroll1.Width = picHScroll.ScaleWidth
    
    picSize.Width = picVScroll.Left - 5
    picSize.Height = picHScroll.Top - 5
    
    cmdEnd.Left = picVScroll.Left + 1
    cmdEnd.Top = picHScroll.Top + 1
    cmdEnd.Width = 20
    cmdEnd.Height = 20
End Sub

Private Sub HScroll1_Change()
    picMain.Left = 0 - HScroll1.Value
End Sub

Private Sub Timer1_Timer()
    Static i As Integer
    Dim fn As String
    
    fn = Dir1.Path + "\"
    If i < List2.ListCount Then
        lblCurrent.Caption = List2.List(i) + " (" + Format(i + 1) + "/" + Format(List2.ListCount) + ")"
        Display_Picture (fn + List2.List(i))
        i = i + 1
    Else
        i = 0
        If chkLoop.Value = False Then
            Timer1.Enabled = False
        End If
    End If
End Sub

Private Sub VScroll1_Change()
    picMain.Top = 0 - VScroll1.Value
End Sub

Private Sub vsSpeed_Change()
    lblSpeed = Format(vsSpeed.Value)
End Sub
