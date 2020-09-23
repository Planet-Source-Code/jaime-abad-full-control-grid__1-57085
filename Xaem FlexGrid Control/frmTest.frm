VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmTest 
   Caption         =   "Test"
   ClientHeight    =   6615
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   8880
   LinkTopic       =   "Form1"
   ScaleHeight     =   6615
   ScaleWidth      =   8880
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox pcbText 
      Align           =   1  'Align Top
      Height          =   540
      Left            =   0
      ScaleHeight     =   480
      ScaleWidth      =   8820
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   8880
      Begin VB.Label Label1 
         Caption         =   "Use right  click to use the menu (Copy) (Cut) (Paste) (Insert) (Clear Text)"
         Height          =   315
         Index           =   1
         Left            =   75
         TabIndex        =   5
         Top             =   225
         Width           =   9015
      End
      Begin VB.Label Label1 
         Caption         =   "Try move the cols or edit the cells, you can use clipboard to export o import to excel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   75
         TabIndex        =   4
         Top             =   0
         Width           =   9015
      End
   End
   Begin VB.PictureBox pcbIcon 
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   7725
      Picture         =   "frmTest.frx":0000
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   2
      Top             =   1275
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.TextBox txtEditar 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   3300
      MultiLine       =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   4650
      Visible         =   0   'False
      Width           =   915
   End
   Begin MSFlexGridLib.MSFlexGrid msfTest 
      Height          =   5265
      Left            =   150
      TabIndex        =   0
      Top             =   975
      Width           =   6765
      _ExtentX        =   11933
      _ExtentY        =   9287
      _Version        =   393216
      Rows            =   12
      Cols            =   12
      WordWrap        =   -1  'True
      AllowUserResizing=   3
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      Begin VB.Menu mnuCut 
         Caption         =   "Cut"
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "Copy"
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "Paste"
      End
      Begin VB.Menu mnuSeparator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInsert 
         Caption         =   "Insert"
      End
      Begin VB.Menu mnuClear 
         Caption         =   "Clear Cell"
      End
      Begin VB.Menu mnusep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSelectAll 
         Caption         =   "Select All"
      End
   End
   Begin VB.Menu mnuOption 
      Caption         =   "Option"
      Begin VB.Menu mnuAEdit 
         Caption         =   "Allow Edit"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuACols 
         Caption         =   "Allow Move Cols"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuADel 
         Caption         =   "Allow Key Del"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuAEsc 
         Caption         =   "Allow Key Esc"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuSeparator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAUp 
         Caption         =   "Allow Key Up"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuADown 
         Caption         =   "Allow Key Down"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuALeft 
         Caption         =   "Allow Key Left"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuARight 
         Caption         =   "Allow Key Right"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuATab 
         Caption         =   "Allow Key Tab"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuAEnter 
         Caption         =   "Allow Key Enter"
         Checked         =   -1  'True
      End
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim msfeDATA As xssEditFlexGrid

Private Sub Form_Load()
    Dim i As Long, x As Integer, y As Integer
    Set msfeDATA = New xssEditFlexGrid
    msfeDATA.SetFlexGrid msfTest
    msfeDATA.SetTextBox txtEditar
    With msfTest
        For y = 1 To 10
            .Row = y
            For x = 1 To 10
                i = i + 1
                .Col = x
                msfTest.Text = i
            Next
        Next
    End With
End Sub

Private Sub Form_Resize()
    msfTest.Move Me.ScaleLeft, Me.ScaleTop + pcbText.Height, Me.ScaleWidth, Me.ScaleHeight
End Sub

Private Sub mnuACols_Click()
    mnuACols.Checked = Not mnuACols.Checked
    msfeDATA.AllowMoveCols = mnuACols.Checked
End Sub

Private Sub mnuADel_Click()
    mnuADel.Checked = Not mnuADel.Checked
    msfeDATA.AllowKeyDel = mnuADel.Checked
End Sub

Private Sub mnuADown_Click()
    mnuADown.Checked = Not mnuADown.Checked
    msfeDATA.AllowKeyDown = mnuADown.Checked
End Sub

Private Sub mnuAEdit_Click()
    mnuAEdit.Checked = Not mnuAEdit.Checked
    msfeDATA.AllowEditCells = mnuAEdit.Checked
End Sub

Private Sub mnuAEnter_Click()
    mnuAEnter.Checked = Not mnuAEnter.Checked
    msfeDATA.AllowKeyEnter = mnuAEnter.Checked
End Sub

Private Sub mnuAEsc_Click()
    mnuAEsc.Checked = Not mnuAEsc.Checked
    msfeDATA.AllowKeyEsc = mnuAEsc.Checked
End Sub

Private Sub mnuALeft_Click()
    mnuALeft.Checked = Not mnuALeft.Checked
    msfeDATA.AllowKeyLeft = mnuALeft.Checked
End Sub

Private Sub mnuARight_Click()
    mnuARight.Checked = Not mnuARight.Checked
    msfeDATA.AllowKeyRight = mnuARight.Checked
End Sub

Private Sub mnuATab_Click()
    mnuATab.Checked = Not mnuATab.Checked
    msfeDATA.AllowKeyTab = mnuATab.Checked
End Sub

Private Sub mnuAUp_Click()
    mnuAUp.Checked = Not mnuAUp.Checked
    msfeDATA.AllowKeyUp = mnuAUp.Checked
End Sub

Private Sub mnuClear_Click()
    msfeDATA.ClearCells
End Sub

Private Sub mnuCopy_Click()
    msfeDATA.Copy
End Sub

Private Sub mnuCut_Click()
    msfeDATA.Copy
    msfeDATA.ClearCells
End Sub

Private Sub mnuInsert_Click()
    frmSlide.Show vbModal
    If InsertMethod <> -1 Then
        msfeDATA.Insert (InsertMethod)
    End If
End Sub

Private Sub mnuPaste_Click()
    msfeDATA.Paste
End Sub

Private Sub mnuSelectAll_Click()
    msfeDATA.SelectAll
End Sub

Private Sub msfTest_DblClick()
    msfeDATA.StarEdit
End Sub

Private Sub msfTest_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
        If Not Clipboard.GetFormat(vbCFText) Then
            mnuPaste.Enabled = False
            mnuInsert.Enabled = False
        Else
            mnuPaste.Enabled = True
            mnuInsert.Enabled = True
        End If
            PopupMenu mnuEdit
    Else
        msfTest.DragIcon = pcbIcon
    End If
End Sub
