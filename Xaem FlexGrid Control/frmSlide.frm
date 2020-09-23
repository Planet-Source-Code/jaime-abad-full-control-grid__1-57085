VERSION 5.00
Begin VB.Form frmSlide 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Select insert method"
   ClientHeight    =   1485
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3660
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1485
   ScaleWidth      =   3660
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   2400
      TabIndex        =   4
      Top             =   975
      Width           =   1140
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   315
      Left            =   2400
      TabIndex        =   3
      Top             =   600
      Width           =   1140
   End
   Begin VB.OptionButton OptDown 
      Caption         =   "Slide cells to the down"
      Height          =   240
      Left            =   150
      TabIndex        =   1
      Top             =   975
      Width           =   1965
   End
   Begin VB.OptionButton optRight 
      Caption         =   "Slide cells to the right"
      Height          =   240
      Left            =   150
      TabIndex        =   0
      Top             =   600
      Value           =   -1  'True
      Width           =   1965
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      Index           =   1
      X1              =   750
      X2              =   3525
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      Index           =   0
      X1              =   750
      X2              =   3525
      Y1              =   225
      Y2              =   225
   End
   Begin VB.Label lblInsert 
      Caption         =   "Insert"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   150
      TabIndex        =   2
      Top             =   150
      Width           =   615
   End
End
Attribute VB_Name = "frmSlide"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOk_Click()
    InsertMethod = IIf(optRight.Value, 0, 1)
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    InsertMethod = -1
    Unload Me
End Sub
