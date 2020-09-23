VERSION 5.00
Begin VB.Form frmExtend 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   750
      Style           =   1  'Simple Combo
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   1425
      Width           =   1440
   End
   Begin VB.ListBox List1 
      Height          =   255
      Left            =   675
      TabIndex        =   1
      Top             =   1050
      Width           =   1515
   End
   Begin VB.TextBox txtEditar 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   915
   End
End
Attribute VB_Name = "frmExtend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

