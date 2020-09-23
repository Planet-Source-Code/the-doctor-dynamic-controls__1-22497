VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1560
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7680
   LinkTopic       =   "Form1"
   ScaleHeight     =   1560
   ScaleWidth      =   7680
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   855
      Left            =   1320
      TabIndex        =   2
      Top             =   600
      Width           =   6255
   End
   Begin VB.CommandButton cmdLoadInFrame 
      Caption         =   "Create 2"
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Create 1"
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents txtTextbox1 As TextBox
Attribute txtTextbox1.VB_VarHelpID = -1
Private WithEvents txtTextbox2 As TextBox
Attribute txtTextbox2.VB_VarHelpID = -1

Private Sub cmdLoad_Click()
    Set txtTextbox1 = Form1.Controls.Add("VB.Textbox", "txtOne")
    With txtTextbox1
        .Visible = True
        .Text = "Dynamisch textbox 1"
        .Height = 315
        .Width = 5000
        .Left = 1440
        .Top = 120
    End With
    cmdLoad.Enabled = False
End Sub

Private Sub cmdLoadInFrame_Click()
    Set txtTextbox2 = Form1.Controls.Add("VB.Textbox", "txtTwo", Frame1)
    With txtTextbox2
        .Visible = True
        .Text = "Dynamisch textbox 2 "
        .Height = 315
        .Width = 5000
        .Left = 120
        .Top = 240
    End With
    cmdLoadInFrame.Enabled = False
End Sub

Private Sub txtTextbox1_DblClick()
    MsgBox txtTextbox2.Container.Name, , "Container"
End Sub

Private Sub txtTextbox2_DblClick()
    MsgBox txtTextbox2.Container.Name, , "Container"
End Sub
