VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Convert doc types"
   ClientHeight    =   2595
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4200
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2595
   ScaleWidth      =   4200
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Convert >"
      Height          =   495
      Left            =   1320
      TabIndex        =   0
      Top             =   1080
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    ConvertFile "c:\test.rtf", "c:\test.doc"
End Sub
