VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4575
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5745
   LinkTopic       =   "Form1"
   ScaleHeight     =   4575
   ScaleWidth      =   5745
   StartUpPosition =   3  'Windows Default
   Begin Project1.Duncan_TitleButton Duncan_TitleButton1 
      Height          =   330
      Left            =   1560
      TabIndex        =   0
      Top             =   360
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   582
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Duncan_TitleButton1_Click()
    Debug.Print "Clicked " & Now
End Sub

