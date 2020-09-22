VERSION 5.00
Begin VB.Form frmMsg 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Private Message"
   ClientHeight    =   2760
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Send"
      Height          =   495
      Left            =   3960
      TabIndex        =   2
      Top             =   2160
      Width           =   615
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   2160
      Width           =   3855
   End
   Begin VB.TextBox Text1 
      Height          =   2055
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "frmMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
