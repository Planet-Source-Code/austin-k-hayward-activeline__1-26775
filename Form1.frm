VERSION 5.00
Object = "{7E31D7F3-9577-497B-9232-98085CB0BA83}#1.0#0"; "ActiveLine.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H8000000A&
   Caption         =   "frmMain"
   ClientHeight    =   5340
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9285
   LinkTopic       =   "Form1"
   ScaleHeight     =   5340
   ScaleWidth      =   9285
   StartUpPosition =   3  'Windows Default
   Begin prjLineProject.ActiveLine ActiveLine3 
      Height          =   1290
      Left            =   2820
      TabIndex        =   5
      Top             =   3540
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   2275
      Alignment       =   1
   End
   Begin prjLineProject.ActiveLine ActiveLine2 
      Height          =   150
      Left            =   660
      TabIndex        =   4
      Top             =   3180
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   265
   End
   Begin prjLineProject.ActiveLine ActiveLine1 
      Height          =   150
      Left            =   660
      TabIndex        =   2
      Top             =   1980
      Width           =   7830
      _ExtentX        =   13811
      _ExtentY        =   265
      BackColor       =   -2147483638
   End
   Begin prjLineProject.ActiveLine ActiveLine4 
      Height          =   1290
      Left            =   6300
      TabIndex        =   6
      Top             =   3540
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   2275
      Alignment       =   1
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Now with vertical alignment as well as horizontal!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   3585
      TabIndex        =   3
      Top             =   3840
      Width           =   2115
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000A&
      Caption         =   $"Form1.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1275
      Left            =   705
      TabIndex        =   0
      Top             =   420
      Width           =   7815
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000A&
      Caption         =   $"Form1.frx":01E1
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   675
      Left            =   720
      TabIndex        =   1
      Top             =   2280
      Width           =   7815
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

