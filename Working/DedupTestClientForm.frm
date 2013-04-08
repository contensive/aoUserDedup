VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Dedup Client"
   ClientHeight    =   3795
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   7995
   LinkTopic       =   "Form1"
   ScaleHeight     =   3795
   ScaleWidth      =   7995
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Begin Dedup"
      Height          =   375
      Left            =   2280
      TabIndex        =   3
      Top             =   1440
      Width           =   2055
   End
   Begin VB.TextBox Appname 
      Height          =   285
      Left            =   2280
      TabIndex        =   1
      Top             =   1080
      Width           =   3495
   End
   Begin VB.Label StatusLabel 
      Caption         =   "status"
      Height          =   735
      Left            =   2280
      TabIndex        =   4
      Top             =   1920
      Width           =   3615
   End
   Begin VB.Label Label2 
      Caption         =   "App Name"
      Height          =   255
      Left            =   960
      TabIndex        =   2
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Dedup people records on email"
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   6855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim ao As New aoUserDeDup.AutoProcessClass
    Dim csv As New ContentServerClass
    Dim CSConn As CSConnectionType
    '
    CSConn = csv.OpenConnection(Appname.Text)
    If CSConn.ApplicationStatus <> ApplicationStatusRunning Then
        StatusLabel.Caption = "Status: Application not running"
    Else
        StatusLabel.Caption = "Status: Runing dedup"
    End If
    Call ao.Main(csv, "")
    Set csv = Nothing
    StatusLabel.Caption = "Status: Finished"
    '
End Sub
