VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "倉庫管理システム - メインメニュー"
   ClientHeight    =   4800
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4800
   ScaleWidth      =   6000
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdLocation 
      Caption         =   "ロケーション管理"
      Height          =   735
      Left            =   1560
      TabIndex        =   3
      Top             =   3120
      Width           =   2895
   End
   Begin VB.CommandButton cmdActualReceiving 
      Caption         =   "入庫実績"
      Height          =   735
      Left            =   1560
      TabIndex        =   2
      Top             =   2160
      Width           =   2895
   End
   Begin VB.CommandButton cmdScheduledReceiving 
      Caption         =   "入庫予定"
      Height          =   735
      Left            =   1560
      TabIndex        =   1
      Top             =   1200
      Width           =   2895
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Caption         =   "倉庫管理システム"
      BeginProperty Font 
         Name            =   "MS ゴシック"
         Size            =   18
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   5055
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    ' データベース初期化
    Call InitializeDatabase
End Sub

Private Sub cmdScheduledReceiving_Click()
    frmScheduledReceiving.Show
End Sub

Private Sub cmdActualReceiving_Click()
    frmActualReceiving.Show
End Sub

Private Sub cmdLocation_Click()
    frmLocation.Show
End Sub
