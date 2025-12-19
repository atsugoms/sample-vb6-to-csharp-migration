VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLocation 
   Caption         =   "ロケーション管理"
   ClientHeight    =   6600
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10800
   LinkTopic       =   "Form1"
   ScaleHeight     =   6600
   ScaleWidth      =   10800
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "ロケーション詳細"
      Height          =   1815
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   10575
      Begin VB.TextBox txtAssignmentRule 
         Height          =   375
         Left            =   1680
         TabIndex        =   2
         Top             =   1080
         Width           =   2415
      End
      Begin VB.TextBox txtShelfAttribute 
         Height          =   375
         Left            =   1680
         TabIndex        =   1
         Top             =   600
         Width           =   2415
      End
      Begin VB.TextBox txtShelfNumber 
         Height          =   375
         Left            =   1680
         TabIndex        =   0
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label Label3 
         Caption         =   "割り当てルール:"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   1140
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "棚属性:"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   660
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "棚番号:"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   300
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "閉じる"
      Height          =   495
      Left            =   9120
      TabIndex        =   10
      Top             =   5880
      Width           =   1455
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "削除"
      Height          =   495
      Left            =   6120
      TabIndex        =   9
      Top             =   2040
      Width           =   1455
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "編集"
      Height          =   495
      Left            =   4440
      TabIndex        =   8
      Top             =   2040
      Width           =   1455
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "追加"
      Height          =   495
      Left            =   2760
      TabIndex        =   7
      Top             =   2040
      Width           =   1455
   End
   Begin MSComctlLib.ListView lstLocation 
      Height          =   3255
      Left            =   120
      TabIndex        =   3
      Top             =   2640
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   5741
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "棚番号"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "棚属性"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "割り当てルール"
         Object.Width           =   5292
      EndProperty
   End
   Begin VB.Label lblSelectedId 
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   2160
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "frmLocation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Call LoadList
End Sub

Private Sub Form_Activate()
    Call LoadList
End Sub

Private Sub LoadList()
    Dim loc As clsLocation
    Dim li As ListItem
    
    lstLocation.ListItems.Clear
    
    For Each loc In Locations
        Set li = lstLocation.ListItems.Add(, , CStr(loc.Id))
        li.SubItems(1) = loc.ShelfNumber
        li.SubItems(2) = loc.ShelfAttribute
        li.SubItems(3) = loc.AssignmentRule
    Next loc
    
    Call ClearForm
End Sub

Private Sub ClearForm()
    txtShelfNumber.Text = ""
    txtShelfAttribute.Text = ""
    txtAssignmentRule.Text = ""
    lblSelectedId.Caption = ""
End Sub

Private Sub cmdAdd_Click()
    Dim loc As clsLocation
    
    If Trim(txtShelfNumber.Text) = "" Then
        MsgBox "棚番号を入力してください。", vbExclamation
        Exit Sub
    End If
    
    Set loc = New clsLocation
    loc.ShelfNumber = txtShelfNumber.Text
    loc.ShelfAttribute = txtShelfAttribute.Text
    loc.AssignmentRule = txtAssignmentRule.Text
    
    Call AddLocation(loc)
    
    MsgBox "追加しました。", vbInformation
    Call LoadList
End Sub

Private Sub cmdEdit_Click()
    Dim loc As clsLocation
    Dim Id As Long
    
    If lblSelectedId.Caption = "" Then
        MsgBox "編集する項目を選択してください。", vbExclamation
        Exit Sub
    End If
    
    If Trim(txtShelfNumber.Text) = "" Then
        MsgBox "棚番号を入力してください。", vbExclamation
        Exit Sub
    End If
    
    Id = CLng(lblSelectedId.Caption)
    Set loc = GetLocation(Id)
    
    loc.ShelfNumber = txtShelfNumber.Text
    loc.ShelfAttribute = txtShelfAttribute.Text
    loc.AssignmentRule = txtAssignmentRule.Text
    
    Call UpdateLocation(loc)
    
    MsgBox "更新しました。", vbInformation
    Call LoadList
End Sub

Private Sub cmdDelete_Click()
    Dim Id As Long
    
    If lblSelectedId.Caption = "" Then
        MsgBox "削除する項目を選択してください。", vbExclamation
        Exit Sub
    End If
    
    If MsgBox("削除してもよろしいですか？", vbQuestion + vbYesNo) = vbNo Then
        Exit Sub
    End If
    
    Id = CLng(lblSelectedId.Caption)
    Call DeleteLocation(Id)
    
    MsgBox "削除しました。", vbInformation
    Call LoadList
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub lstLocation_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim loc As clsLocation
    Dim Id As Long
    
    Id = CLng(Item.Text)
    lblSelectedId.Caption = CStr(Id)
    
    Set loc = GetLocation(Id)
    txtShelfNumber.Text = loc.ShelfNumber
    txtShelfAttribute.Text = loc.ShelfAttribute
    txtAssignmentRule.Text = loc.AssignmentRule
End Sub
