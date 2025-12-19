VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmScheduledReceiving 
   Caption         =   "入庫予定管理"
   ClientHeight    =   6600
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10800
   LinkTopic       =   "Form1"
   ScaleHeight     =   6600
   ScaleWidth      =   10800
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "入庫予定詳細"
      Height          =   2295
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   10575
      Begin VB.TextBox txtDeliveryDate 
         Height          =   375
         Left            =   1680
         TabIndex        =   3
         Top             =   1560
         Width           =   2415
      End
      Begin VB.TextBox txtQuantity 
         Height          =   375
         Left            =   1680
         TabIndex        =   2
         Top             =   1080
         Width           =   2415
      End
      Begin VB.TextBox txtProductCode 
         Height          =   375
         Left            =   1680
         TabIndex        =   1
         Top             =   600
         Width           =   2415
      End
      Begin VB.TextBox txtOrderNumber 
         Height          =   375
         Left            =   1680
         TabIndex        =   0
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label Label4 
         Caption         =   "納期:"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   1620
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "数量:"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   1140
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "商品コード:"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   660
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "発注番号:"
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
      Top             =   2520
      Width           =   1455
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "編集"
      Height          =   495
      Left            =   4440
      TabIndex        =   8
      Top             =   2520
      Width           =   1455
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "追加"
      Height          =   495
      Left            =   2760
      TabIndex        =   7
      Top             =   2520
      Width           =   1455
   End
   Begin MSComctlLib.ListView lstScheduledReceiving 
      Height          =   3135
      Left            =   120
      TabIndex        =   4
      Top             =   3240
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   5530
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
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "発注番号"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "商品コード"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "数量"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "納期"
         Object.Width           =   3528
      EndProperty
   End
   Begin VB.Label lblSelectedId 
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   2640
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "frmScheduledReceiving"
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
    Dim sr As clsScheduledReceiving
    Dim li As ListItem
    
    lstScheduledReceiving.ListItems.Clear
    
    For Each sr In ScheduledReceivings
        Set li = lstScheduledReceiving.ListItems.Add(, , CStr(sr.Id))
        li.SubItems(1) = sr.OrderNumber
        li.SubItems(2) = sr.ProductCode
        li.SubItems(3) = CStr(sr.Quantity)
        li.SubItems(4) = Format(sr.DeliveryDate, "yyyy/mm/dd")
    Next sr
    
    Call ClearForm
End Sub

Private Sub ClearForm()
    txtOrderNumber.Text = ""
    txtProductCode.Text = ""
    txtQuantity.Text = ""
    txtDeliveryDate.Text = ""
    lblSelectedId.Caption = ""
End Sub

Private Sub cmdAdd_Click()
    Dim sr As clsScheduledReceiving
    
    If Trim(txtOrderNumber.Text) = "" Then
        MsgBox "発注番号を入力してください。", vbExclamation
        Exit Sub
    End If
    
    Set sr = New clsScheduledReceiving
    sr.OrderNumber = txtOrderNumber.Text
    sr.ProductCode = txtProductCode.Text
    sr.Quantity = Val(txtQuantity.Text)
    If IsDate(txtDeliveryDate.Text) Then
        sr.DeliveryDate = CDate(txtDeliveryDate.Text)
    Else
        sr.DeliveryDate = Date
    End If
    
    Call AddScheduledReceiving(sr)
    
    MsgBox "追加しました。", vbInformation
    Call LoadList
End Sub

Private Sub cmdEdit_Click()
    Dim sr As clsScheduledReceiving
    Dim Id As Long
    
    If lblSelectedId.Caption = "" Then
        MsgBox "編集する項目を選択してください。", vbExclamation
        Exit Sub
    End If
    
    If Trim(txtOrderNumber.Text) = "" Then
        MsgBox "発注番号を入力してください。", vbExclamation
        Exit Sub
    End If
    
    Id = CLng(lblSelectedId.Caption)
    Set sr = GetScheduledReceiving(Id)
    
    sr.OrderNumber = txtOrderNumber.Text
    sr.ProductCode = txtProductCode.Text
    sr.Quantity = Val(txtQuantity.Text)
    If IsDate(txtDeliveryDate.Text) Then
        sr.DeliveryDate = CDate(txtDeliveryDate.Text)
    Else
        sr.DeliveryDate = Date
    End If
    
    Call UpdateScheduledReceiving(sr)
    
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
    Call DeleteScheduledReceiving(Id)
    
    MsgBox "削除しました。", vbInformation
    Call LoadList
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub lstScheduledReceiving_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim sr As clsScheduledReceiving
    Dim Id As Long
    
    Id = CLng(Item.Text)
    lblSelectedId.Caption = CStr(Id)
    
    Set sr = GetScheduledReceiving(Id)
    txtOrderNumber.Text = sr.OrderNumber
    txtProductCode.Text = sr.ProductCode
    txtQuantity.Text = CStr(sr.Quantity)
    txtDeliveryDate.Text = Format(sr.DeliveryDate, "yyyy/mm/dd")
End Sub
