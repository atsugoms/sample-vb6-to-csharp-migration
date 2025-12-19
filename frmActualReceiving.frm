VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmActualReceiving 
   Caption         =   "入庫実績管理"
   ClientHeight    =   7200
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11400
   LinkTopic       =   "Form1"
   ScaleHeight     =   7200
   ScaleWidth      =   11400
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "入庫実績詳細"
      Height          =   2895
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   11175
      Begin VB.TextBox txtOrderNumber 
         Height          =   375
         Left            =   6360
         TabIndex        =   0
         Top             =   240
         Width           =   2415
      End
      Begin VB.TextBox txtProductCode 
         Height          =   375
         Left            =   6360
         TabIndex        =   1
         Top             =   480
         Width           =   2415
      End
      Begin VB.TextBox txtQuantity 
         Height          =   375
         Left            =   1680
         TabIndex        =   2
         Top             =   480
         Width           =   2415
      End
      Begin VB.TextBox txtBarcodeType 
         Height          =   375
         Left            =   1680
         TabIndex        =   3
         Top             =   960
         Width           =   2415
      End
      Begin VB.TextBox txtInspectionError 
         Height          =   375
         Left            =   1680
         TabIndex        =   4
         Top             =   1440
         Width           =   2415
      End
      Begin VB.TextBox txtShelfNumber 
         Height          =   375
         Left            =   1680
         TabIndex        =   5
         Top             =   1920
         Width           =   2415
      End
      Begin VB.TextBox txtReceivingDate 
         Height          =   375
         Left            =   1680
         TabIndex        =   6
         Top             =   2400
         Width           =   2415
      End
      Begin VB.Label Label7 
         Caption         =   "入庫日:"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   2460
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "棚番号:"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   1980
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "検品時エラー:"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   1500
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "バーコード種類:"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   1020
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "数量:"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   540
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "商品コード:"
         Height          =   255
         Left            =   4920
         TabIndex        =   15
         Top             =   540
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "発注番号:"
         Height          =   255
         Left            =   4920
         TabIndex        =   14
         Top             =   300
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "閉じる"
      Height          =   495
      Left            =   9720
      TabIndex        =   13
      Top             =   6480
      Width           =   1455
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "削除"
      Height          =   495
      Left            =   6120
      TabIndex        =   12
      Top             =   3120
      Width           =   1455
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "編集"
      Height          =   495
      Left            =   4440
      TabIndex        =   11
      Top             =   3120
      Width           =   1455
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "追加"
      Height          =   495
      Left            =   2760
      TabIndex        =   10
      Top             =   3120
      Width           =   1455
   End
   Begin MSComctlLib.ListView lstActualReceiving 
      Height          =   2775
      Left            =   120
      TabIndex        =   7
      Top             =   3720
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   4895
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
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "発注番号"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "商品コード"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "数量"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "バーコード"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "エラー処理"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "棚番号"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "入庫日"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label lblSelectedId 
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   3240
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "frmActualReceiving"
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
    Dim ar As clsActualReceiving
    Dim li As ListItem
    
    lstActualReceiving.ListItems.Clear
    
    For Each ar In ActualReceivings
        Set li = lstActualReceiving.ListItems.Add(, , CStr(ar.Id))
        li.SubItems(1) = ar.OrderNumber
        li.SubItems(2) = ar.ProductCode
        li.SubItems(3) = CStr(ar.Quantity)
        li.SubItems(4) = ar.BarcodeType
        li.SubItems(5) = ar.InspectionError
        li.SubItems(6) = ar.ShelfNumber
        li.SubItems(7) = Format(ar.ReceivingDate, "yyyy/mm/dd")
    Next ar
    
    Call ClearForm
End Sub

Private Sub ClearForm()
    txtOrderNumber.Text = ""
    txtProductCode.Text = ""
    txtQuantity.Text = ""
    txtBarcodeType.Text = ""
    txtInspectionError.Text = ""
    txtShelfNumber.Text = ""
    txtReceivingDate.Text = ""
    lblSelectedId.Caption = ""
End Sub

Private Sub cmdAdd_Click()
    Dim ar As clsActualReceiving
    
    If Trim(txtOrderNumber.Text) = "" Then
        MsgBox "発注番号を入力してください。", vbExclamation
        Exit Sub
    End If
    
    Set ar = New clsActualReceiving
    ar.OrderNumber = txtOrderNumber.Text
    ar.ProductCode = txtProductCode.Text
    ar.Quantity = Val(txtQuantity.Text)
    ar.BarcodeType = txtBarcodeType.Text
    ar.InspectionError = txtInspectionError.Text
    ar.ShelfNumber = txtShelfNumber.Text
    If IsDate(txtReceivingDate.Text) Then
        ar.ReceivingDate = CDate(txtReceivingDate.Text)
    Else
        ar.ReceivingDate = Date
    End If
    
    Call AddActualReceiving(ar)
    
    MsgBox "追加しました。", vbInformation
    Call LoadList
End Sub

Private Sub cmdEdit_Click()
    Dim ar As clsActualReceiving
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
    Set ar = GetActualReceiving(Id)
    
    ar.OrderNumber = txtOrderNumber.Text
    ar.ProductCode = txtProductCode.Text
    ar.Quantity = Val(txtQuantity.Text)
    ar.BarcodeType = txtBarcodeType.Text
    ar.InspectionError = txtInspectionError.Text
    ar.ShelfNumber = txtShelfNumber.Text
    If IsDate(txtReceivingDate.Text) Then
        ar.ReceivingDate = CDate(txtReceivingDate.Text)
    Else
        ar.ReceivingDate = Date
    End If
    
    Call UpdateActualReceiving(ar)
    
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
    Call DeleteActualReceiving(Id)
    
    MsgBox "削除しました。", vbInformation
    Call LoadList
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub lstActualReceiving_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim ar As clsActualReceiving
    Dim Id As Long
    
    Id = CLng(Item.Text)
    lblSelectedId.Caption = CStr(Id)
    
    Set ar = GetActualReceiving(Id)
    txtOrderNumber.Text = ar.OrderNumber
    txtProductCode.Text = ar.ProductCode
    txtQuantity.Text = CStr(ar.Quantity)
    txtBarcodeType.Text = ar.BarcodeType
    txtInspectionError.Text = ar.InspectionError
    txtShelfNumber.Text = ar.ShelfNumber
    txtReceivingDate.Text = Format(ar.ReceivingDate, "yyyy/mm/dd")
End Sub
