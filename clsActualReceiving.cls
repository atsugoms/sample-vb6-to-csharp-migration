VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsActualReceiving"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' 入庫実績データ
Public Id As Long
Public OrderNumber As String      ' 発注番号
Public ProductCode As String      ' 商品コード
Public Quantity As Long           ' 数量
Public BarcodeType As String      ' バーコードの種類
Public InspectionError As String  ' 検品時エラー処理
Public ShelfNumber As String      ' 棚番号
Public ReceivingDate As Date      ' 入庫日
