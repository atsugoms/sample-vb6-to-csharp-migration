Attribute VB_Name = "modDatabase"
Option Explicit

' In-memory database collections
Public ScheduledReceivings As Collection
Public ActualReceivings As Collection
Public Locations As Collection

' ID generators
Private ScheduledReceivingNextId As Long
Private ActualReceivingNextId As Long
Private LocationNextId As Long

' Initialize database
Public Sub InitializeDatabase()
    Set ScheduledReceivings = New Collection
    Set ActualReceivings = New Collection
    Set Locations = New Collection
    
    ScheduledReceivingNextId = 1
    ActualReceivingNextId = 1
    LocationNextId = 1
    
    ' サンプルデータの追加
    Call AddSampleData
End Sub

' サンプルデータの追加
Private Sub AddSampleData()
    Dim sr As clsScheduledReceiving
    Dim ar As clsActualReceiving
    Dim loc As clsLocation
    
    ' 入庫予定サンプル
    Set sr = New clsScheduledReceiving
    sr.Id = GetNextScheduledReceivingId()
    sr.OrderNumber = "PO-001"
    sr.ProductCode = "PROD-001"
    sr.Quantity = 100
    sr.DeliveryDate = DateAdd("d", 3, Date)
    ScheduledReceivings.Add sr, CStr(sr.Id)
    
    Set sr = New clsScheduledReceiving
    sr.Id = GetNextScheduledReceivingId()
    sr.OrderNumber = "PO-002"
    sr.ProductCode = "PROD-002"
    sr.Quantity = 50
    sr.DeliveryDate = DateAdd("d", 5, Date)
    ScheduledReceivings.Add sr, CStr(sr.Id)
    
    ' ロケーションサンプル
    Set loc = New clsLocation
    loc.Id = GetNextLocationId()
    loc.ShelfNumber = "A-01"
    loc.ShelfAttribute = "常温"
    loc.AssignmentRule = "先入先出"
    Locations.Add loc, CStr(loc.Id)
    
    Set loc = New clsLocation
    loc.Id = GetNextLocationId()
    loc.ShelfNumber = "A-02"
    loc.ShelfAttribute = "常温"
    loc.AssignmentRule = "先入先出"
    Locations.Add loc, CStr(loc.Id)
    
    Set loc = New clsLocation
    loc.Id = GetNextLocationId()
    loc.ShelfNumber = "B-01"
    loc.ShelfAttribute = "冷蔵"
    loc.AssignmentRule = "後入先出"
    Locations.Add loc, CStr(loc.Id)
End Sub

' ID generators
Public Function GetNextScheduledReceivingId() As Long
    GetNextScheduledReceivingId = ScheduledReceivingNextId
    ScheduledReceivingNextId = ScheduledReceivingNextId + 1
End Function

Public Function GetNextActualReceivingId() As Long
    GetNextActualReceivingId = ActualReceivingNextId
    ActualReceivingNextId = ActualReceivingNextId + 1
End Function

Public Function GetNextLocationId() As Long
    GetNextLocationId = LocationNextId
    LocationNextId = LocationNextId + 1
End Function

' CRUD operations for Scheduled Receiving
Public Sub AddScheduledReceiving(sr As clsScheduledReceiving)
    sr.Id = GetNextScheduledReceivingId()
    ScheduledReceivings.Add sr, CStr(sr.Id)
End Sub

Public Sub UpdateScheduledReceiving(sr As clsScheduledReceiving)
    ScheduledReceivings.Remove CStr(sr.Id)
    ScheduledReceivings.Add sr, CStr(sr.Id)
End Sub

Public Sub DeleteScheduledReceiving(Id As Long)
    ScheduledReceivings.Remove CStr(Id)
End Sub

Public Function GetScheduledReceiving(Id As Long) As clsScheduledReceiving
    Set GetScheduledReceiving = ScheduledReceivings(CStr(Id))
End Function

' CRUD operations for Actual Receiving
Public Sub AddActualReceiving(ar As clsActualReceiving)
    ar.Id = GetNextActualReceivingId()
    ActualReceivings.Add ar, CStr(ar.Id)
End Sub

Public Sub UpdateActualReceiving(ar As clsActualReceiving)
    ActualReceivings.Remove CStr(ar.Id)
    ActualReceivings.Add ar, CStr(ar.Id)
End Sub

Public Sub DeleteActualReceiving(Id As Long)
    ActualReceivings.Remove CStr(Id)
End Sub

Public Function GetActualReceiving(Id As Long) As clsActualReceiving
    Set GetActualReceiving = ActualReceivings(CStr(Id))
End Function

' CRUD operations for Location
Public Sub AddLocation(loc As clsLocation)
    loc.Id = GetNextLocationId()
    Locations.Add loc, CStr(loc.Id)
End Sub

Public Sub UpdateLocation(loc As clsLocation)
    Locations.Remove CStr(loc.Id)
    Locations.Add loc, CStr(loc.Id)
End Sub

Public Sub DeleteLocation(Id As Long)
    Locations.Remove CStr(Id)
End Sub

Public Function GetLocation(Id As Long) As clsLocation
    Set GetLocation = Locations(CStr(Id))
End Function
