' Discount calculation and tracking models
Public Class DiscountChange
    Public Property CheckDate As String
    Public Property RoomType As String
    Public Property OldDiscount As Double
    Public Property NewDiscount As Double
    Public Property OccupancyPct As Double
    Public Property AvailableUnits As Integer
    Public Property DaysAhead As Integer
End Class

Public Class PreviousDiscount
    Public Property DormDiscount As Double
    Public Property PrivateDiscount As Double
    Public Property EnsuiteDiscount As Double
    Public Property TwinDiscount As Double
    Public Property LastUpdated As DateTime
End Class
' Rate change tracking models
Public Class RateChange
    Public Property CheckDate As String
    Public Property RoomType As String
    Public Property OldRate As Double
    Public Property NewRate As Double
    Public Property OccupancyPct As Double
    Public Property AvailableUnits As Integer
    Public Property DaysAhead As Integer
End Class

Public Class PreviousRate
    Public Property DormRate As Double
    Public Property PrivateRate As Double
    Public Property EnsuiteRate As Double
    Public Property TwinRate As Double
    Public Property LastUpdated As DateTime
End Class
