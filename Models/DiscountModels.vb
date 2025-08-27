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
