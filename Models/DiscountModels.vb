' Rate change tracking models
Public Class RateChange
    Public Property CheckDate As String
    Public Property RoomType As String
    Public Property OldRegularRate As Double
    Public Property NewRegularRate As Double
    Public Property OldWalkInRate As Double
    Public Property NewWalkInRate As Double
    Public Property AvailableUnits As Integer
    Public Property DaysAhead As Integer
End Class

Public Class PreviousRate
    Public Property DormRegularRate As Double
    Public Property DormWalkInRate As Double
    Public Property PrivateRegularRate As Double
    Public Property PrivateWalkInRate As Double
    Public Property EnsuiteRegularRate As Double
    Public Property EnsuiteWalkInRate As Double
    Public Property LastUpdated As DateTime
End Class

' NEW: QuietPeriod model class
Public Class QuietPeriod
    Public Property Name As String
    Public Property DayOfWeek As DayOfWeek
    Public Property StartTime As TimeSpan
    Public Property EndTime As TimeSpan
    Public Property Enabled As Boolean
End Class

' NEW: Price History Analytics Entry
Public Class PriceHistoryEntry
    Public Property Timestamp As DateTime
    Public Property CheckInDate As String
    Public Property RoomType As String
    Public Property RoomCategory As String
    Public Property AvailableUnits As Integer
    Public Property TotalCapacity As Integer
    Public Property OccupancyRate As Double
    Public Property DaysAhead As Integer
    Public Property OldRate As Double
    Public Property NewRate As Double
    Public Property RateChange As Double
    Public Property PercentChange As Double
    Public Property PriceTier As String
    Public Property DayType As String
    Public Property BusinessHour As Integer
    Public Property ConfigSource As String
End Class
