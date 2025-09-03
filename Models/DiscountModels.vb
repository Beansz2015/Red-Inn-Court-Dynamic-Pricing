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
