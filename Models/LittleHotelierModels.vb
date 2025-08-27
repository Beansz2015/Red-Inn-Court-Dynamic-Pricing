' LittleHotelier API response models matching your JSON structure
Public Class LittleHotelierResponse
    Public Property name As String
    Public Property rate_plans As List(Of RatePlan)
End Class

Public Class RatePlan
    Public Property id As Integer
    Public Property name As String
    Public Property rate_plan_dates As List(Of RatePlanDate)
End Class

Public Class RatePlanDate
    Public Property id As Integer?
    Public Property [date] As String
    Public Property rate As Decimal
    Public Property min_stay As Integer
    Public Property stop_online_sell As Boolean
    Public Property close_to_arrival As Boolean
    Public Property close_to_departure As Boolean
    Public Property max_stay As Integer?
    Public Property available As Integer
End Class
