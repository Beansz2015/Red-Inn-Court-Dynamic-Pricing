Imports System.Net.Http
Imports System.Threading.Tasks
Imports Newtonsoft.Json
Imports System.Text
Imports System.IO
Imports System.Configuration
Imports System.Net.Mail
Imports System.Net

Public Class DynamicPricingService
    Private ReadOnly httpClient As New HttpClient()

    ' Configuration settings
    Private ReadOnly apiKey As String = ConfigurationManager.AppSettings("LittleHotelierApiKey")
    Private ReadOnly propertyId As String = ConfigurationManager.AppSettings("PropertyId")
    Private ReadOnly apiBaseUrl As String = ConfigurationManager.AppSettings("ApiBaseUrl")

    ' Email Settings
    Private ReadOnly smtpHost As String = ConfigurationManager.AppSettings("SmtpHost")
    Private ReadOnly smtpPort As Integer = CInt(ConfigurationManager.AppSettings("SmtpPort"))
    Private ReadOnly smtpUsername As String = ConfigurationManager.AppSettings("SmtpUsername")
    Private ReadOnly smtpPassword As String = ConfigurationManager.AppSettings("SmtpPassword")
    Private ReadOnly emailFromAddress As String = ConfigurationManager.AppSettings("EmailFromAddress")
    Private ReadOnly emailFromName As String = ConfigurationManager.AppSettings("EmailFromName")
    Private ReadOnly emailToAddress As String = ConfigurationManager.AppSettings("EmailToAddress")

    ' Room capacity configuration with temporary closures
    Private ReadOnly baseDormBeds As Integer = 20
    Private ReadOnly basePrivateRooms As Integer = 3
    Private ReadOnly baseEnsuiteRooms As Integer = 2

    ' Calculate effective capacity minus temporarily closed units
    Private ReadOnly totalDormBeds As Integer = baseDormBeds - CInt(If(ConfigurationManager.AppSettings("TemporaryClosedDormBeds"), "0"))
    Private ReadOnly totalPrivateRooms As Integer = basePrivateRooms - CInt(If(ConfigurationManager.AppSettings("TemporaryClosedPrivateRooms"), "0"))
    Private ReadOnly totalEnsuiteRooms As Integer = baseEnsuiteRooms - CInt(If(ConfigurationManager.AppSettings("TemporaryClosedEnsuiteRooms"), "0"))


    ' Business timezone helper
    Private Function GetBusinessDateTime() As DateTime
        Dim businessTimeZone = TimeZoneInfo.FindSystemTimeZoneById("Singapore Standard Time")
        Return TimeZoneInfo.ConvertTime(DateTime.UtcNow, businessTimeZone)
    End Function

    Private Function GetBusinessToday() As DateTime
        Return GetBusinessDateTime().Date
    End Function


    Public Async Function RunDynamicPricingCheck() As Task
        Dim errorToReport As String = Nothing
        Dim propertyData As LittleHotelierResponse = Nothing

        Try
            Console.WriteLine($"Starting dynamic pricing check at {DateTime.Now}")

            ' Get availability and store the API response
            Dim availabilityResult = Await GetRoomAvailabilityWithResponseAsync()

            If availabilityResult.AvailabilityData.Count = 0 Then
                Console.WriteLine("No availability data retrieved. Exiting.")
                Return
            End If

            ' Calculate new rates and detect changes using API data
            Dim rateChanges = CalculateRateChanges(availabilityResult.AvailabilityData, availabilityResult.PropertyData)

            ' Send email notifications if there are changes
            If rateChanges.Any() Then
                Await SendEmailNotificationAsync(rateChanges)
                Console.WriteLine($"Sent email notification for {rateChanges.Count} rate changes")
            Else
                Console.WriteLine("No rate changes detected")
            End If

            Console.WriteLine("Dynamic pricing check completed successfully")

        Catch ex As Exception
            Console.WriteLine($"Error in dynamic pricing check: {ex.Message}")
            Console.WriteLine($"Stack trace: {ex.StackTrace}")
            errorToReport = ex.Message
        End Try

        ' Send error notification if there was an error
        If errorToReport IsNot Nothing Then
            Try
                Await SendEmailErrorNotificationAsync(errorToReport)
            Catch notificationEx As Exception
                Console.WriteLine($"Failed to send email error notification: {notificationEx.Message}")
            End Try
        End If
    End Function

    Public Async Function GetRoomAvailabilityAsync() As Task(Of Dictionary(Of String, RoomAvailability))
        Dim availabilityData As New Dictionary(Of String, RoomAvailability)

        Try
            Dim startDate = GetBusinessDateTime().ToString("yyyy-MM-dd")
            Dim url = $"{apiBaseUrl}properties/{propertyId}/rates.json?start_date={startDate}"

            Console.WriteLine($"API URL: {url}")

            httpClient.DefaultRequestHeaders.Clear()
            httpClient.DefaultRequestHeaders.Add("Accept", "application/json")
            httpClient.DefaultRequestHeaders.Add("User-Agent", "RedInnDynamicPricing/1.0")

            Dim response = Await httpClient.GetAsync(url)

            If response.IsSuccessStatusCode Then
                Dim jsonContent = Await response.Content.ReadAsStringAsync()
                Console.WriteLine("Successfully retrieved API data")
                Console.WriteLine($"Response length: {jsonContent.Length} characters")

                Dim apiDataArray = JsonConvert.DeserializeObject(Of List(Of LittleHotelierResponse))(jsonContent)

                If apiDataArray.Count > 0 Then
                    Dim propertyData = apiDataArray(0)
                    availabilityData = ParseAvailabilityByDate(propertyData)
                    Console.WriteLine($"Parsed availability for {availabilityData.Count} dates")
                End If
            Else
                Dim errorContent = Await response.Content.ReadAsStringAsync()
                Console.WriteLine($"API Error: {response.StatusCode} - {errorContent}")
                Console.WriteLine($"Request URL: {url}")
            End If

        Catch ex As Exception
            Console.WriteLine($"Error fetching room availability: {ex.Message}")
            Throw
        End Try

        Return availabilityData
    End Function

    Private Async Function GetRoomAvailabilityWithResponseAsync() As Task(Of (AvailabilityData As Dictionary(Of String, RoomAvailability), PropertyData As LittleHotelierResponse))
        Dim availabilityData As New Dictionary(Of String, RoomAvailability)
        Dim propertyData As LittleHotelierResponse = Nothing

        Try
            Dim startDate = GetBusinessDateTime().ToString("yyyy-MM-dd")
            Dim url = $"{apiBaseUrl}properties/{propertyId}/rates.json?start_date={startDate}"

            ' Existing API call code...
            Dim response = Await httpClient.GetAsync(url)

            If response.IsSuccessStatusCode Then
                Dim jsonContent = Await response.Content.ReadAsStringAsync()
                Dim apiDataArray = JsonConvert.DeserializeObject(Of List(Of LittleHotelierResponse))(jsonContent)

                If apiDataArray.Count > 0 Then
                    propertyData = apiDataArray(0)
                    availabilityData = ParseAvailabilityByDate(propertyData)
                End If
            End If
        Catch ex As Exception
            Console.WriteLine($"Error fetching room availability: {ex.Message}")
            Throw
        End Try

        Return (availabilityData, propertyData)
    End Function


    Private Function ParseAvailabilityByDate(propertyData As LittleHotelierResponse) As Dictionary(Of String, RoomAvailability)
        Dim availabilityByDate As New Dictionary(Of String, RoomAvailability)

        ' Get only 3 days: Today, Day +1, Day >+2
        Dim targetDates As New List(Of String)
        For i As Integer = 0 To 2
            targetDates.Add(GetBusinessDateTime().AddDays(i).ToString("yyyy-MM-dd"))
        Next

        ' Process each target date
        For Each targetDate In targetDates
            Dim availability As New RoomAvailability With {
            .CheckDate = targetDate,
            .DormBedsAvailable = 0,
            .PrivateRoomsAvailable = 0,
            .PrivateEnsuitesAvailable = 0
        }

            ' NEW: Capture current rates from API
            Dim currentApiRates As New PreviousRate With {
            .DormRegularRate = 0,
            .DormWalkInRate = 0,
            .PrivateRegularRate = 0,
            .PrivateWalkInRate = 0,
            .EnsuiteRegularRate = 0,
            .EnsuiteWalkInRate = 0
        }

            ' Sum up availability and capture rates by room category
            For Each ratePlan In propertyData.rate_plans
                Dim dateEntry = ratePlan.rate_plan_dates.FirstOrDefault(Function(d) d.date = targetDate)

                If dateEntry IsNot Nothing Then
                    Select Case ratePlan.name.ToLower()
                        Case "2 bed mixed dorm", "4 bed female dorm", "4 bed mixed dorm", "6 bed mixed dorm"
                            availability.DormBedsAvailable += dateEntry.available
                            ' Capture dorm rate from API (assuming this represents regular rate)
                            If currentApiRates.DormRegularRate = 0 Then
                                currentApiRates.DormRegularRate = CDbl(dateEntry.rate)
                            End If
                        Case "superior double (shared bathroom)"
                            availability.PrivateRoomsAvailable += dateEntry.available
                            If currentApiRates.PrivateRegularRate = 0 Then
                                currentApiRates.PrivateRegularRate = CDbl(dateEntry.rate)
                            End If
                        Case "superior queen ensuite"
                            availability.PrivateEnsuitesAvailable += dateEntry.available
                            If currentApiRates.EnsuiteRegularRate = 0 Then
                                currentApiRates.EnsuiteRegularRate = CDbl(dateEntry.rate)
                            End If
                    End Select
                End If
            Next

            availabilityByDate(targetDate) = availability

            ' Compare API rates vs calculated rates (simplified display)
            Dim calculatedRates = GetCurrentRates(availability, targetDate)

            ' Display rate comparison
            Dim daysAhead = DateDiff(DateInterval.Day, GetBusinessToday(), DateTime.Parse(targetDate))
            Dim dayLabel = If(daysAhead = 0, "Today", If(daysAhead = 1, "Day +1", "Day >+2"))

            Console.WriteLine($"Date: {targetDate} ({dayLabel})")
            Console.WriteLine($"  Dorms: {availability.DormBedsAvailable}/{totalDormBeds} available - " &
                    If(currentApiRates.DormRegularRate = calculatedRates.DormRegularRate,
                       $"RM{calculatedRates.DormRegularRate}/RM{calculatedRates.DormWalkInRate}",
                       $"RM{currentApiRates.DormRegularRate} → RM{calculatedRates.DormRegularRate} (Walk-in: RM{calculatedRates.DormWalkInRate})"))

            Console.WriteLine($"  Private: {availability.PrivateRoomsAvailable}/{totalPrivateRooms} available - " &
                    If(currentApiRates.PrivateRegularRate = calculatedRates.PrivateRegularRate,
                       $"RM{calculatedRates.PrivateRegularRate}/RM{calculatedRates.PrivateWalkInRate}",
                       $"RM{currentApiRates.PrivateRegularRate} → RM{calculatedRates.PrivateRegularRate} (Walk-in: RM{calculatedRates.PrivateWalkInRate})"))

            Console.WriteLine($"  Ensuite: {availability.PrivateEnsuitesAvailable}/{totalEnsuiteRooms} available - " &
                    If(currentApiRates.EnsuiteRegularRate = calculatedRates.EnsuiteRegularRate,
                       $"RM{calculatedRates.EnsuiteRegularRate}/RM{calculatedRates.EnsuiteWalkInRate}",
                       $"RM{currentApiRates.EnsuiteRegularRate} → RM{calculatedRates.EnsuiteRegularRate} (Walk-in: RM{calculatedRates.EnsuiteWalkInRate})"))

        Next

        Return availabilityByDate
    End Function

    Private Function GetApiRatesFromResponse(propertyData As LittleHotelierResponse, targetDate As String) As PreviousRate
        Dim apiRates As New PreviousRate With {
        .DormRegularRate = 0,
        .DormWalkInRate = 0,
        .PrivateRegularRate = 0,
        .PrivateWalkInRate = 0,
        .EnsuiteRegularRate = 0,
        .EnsuiteWalkInRate = 0
    }

        ' Add null checks for propertyData and rate_plans
        If propertyData Is Nothing OrElse propertyData.rate_plans Is Nothing Then
            Return apiRates
        End If

        For Each ratePlan In propertyData.rate_plans
            ' Check if ratePlan and its properties are not null
            If ratePlan Is Nothing OrElse ratePlan.rate_plan_dates Is Nothing OrElse String.IsNullOrEmpty(ratePlan.name) Then
                Continue For
            End If

            Dim dateEntry = ratePlan.rate_plan_dates.FirstOrDefault(Function(d) d IsNot Nothing AndAlso d.date = targetDate)

            ' Check if dateEntry was found and is not null
            If dateEntry Is Nothing Then
                Continue For
            End If

            Select Case ratePlan.name.ToLower()
                Case "2 bed mixed dorm", "4 bed female dorm", "4 bed mixed dorm", "6 bed mixed dorm"
                    ' Only set if not already set (first match wins)
                    If apiRates.DormRegularRate = 0 Then
                        apiRates.DormRegularRate = CDbl(dateEntry.rate)
                    End If
                Case "superior double (shared bathroom)"
                    If apiRates.PrivateRegularRate = 0 Then
                        apiRates.PrivateRegularRate = CDbl(dateEntry.rate)
                    End If
                Case "superior queen ensuite"
                    If apiRates.EnsuiteRegularRate = 0 Then
                        apiRates.EnsuiteRegularRate = CDbl(dateEntry.rate)
                    End If
            End Select
        Next

        Return apiRates
    End Function



    Private Function GetCurrentRates(avail As RoomAvailability, dateStr As String) As PreviousRate
        ' FIXED: Use DateTime.Today and DateDiff for accurate day calculation
        Dim daysAhead = DateDiff(DateInterval.Day, GetBusinessToday(), DateTime.Parse(dateStr))
        Dim dayPrefix As String

        If daysAhead = 0 Then
            dayPrefix = "Today"
        ElseIf daysAhead = 1 Then
            dayPrefix = "Day1_"
        Else
            dayPrefix = "Day2Plus_"
        End If

        ' Get Dorm rates
        Dim dormRates = GetRoomTypeRates("Dorm", dayPrefix, avail.DormBedsAvailable)

        ' Get Private rates  
        Dim privateRates = GetRoomTypeRates("Private", dayPrefix, avail.PrivateRoomsAvailable)

        ' Get Ensuite rates
        Dim ensuiteRates = GetRoomTypeRates("Ensuite", dayPrefix, avail.PrivateEnsuitesAvailable)

        Return New PreviousRate With {
        .DormRegularRate = dormRates.RegularRate,
        .DormWalkInRate = dormRates.WalkInRate,
        .PrivateRegularRate = privateRates.RegularRate,
        .PrivateWalkInRate = privateRates.WalkInRate,
        .EnsuiteRegularRate = ensuiteRates.RegularRate,
        .EnsuiteWalkInRate = ensuiteRates.WalkInRate
    }
    End Function


    Private Function GetRoomTypeRates(roomType As String, dayPrefix As String, available As Integer) As (RegularRate As Double, WalkInRate As Double)
        Dim configKey As String = ""

        Select Case roomType.ToLower()
            Case "dorm"
                If available >= 8 Then
                    configKey = $"{roomType}{dayPrefix}8Plus"
                ElseIf available >= 4 Then
                    configKey = $"{roomType}{dayPrefix}4to7"
                ElseIf available >= 2 Then
                    configKey = $"{roomType}{dayPrefix}2to3"
                Else
                    configKey = $"{roomType}{dayPrefix}1"
                End If
            Case "private"
                If available >= 3 Then
                    configKey = $"{roomType}{dayPrefix}3Rooms"
                ElseIf available >= 2 Then
                    configKey = $"{roomType}{dayPrefix}2Rooms"
                Else
                    configKey = $"{roomType}{dayPrefix}1Room"
                End If
            Case "ensuite"
                If available >= 2 Then
                    configKey = $"{roomType}{dayPrefix}2Rooms"
                Else
                    configKey = $"{roomType}{dayPrefix}1Room"
                End If
        End Select

        Dim rateString = ConfigurationManager.AppSettings(configKey)
        If Not String.IsNullOrEmpty(rateString) Then
            Dim rates = rateString.Split(","c)
            If rates.Length = 2 Then
                Return (Double.Parse(rates(0)), Double.Parse(rates(1)))
            End If
        End If

        ' Default rates if config not found
        Return (50, 40)
    End Function


    Private Function CalculateRateChanges(availabilityData As Dictionary(Of String, RoomAvailability), propertyData As LittleHotelierResponse) As List(Of RateChange)
        Dim changes As New List(Of RateChange)

        For Each kvp In availabilityData
            Dim dateStr = kvp.Key
            Dim availability = kvp.Value
            Dim daysAhead = DateDiff(DateInterval.Day, GetBusinessToday(), DateTime.Parse(dateStr))

            ' Only apply dynamic pricing for <15 days
            If daysAhead < 15 Then
                Dim calculatedRates = GetCurrentRates(availability, dateStr)
                Dim apiCurrentRates = GetApiRatesFromResponse(propertyData, dateStr)

                ' Check Dorm REGULAR rate changes only
                If calculatedRates.DormRegularRate <> apiCurrentRates.DormRegularRate Then
                    changes.Add(New RateChange With {
                    .CheckDate = dateStr,
                    .RoomType = GetAvailabilityDescription("Dorm", availability.DormBedsAvailable),
                    .OldRegularRate = apiCurrentRates.DormRegularRate,
                    .NewRegularRate = calculatedRates.DormRegularRate,
                    .OldWalkInRate = -1, ' Not used anymore
                    .NewWalkInRate = calculatedRates.DormWalkInRate, ' Show current walk-in rate
                    .AvailableUnits = availability.DormBedsAvailable,
                    .DaysAhead = daysAhead
                })
                End If

                ' Check Private REGULAR rate changes only
                If calculatedRates.PrivateRegularRate <> apiCurrentRates.PrivateRegularRate Then
                    changes.Add(New RateChange With {
                    .CheckDate = dateStr,
                    .RoomType = GetAvailabilityDescription("Private", availability.PrivateRoomsAvailable),
                    .OldRegularRate = apiCurrentRates.PrivateRegularRate,
                    .NewRegularRate = calculatedRates.PrivateRegularRate,
                    .OldWalkInRate = -1, ' Not used anymore
                    .NewWalkInRate = calculatedRates.PrivateWalkInRate, ' Show current walk-in rate
                    .AvailableUnits = availability.PrivateRoomsAvailable,
                    .DaysAhead = daysAhead
                })
                End If

                ' Check Ensuite REGULAR rate changes only
                If calculatedRates.EnsuiteRegularRate <> apiCurrentRates.EnsuiteRegularRate Then
                    changes.Add(New RateChange With {
                    .CheckDate = dateStr,
                    .RoomType = GetAvailabilityDescription("Ensuite", availability.PrivateEnsuitesAvailable),
                    .OldRegularRate = apiCurrentRates.EnsuiteRegularRate,
                    .NewRegularRate = calculatedRates.EnsuiteRegularRate,
                    .OldWalkInRate = -1, ' Not used anymore
                    .NewWalkInRate = calculatedRates.EnsuiteWalkInRate, ' Show current walk-in rate
                    .AvailableUnits = availability.PrivateEnsuitesAvailable,
                    .DaysAhead = daysAhead
                })
                End If
            End If
        Next

        Return changes
    End Function


    Private Function GetAvailabilityDescription(roomType As String, available As Integer) As String
        Select Case roomType.ToLower()
            Case "dorm"
                If available >= 8 Then
                    Return "Dorm Beds (8+ available)"
                ElseIf available >= 4 Then
                    Return "Dorm Beds (4-7 available)"
                ElseIf available >= 2 Then
                    Return "Dorm Beds (2-3 available)"
                Else
                    Return "Dorm Beds (1 available)"
                End If
            Case "private"
                If available >= 3 Then
                    Return "Private Rooms - Shared Bath (3 available)"
                ElseIf available >= 2 Then
                    Return "Private Rooms - Shared Bath (2 available)"
                Else
                    Return "Private Rooms - Shared Bath (1 available)"
                End If
            Case "ensuite"
                If available >= 2 Then
                    Return "Queen Ensuite (2 available)"
                Else
                    Return "Queen Ensuite (1 available)"
                End If
        End Select
        Return roomType
    End Function

    Public Async Function SendEmailNotificationAsync(changes As List(Of RateChange)) As Task
        Try
            Dim subject = "🏨 Red Inn Court - Rate Update"
            Dim body = BuildRateChangeEmailBody(changes)

            Await SendEmailAsync(subject, body, False)
            Console.WriteLine("Email notification sent successfully")

        Catch ex As Exception
            Console.WriteLine($"Error sending email notification: {ex.Message}")
        End Try
    End Function

    Private Function BuildRateChangeEmailBody(changes As List(Of RateChange)) As String
        Dim body As New StringBuilder()
        body.AppendLine($"<h2>🏨 RED INN COURT - RATE UPDATE 🏨</h2>")
        body.AppendLine($"<p><strong>📅 {DateTime.Now:yyyy-MM-dd HH:mm}</strong></p>")
        body.AppendLine($"<br>")

        For Each change In changes
            Dim dayLabel = If(change.DaysAhead = 0, "Today", If(change.DaysAhead = 1, "Tomorrow (Day +1)", $"Day >+2 ({change.DaysAhead} days ahead)"))

            body.AppendLine($"<div style='margin-bottom: 20px; padding: 15px; border-left: 4px solid #007bff; background-color: #f8f9fa;'>")
            body.AppendLine($"<h3>📅 {change.CheckDate} ({dayLabel})</h3>")
            body.AppendLine($"<p><strong>🏠 Room Type:</strong> {change.RoomType}</p>")
            body.AppendLine($"<p><strong>🛏️ Available Units:</strong> {change.AvailableUnits}</p>")

            If change.OldRegularRate = 0 Then
                ' First time setting rates (API had 0)
                body.AppendLine($"<p><strong>💰 New Rates:</strong></p>")
                body.AppendLine($"<p>• Regular: <span style='color: #28a745; font-size: 1.1em;'>RM{change.NewRegularRate}</span></p>")
                body.AppendLine($"<p>• Walk-in: <span style='color: #6c757d; font-size: 1.0em;'>RM{change.NewWalkInRate}</span></p>")
            Else
                ' Rate change detected
                body.AppendLine($"<p><strong>💰 Rate Change:</strong></p>")
                body.AppendLine($"<p>• Regular: RM{change.OldRegularRate} → <span style='color: #28a745; font-size: 1.1em;'>RM{change.NewRegularRate}</span></p>")
                body.AppendLine($"<p>• Walk-in: <span style='color: #6c757d; font-size: 1.0em;'>RM{change.NewWalkInRate}</span> <em>(New rate)</em></p>")
            End If
            body.AppendLine($"</div>")
        Next

        body.AppendLine($"<br><hr>")
        body.AppendLine($"<p><em>Generated by Dynamic Pricing Bot 🤖</em></p>")
        body.AppendLine($"<p style='font-size: 0.9em; color: #6c757d;'><em>Note: Walk-in rates are automatically calculated from regular rates and shown for reference.</em></p>")

        Return body.ToString()
    End Function


    Public Async Function SendEmailErrorNotificationAsync(errorMessage As String) As Task
        Try
            Dim subject = "🚨 Dynamic Pricing Error Alert"
            Dim body = $"<h2>🚨 DYNAMIC PRICING ERROR 🚨</h2>" &
                      $"<div style='padding: 15px; background-color: #f8d7da; border: 1px solid #f5c6cb; border-radius: 5px; color: #721c24;'>" &
                      $"<p><strong>❌ Error:</strong> {errorMessage}</p>" &
                      $"<p><strong>⏰ Time:</strong> {DateTime.Now:yyyy-MM-dd HH:mm}</p>" &
                      $"</div>" &
                      $"<p>Please check the application logs for more details.</p>"

            Await SendEmailAsync(subject, body, True)
            Console.WriteLine("Error email notification sent successfully")

        Catch ex As Exception
            Console.WriteLine($"Error sending email error notification: {ex.Message}")
        End Try
    End Function

    Private Async Function SendEmailAsync(subject As String, body As String, isError As Boolean) As Task
        Try
            Using smtpClient As New SmtpClient(smtpHost, smtpPort)
                smtpClient.EnableSsl = True
                smtpClient.Credentials = New NetworkCredential(smtpUsername, smtpPassword)

                Using message As New MailMessage()
                    message.From = New MailAddress(emailFromAddress, emailFromName)
                    message.To.Add(emailToAddress)
                    message.Subject = subject
                    message.Body = body
                    message.IsBodyHtml = True

                    If isError Then
                        message.Priority = MailPriority.High
                    End If

                    Await smtpClient.SendMailAsync(message)
                End Using
            End Using

        Catch ex As Exception
            Console.WriteLine($"Error sending email: {ex.Message}")
            Throw
        End Try
    End Function

End Class
