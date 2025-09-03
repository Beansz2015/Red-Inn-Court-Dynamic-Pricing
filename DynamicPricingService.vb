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

    ' Storage file for rates
    Private ReadOnly rateStorageFile As String = "previous_rates.json"

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

        Try
            Console.WriteLine($"Starting dynamic pricing check at {DateTime.Now}")

            ' Get availability for next 3 days
            Dim availabilityData = Await GetRoomAvailabilityAsync()

            If availabilityData.Count = 0 Then
                Console.WriteLine("No availability data retrieved. Exiting.")
                Return
            End If

            ' Calculate new rates and detect changes
            Dim rateChanges = CalculateRateChanges(availabilityData)

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

            ' Sum up availability by room category for this specific date
            For Each ratePlan In propertyData.rate_plans
                Dim dateEntry = ratePlan.rate_plan_dates.FirstOrDefault(Function(d) d.date = targetDate)

                If dateEntry IsNot Nothing Then
                    Select Case ratePlan.name.ToLower()
                        Case "2 bed mixed dorm", "4 bed female dorm", "4 bed mixed dorm", "6 bed mixed dorm"
                            availability.DormBedsAvailable += dateEntry.available
                        Case "superior double (shared bathroom)"
                            availability.PrivateRoomsAvailable += dateEntry.available
                        Case "superior queen ensuite"
                            availability.PrivateEnsuitesAvailable += dateEntry.available
                    End Select
                End If
            Next

            availabilityByDate(targetDate) = availability

            ' Display current rates
            Dim prevRates = GetPreviousRates(targetDate)
            Dim currRates = GetCurrentRates(availability, targetDate)

            ' FIXED: Use DateTime.Today instead of DateTime.Now for accurate day calculation
            Dim daysAhead = DateDiff(DateInterval.Day, GetBusinessToday(), DateTime.Parse(targetDate))
            Dim dayLabel = If(daysAhead = 0, "Today", If(daysAhead = 1, "Day +1", "Day >+2"))

            Console.WriteLine($"Date: {targetDate} ({dayLabel})")
            Console.WriteLine($"  Dorms: {availability.DormBedsAvailable}/{totalDormBeds} available - " &
                        If(prevRates.DormRegularRate = -1 Or (prevRates.DormRegularRate = currRates.DormRegularRate And prevRates.DormWalkInRate = currRates.DormWalkInRate),
                           $"RM{currRates.DormRegularRate}/RM{currRates.DormWalkInRate}",
                           $"RM{prevRates.DormRegularRate}/RM{prevRates.DormWalkInRate} → RM{currRates.DormRegularRate}/RM{currRates.DormWalkInRate}"))

            Console.WriteLine($"  Private: {availability.PrivateRoomsAvailable}/{totalPrivateRooms} available - " &
                        If(prevRates.PrivateRegularRate = -1 Or (prevRates.PrivateRegularRate = currRates.PrivateRegularRate And prevRates.PrivateWalkInRate = currRates.PrivateWalkInRate),
                           $"RM{currRates.PrivateRegularRate}/RM{currRates.PrivateWalkInRate}",
                           $"RM{prevRates.PrivateRegularRate}/RM{prevRates.PrivateWalkInRate} → RM{currRates.PrivateRegularRate}/RM{currRates.PrivateWalkInRate}"))

            Console.WriteLine($"  Ensuite: {availability.PrivateEnsuitesAvailable}/{totalEnsuiteRooms} available - " &
                        If(prevRates.EnsuiteRegularRate = -1 Or (prevRates.EnsuiteRegularRate = currRates.EnsuiteRegularRate And prevRates.EnsuiteWalkInRate = currRates.EnsuiteWalkInRate),
                           $"RM{currRates.EnsuiteRegularRate}/RM{currRates.EnsuiteWalkInRate}",
                           $"RM{prevRates.EnsuiteRegularRate}/RM{prevRates.EnsuiteWalkInRate} → RM{currRates.EnsuiteRegularRate}/RM{currRates.EnsuiteWalkInRate}"))
        Next

        Return availabilityByDate
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

    Private Function GetPreviousRates(dateStr As String) As PreviousRate
        Try
            If File.Exists(rateStorageFile) Then
                Dim json = File.ReadAllText(rateStorageFile)
                Dim allRates = JsonConvert.DeserializeObject(Of Dictionary(Of String, PreviousRate))(json)

                If allRates.ContainsKey(dateStr) Then
                    Return allRates(dateStr)
                End If
            End If
        Catch ex As Exception
            Console.WriteLine($"Error reading previous rates: {ex.Message}")
        End Try

        Return New PreviousRate With {
            .DormRegularRate = -1,
            .DormWalkInRate = -1,
            .PrivateRegularRate = -1,
            .PrivateWalkInRate = -1,
            .EnsuiteRegularRate = -1,
            .EnsuiteWalkInRate = -1
        }
    End Function

    Private Function CalculateRateChanges(availabilityData As Dictionary(Of String, RoomAvailability)) As List(Of RateChange)
        Dim changes As New List(Of RateChange)

        For Each kvp In availabilityData
            Dim dateStr = kvp.Key
            Dim availability = kvp.Value
            ' FIXED: Use DateTime.Today and DateDiff for accurate day calculation
            Dim daysAhead = DateDiff(DateInterval.Day, GetBusinessToday(), DateTime.Parse(dateStr))

            ' Only apply dynamic pricing for <15 days
            If daysAhead < 15 Then
                Dim currentRates = GetCurrentRates(availability, dateStr)
                Dim previousRates = GetPreviousRates(dateStr)
                ' Check Dorm changes
                If currentRates.DormRegularRate <> previousRates.DormRegularRate Or currentRates.DormWalkInRate <> previousRates.DormWalkInRate Then
                    changes.Add(New RateChange With {
                        .CheckDate = dateStr,
                        .RoomType = GetAvailabilityDescription("Dorm", availability.DormBedsAvailable),
                        .OldRegularRate = previousRates.DormRegularRate,
                        .NewRegularRate = currentRates.DormRegularRate,
                        .OldWalkInRate = previousRates.DormWalkInRate,
                        .NewWalkInRate = currentRates.DormWalkInRate,
                        .AvailableUnits = availability.DormBedsAvailable,
                        .DaysAhead = daysAhead
                    })
                End If

                ' Check Private changes
                If currentRates.PrivateRegularRate <> previousRates.PrivateRegularRate Or currentRates.PrivateWalkInRate <> previousRates.PrivateWalkInRate Then
                    changes.Add(New RateChange With {
                        .CheckDate = dateStr,
                        .RoomType = GetAvailabilityDescription("Private", availability.PrivateRoomsAvailable),
                        .OldRegularRate = previousRates.PrivateRegularRate,
                        .NewRegularRate = currentRates.PrivateRegularRate,
                        .OldWalkInRate = previousRates.PrivateWalkInRate,
                        .NewWalkInRate = currentRates.PrivateWalkInRate,
                        .AvailableUnits = availability.PrivateRoomsAvailable,
                        .DaysAhead = daysAhead
                    })
                End If

                ' Check Ensuite changes
                If currentRates.EnsuiteRegularRate <> previousRates.EnsuiteRegularRate Or currentRates.EnsuiteWalkInRate <> previousRates.EnsuiteWalkInRate Then
                    changes.Add(New RateChange With {
                        .CheckDate = dateStr,
                        .RoomType = GetAvailabilityDescription("Ensuite", availability.PrivateEnsuitesAvailable),
                        .OldRegularRate = previousRates.EnsuiteRegularRate,
                        .NewRegularRate = currentRates.EnsuiteRegularRate,
                        .OldWalkInRate = previousRates.EnsuiteWalkInRate,
                        .NewWalkInRate = currentRates.EnsuiteWalkInRate,
                        .AvailableUnits = availability.PrivateEnsuitesAvailable,
                        .DaysAhead = daysAhead
                    })
                End If

                ' Update stored rates
                UpdateStoredRates(dateStr, currentRates)
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

    Private Sub UpdateStoredRates(dateStr As String, rates As PreviousRate)
        Try
            Dim allRates As New Dictionary(Of String, PreviousRate)

            If File.Exists(rateStorageFile) Then
                Dim json = File.ReadAllText(rateStorageFile)
                allRates = JsonConvert.DeserializeObject(Of Dictionary(Of String, PreviousRate))(json)
            End If

            rates.LastUpdated = DateTime.Now
            allRates(dateStr) = rates

            ' Clean up old entries (older than 7 days)
            Dim cutoffDate = DateTime.Now.AddDays(-7).ToString("yyyy-MM-dd")
            Dim keysToRemove = allRates.Keys.Where(Function(k) String.Compare(k, cutoffDate) < 0).ToList()
            For Each key In keysToRemove
                allRates.Remove(key)
            Next

            Dim updatedJson = JsonConvert.SerializeObject(allRates, Formatting.Indented)
            File.WriteAllText(rateStorageFile, updatedJson)

        Catch ex As Exception
            Console.WriteLine($"Error updating stored rates: {ex.Message}")
        End Try
    End Sub

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

            If change.OldRegularRate = -1 Then
                body.AppendLine($"<p><strong>💰 New Rates:</strong></p>")
                body.AppendLine($"<p>• Regular: <span style='color: #28a745; font-size: 1.1em;'>RM{change.NewRegularRate}</span></p>")
                body.AppendLine($"<p>• Walk-in: <span style='color: #28a745; font-size: 1.1em;'>RM{change.NewWalkInRate}</span></p>")
            Else
                body.AppendLine($"<p><strong>💰 Rate Changes:</strong></p>")
                body.AppendLine($"<p>• Regular: RM{change.OldRegularRate} → <span style='color: #28a745; font-size: 1.1em;'>RM{change.NewRegularRate}</span></p>")
                body.AppendLine($"<p>• Walk-in: RM{change.OldWalkInRate} → <span style='color: #28a745; font-size: 1.1em;'>RM{change.NewWalkInRate}</span></p>")
            End If
            body.AppendLine($"</div>")
        Next

        body.AppendLine($"<br><hr>")
        body.AppendLine($"<p><em>Generated by Dynamic Pricing Bot 🤖</em></p>")
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
