Imports System.Net.Http
Imports System.Threading.Tasks
Imports Newtonsoft.Json
Imports System.Text
Imports System.IO
Imports System.Configuration

Public Class DynamicPricingService
    Private ReadOnly httpClient As New HttpClient()

    ' Configuration settings
    Private ReadOnly apiKey As String = ConfigurationManager.AppSettings("LittleHotelierApiKey")
    Private ReadOnly propertyId As String = ConfigurationManager.AppSettings("PropertyId")
    Private ReadOnly apiBaseUrl As String = ConfigurationManager.AppSettings("ApiBaseUrl")

    ' Twilio WhatsApp Settings
    Private ReadOnly twilioAccountSid As String = ConfigurationManager.AppSettings("TwilioAccountSid")
    Private ReadOnly twilioAuthToken As String = ConfigurationManager.AppSettings("TwilioAuthToken")
    Private ReadOnly twilioFromNumber As String = ConfigurationManager.AppSettings("TwilioFromNumber")
    Private ReadOnly notificationRecipient As String = ConfigurationManager.AppSettings("NotificationRecipient")

    ' Room capacity configuration with temporary closures
    Private ReadOnly baseDormBeds As Integer = 20
    Private ReadOnly basePrivateRooms As Integer = 3
    Private ReadOnly baseEnsuiteRooms As Integer = 2
    Private ReadOnly baseTwinRooms As Integer = 1

    ' Calculate effective capacity minus temporarily closed units
    Private ReadOnly totalDormBeds As Integer = baseDormBeds - CInt(If(ConfigurationManager.AppSettings("TemporaryClosedDormBeds"), "0"))
    Private ReadOnly totalPrivateRooms As Integer = basePrivateRooms - CInt(If(ConfigurationManager.AppSettings("TemporaryClosedPrivateRooms"), "0"))
    Private ReadOnly totalEnsuiteRooms As Integer = baseEnsuiteRooms - CInt(If(ConfigurationManager.AppSettings("TemporaryClosedEnsuiteRooms"), "0"))
    Private ReadOnly totalTwinRooms As Integer = baseTwinRooms - CInt(If(ConfigurationManager.AppSettings("TemporaryClosedTwinRooms"), "0"))

    ' Storage file for rates
    Private ReadOnly rateStorageFile As String = "previous_rates.json"

    ' Rate profile class
    Private Class RateProfile
        Public Thresholds As Double()
        Public Rates As Double()
    End Class

    Private ReadOnly rateProfiles As Dictionary(Of String, RateProfile)

    '--------------------------------------------
    ' Constructor – runs once per service launch
    '--------------------------------------------
    Public Sub New()
        ' Build the rate profiles dictionary from App.config
        rateProfiles = New Dictionary(Of String, RateProfile)(StringComparer.OrdinalIgnoreCase) From {
            {"Dorm", LoadRateProfile("Dorm")},
            {"Private", LoadRateProfile("Private")},
            {"Ensuite", LoadRateProfile("Ensuite")},
            {"Twin", LoadRateProfile("Twin")}
        }
    End Sub

    Private Function LoadRateProfile(tag As String) As RateProfile
        Dim t = ConfigurationManager.AppSettings($"Thresholds{tag}")?.Split(","c)
        Dim r = ConfigurationManager.AppSettings($"Rates{tag}")?.Split(","c)

        ' Handle empty thresholds (like Twin)
        If String.IsNullOrWhiteSpace(ConfigurationManager.AppSettings($"Thresholds{tag}")) Then
            t = New String() {}
        End If

        If t Is Nothing OrElse r Is Nothing OrElse r.Length <> t.Length + 1 Then
            ' Fall back to hard-coded defaults
            Select Case tag.ToLower()
                Case "dorm"
                    t = {"60", "80", "90"}
                    r = {"42", "46", "50", "54"}
                Case "private"
                    t = {"30", "60"}
                    r = {"108", "112", "116"}
                Case "ensuite"
                    t = {"50"}
                    r = {"130", "150"}
                Case "twin"
                    t = New String() {}
                    r = {"89"}
                Case Else
                    t = {"60", "80", "90"}
                    r = {"42", "46", "50", "54"}
            End Select
        End If

        Return New RateProfile With {
            .Thresholds = If(t.Length > 0, t.Select(Function(x) Double.Parse(x.Trim())).ToArray(), New Double() {}),
            .Rates = r.Select(Function(x) Double.Parse(x.Trim())).ToArray()
        }
    End Function

    Private Function GetRate(roomTag As String, occPct As Double) As Double
        Dim p = rateProfiles(roomTag)
        For i = p.Thresholds.Length - 1 To 0 Step -1
            If occPct >= p.Thresholds(i) Then Return p.Rates(i + 1)
        Next
        Return p.Rates(0)
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

            ' Send Twilio WhatsApp notifications if there are changes
            If rateChanges.Any() Then
                Await SendTwilioWhatsAppAsync(rateChanges)
                Console.WriteLine($"Sent Twilio WhatsApp notification for {rateChanges.Count} rate changes")
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
                Await SendTwilioErrorNotificationAsync(errorToReport)
            Catch notificationEx As Exception
                Console.WriteLine($"Failed to send Twilio error notification: {notificationEx.Message}")
            End Try
        End If
    End Function

    Public Async Function GetRoomAvailabilityAsync() As Task(Of Dictionary(Of String, RoomAvailability))
        Dim availabilityData As New Dictionary(Of String, RoomAvailability)

        Try
            Dim startDate = DateTime.Now.ToString("yyyy-MM-dd")
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

        ' Get current day plus next 3 days (total 4 days)
        Dim targetDates As New List(Of String)
        For i As Integer = 0 To 3
            targetDates.Add(DateTime.Now.AddDays(i).ToString("yyyy-MM-dd"))
        Next

        ' Process each target date
        For Each targetDate In targetDates
            Dim availability As New RoomAvailability With {
                .CheckDate = targetDate,
                .DormBedsAvailable = 0,
                .PrivateRoomsAvailable = 0,
                .PrivateEnsuitesAvailable = 0,
                .TwinRoomsAvailable = 0
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
                        Case "twin (shared bathroom)"
                            availability.TwinRoomsAvailable += dateEntry.available
                    End Select
                End If
            Next

            ' Calculate occupancy percentages
            availability.DormOccupancyPct = ((totalDormBeds - availability.DormBedsAvailable) / totalDormBeds) * 100
            availability.PrivateRoomsOccupancyPct = ((totalPrivateRooms - availability.PrivateRoomsAvailable) / totalPrivateRooms) * 100
            availability.PrivateEnsuitesOccupancyPct = ((totalEnsuiteRooms - availability.PrivateEnsuitesAvailable) / totalEnsuiteRooms) * 100
            availability.TwinRoomsOccupancyPct = ((totalTwinRooms - availability.TwinRoomsAvailable) / totalTwinRooms) * 100

            ' Ensure occupancy doesn't go below 0%
            availability.DormOccupancyPct = Math.Max(0, availability.DormOccupancyPct)
            availability.PrivateRoomsOccupancyPct = Math.Max(0, availability.PrivateRoomsOccupancyPct)
            availability.PrivateEnsuitesOccupancyPct = Math.Max(0, availability.PrivateEnsuitesOccupancyPct)
            availability.TwinRoomsOccupancyPct = Math.Max(0, availability.TwinRoomsOccupancyPct)

            availabilityByDate(targetDate) = availability

            ' Display current rates
            Dim prevRates = GetPreviousRates(targetDate)
            Dim currRates = GetCurrentRates(availability)

            Console.WriteLine($"Date: {targetDate}")
            Console.WriteLine($"  Dorms: {availability.DormBedsAvailable}/{totalDormBeds} available ({availability.DormOccupancyPct:F1}% occupied) - " &
                            If(prevRates.DormRate = -1 Or prevRates.DormRate = currRates.DormRate,
                               $"RM{currRates.DormRate}",
                               $"RM{prevRates.DormRate} → RM{currRates.DormRate}"))

            Console.WriteLine($"  Private: {availability.PrivateRoomsAvailable}/{totalPrivateRooms} available ({availability.PrivateRoomsOccupancyPct:F1}% occupied) - " &
                            If(prevRates.PrivateRate = -1 Or prevRates.PrivateRate = currRates.PrivateRate,
                               $"RM{currRates.PrivateRate}",
                               $"RM{prevRates.PrivateRate} → RM{currRates.PrivateRate}"))

            Console.WriteLine($"  Ensuite: {availability.PrivateEnsuitesAvailable}/{totalEnsuiteRooms} available ({availability.PrivateEnsuitesOccupancyPct:F1}% occupied) - " &
                            If(prevRates.EnsuiteRate = -1 Or prevRates.EnsuiteRate = currRates.EnsuiteRate,
                               $"RM{currRates.EnsuiteRate}",
                               $"RM{prevRates.EnsuiteRate} → RM{currRates.EnsuiteRate}"))

            Console.WriteLine($"  Twin: {availability.TwinRoomsAvailable}/{totalTwinRooms} available ({availability.TwinRoomsOccupancyPct:F1}% occupied) - " &
                            If(prevRates.TwinRate = -1 Or prevRates.TwinRate = currRates.TwinRate,
                               $"RM{currRates.TwinRate}",
                               $"RM{prevRates.TwinRate} → RM{currRates.TwinRate}"))
        Next

        Return availabilityByDate
    End Function

    Private Function GetCurrentRates(avail As RoomAvailability) As PreviousRate
        Return New PreviousRate With {
            .DormRate = GetRate("Dorm", avail.DormOccupancyPct),
            .PrivateRate = GetRate("Private", avail.PrivateRoomsOccupancyPct),
            .EnsuiteRate = GetRate("Ensuite", avail.PrivateEnsuitesOccupancyPct),
            .TwinRate = GetRate("Twin", avail.TwinRoomsOccupancyPct)
        }
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
            .DormRate = -1,
            .PrivateRate = -1,
            .EnsuiteRate = -1,
            .TwinRate = -1
        }
    End Function

    Private Function CalculateRateChanges(availabilityData As Dictionary(Of String, RoomAvailability)) As List(Of RateChange)
        Dim changes As New List(Of RateChange)

        For Each kvp In availabilityData
            Dim dateStr = kvp.Key
            Dim availability = kvp.Value
            Dim DaysAhead = (DateTime.Parse(dateStr) - DateTime.Now).Days

            ' Only apply dynamic pricing for <15 days
            If DaysAhead < 15 Then
                ' Calculate new rates based on occupancy thresholds
                Dim newDormRate = GetRate("Dorm", availability.DormOccupancyPct)
                Dim newPrivateRate = GetRate("Private", availability.PrivateRoomsOccupancyPct)
                Dim newEnsuiteRate = GetRate("Ensuite", availability.PrivateEnsuitesOccupancyPct)
                Dim newTwinRate = GetRate("Twin", availability.TwinRoomsOccupancyPct)

                ' Get previous rates from storage
                Dim previousRates = GetPreviousRates(dateStr)

                ' Check for changes and create notifications
                If Math.Abs(newDormRate - previousRates.DormRate) > 0 Then
                    changes.Add(New RateChange With {
                        .CheckDate = dateStr,
                        .RoomType = "Dorm Beds",
                        .OldRate = previousRates.DormRate,
                        .NewRate = newDormRate,
                        .OccupancyPct = availability.DormOccupancyPct,
                        .AvailableUnits = availability.DormBedsAvailable,
                        .DaysAhead = DaysAhead
                    })
                End If

                If Math.Abs(newPrivateRate - previousRates.PrivateRate) > 0 Then
                    changes.Add(New RateChange With {
                        .CheckDate = dateStr,
                        .RoomType = "Private Rooms (Shared Bath)",
                        .OldRate = previousRates.PrivateRate,
                        .NewRate = newPrivateRate,
                        .OccupancyPct = availability.PrivateRoomsOccupancyPct,
                        .AvailableUnits = availability.PrivateRoomsAvailable,
                        .DaysAhead = DaysAhead
                    })
                End If

                If Math.Abs(newEnsuiteRate - previousRates.EnsuiteRate) > 0 Then
                    changes.Add(New RateChange With {
                        .CheckDate = dateStr,
                        .RoomType = "Queen Ensuite",
                        .OldRate = previousRates.EnsuiteRate,
                        .NewRate = newEnsuiteRate,
                        .OccupancyPct = availability.PrivateEnsuitesOccupancyPct,
                        .AvailableUnits = availability.PrivateEnsuitesAvailable,
                        .DaysAhead = DaysAhead
                    })
                End If

                If Math.Abs(newTwinRate - previousRates.TwinRate) > 0 Then
                    changes.Add(New RateChange With {
                        .CheckDate = dateStr,
                        .RoomType = "Twin Room (Shared Bath)",
                        .OldRate = previousRates.TwinRate,
                        .NewRate = newTwinRate,
                        .OccupancyPct = availability.TwinRoomsOccupancyPct,
                        .AvailableUnits = availability.TwinRoomsAvailable,
                        .DaysAhead = DaysAhead
                    })
                End If

                ' Update stored rates
                UpdateStoredRates(dateStr, newDormRate, newPrivateRate, newEnsuiteRate, newTwinRate)
            End If
        Next

        Return changes
    End Function

    Private Sub UpdateStoredRates(dateStr As String, dormRate As Double, privateRate As Double, ensuiteRate As Double, twinRate As Double)
        Try
            Dim allRates As New Dictionary(Of String, PreviousRate)

            If File.Exists(rateStorageFile) Then
                Dim json = File.ReadAllText(rateStorageFile)
                allRates = JsonConvert.DeserializeObject(Of Dictionary(Of String, PreviousRate))(json)
            End If

            allRates(dateStr) = New PreviousRate With {
                .DormRate = dormRate,
                .PrivateRate = privateRate,
                .EnsuiteRate = ensuiteRate,
                .TwinRate = twinRate,
                .LastUpdated = DateTime.Now
            }

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

    Public Async Function SendTwilioWhatsAppAsync(changes As List(Of RateChange)) As Task
        Try
            Dim message = BuildRateChangeMessage(changes)

            Dim url = $"https://api.twilio.com/2010-04-01/Accounts/{twilioAccountSid}/Messages.json"
            Dim credentials = Convert.ToBase64String(Encoding.ASCII.GetBytes($"{twilioAccountSid}:{twilioAuthToken}"))

            Dim formData = New List(Of KeyValuePair(Of String, String)) From {
                New KeyValuePair(Of String, String)("From", twilioFromNumber),
                New KeyValuePair(Of String, String)("To", notificationRecipient),
                New KeyValuePair(Of String, String)("Body", message)
            }

            Dim content = New FormUrlEncodedContent(formData)

            httpClient.DefaultRequestHeaders.Clear()
            httpClient.DefaultRequestHeaders.Add("Authorization", $"Basic {credentials}")

            Dim response = Await httpClient.PostAsync(url, content)

            If response.IsSuccessStatusCode Then
                Console.WriteLine("Twilio WhatsApp message sent successfully")
            Else
                Dim errorContent = Await response.Content.ReadAsStringAsync()
                Console.WriteLine($"Twilio Error: {response.StatusCode} - {errorContent}")
            End If

        Catch ex As Exception
            Console.WriteLine($"Error sending Twilio WhatsApp: {ex.Message}")
        End Try
    End Function

    Private Function BuildRateChangeMessage(changes As List(Of RateChange)) As String
        Dim message As New StringBuilder()
        message.AppendLine("🏨 *RED INN COURT - RATE UPDATE* 🏨")
        message.AppendLine($"📅 {DateTime.Now:yyyy-MM-dd HH:mm}")
        message.AppendLine("")

        For Each change In changes
            message.AppendLine($"📅 *{change.CheckDate}* ({change.DaysAhead} days ahead)")
            message.AppendLine($"🏠 {change.RoomType}")
            message.AppendLine($"📊 Occupancy: {change.OccupancyPct:F1}%")
            message.AppendLine($"🛏️ Available: {change.AvailableUnits} units")

            If change.OldRate = -1 Then
                message.AppendLine($"💰 New Rate: *RM{change.NewRate}*")
            Else
                message.AppendLine($"💰 Rate: RM{change.OldRate} → *RM{change.NewRate}*")
            End If
            message.AppendLine("")
        Next

        message.AppendLine("Generated by Dynamic Pricing Bot 🤖")
        Return message.ToString()
    End Function

    Public Async Function SendTwilioErrorNotificationAsync(errorMessage As String) As Task
        Try
            Dim message = $"🚨 *DYNAMIC PRICING ERROR* 🚨{vbCrLf}{vbCrLf}❌ {errorMessage}{vbCrLf}{vbCrLf}⏰ {DateTime.Now:yyyy-MM-dd HH:mm}{vbCrLf}{vbCrLf}Please check the application logs for more details."

            Dim url = $"https://api.twilio.com/2010-04-01/Accounts/{twilioAccountSid}/Messages.json"
            Dim credentials = Convert.ToBase64String(Encoding.ASCII.GetBytes($"{twilioAccountSid}:{twilioAuthToken}"))

            Dim formData = New List(Of KeyValuePair(Of String, String)) From {
                New KeyValuePair(Of String, String)("From", twilioFromNumber),
                New KeyValuePair(Of String, String)("To", notificationRecipient),
                New KeyValuePair(Of String, String)("Body", message)
            }

            Dim content = New FormUrlEncodedContent(formData)

            httpClient.DefaultRequestHeaders.Clear()
            httpClient.DefaultRequestHeaders.Add("Authorization", $"Basic {credentials}")

            Await httpClient.PostAsync(url, content)

        Catch ex As Exception
            Console.WriteLine($"Error sending Twilio error notification: {ex.Message}")
        End Try
    End Function

End Class
