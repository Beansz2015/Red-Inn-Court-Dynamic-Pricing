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

    ' REPLACE WhatsApp settings with Twilio settings
    Private ReadOnly twilioAccountSid As String = ConfigurationManager.AppSettings("TwilioAccountSid")
    Private ReadOnly twilioAuthToken As String = ConfigurationManager.AppSettings("TwilioAuthToken")
    Private ReadOnly twilioFromNumber As String = ConfigurationManager.AppSettings("TwilioFromNumber")
    Private ReadOnly notificationRecipient As String = ConfigurationManager.AppSettings("NotificationRecipient")

    ' Room capacity configuration - keep as is
    Private ReadOnly totalDormBeds As Integer = 20
    Private ReadOnly totalPrivateRooms As Integer = 3
    Private ReadOnly totalEnsuiteRooms As Integer = 2
    Private ReadOnly totalTwinRooms As Integer = 1

    ' Storage file - keep as is
    Private ReadOnly discountStorageFile As String = "previous_discounts.json"


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

            ' Calculate new discounts and detect changes
            Dim discountChanges = CalculateDiscountChanges(availabilityData)

            ' Send Twilio WhatsApp notifications if there are changes
            If discountChanges.Any() Then
                Await SendTwilioWhatsAppAsync(discountChanges)  ' CHANGED from SendWhatsAppNotificationAsync
                Console.WriteLine($"Sent Twilio WhatsApp notification for {discountChanges.Count} discount changes")
            Else
                Console.WriteLine("No discount changes detected")
            End If

            Console.WriteLine("Dynamic pricing check completed successfully")

        Catch ex As Exception
            Console.WriteLine($"Error in dynamic pricing check: {ex.Message}")
            Console.WriteLine($"Stack trace: {ex.StackTrace}")

            ' Store error for notification outside try-catch
            errorToReport = ex.Message
        End Try

        ' Send error notification if there was an error
        If errorToReport IsNot Nothing Then
            Try
                Await SendTwilioErrorNotificationAsync(errorToReport)  ' CHANGED from SendErrorNotificationAsync
            Catch notificationEx As Exception
                Console.WriteLine($"Failed to send Twilio error notification: {notificationEx.Message}")
            End Try
        End If
    End Function



    Public Async Function GetRoomAvailabilityAsync() As Task(Of Dictionary(Of String, RoomAvailability))
        Dim availabilityData As New Dictionary(Of String, RoomAvailability)

        Try
            ' Correct LittleHotelier API URL format
            Dim startDate = DateTime.Now.ToString("yyyy-MM-dd")  ' Start from today
            Dim url = $"{apiBaseUrl}properties/{propertyId}/rates.json?start_date={startDate}"

            Console.WriteLine($"API URL: {url}")  ' Debug logging

            httpClient.DefaultRequestHeaders.Clear()
            ' Remove authorization header - no API key needed
            httpClient.DefaultRequestHeaders.Add("Accept", "application/json")
            httpClient.DefaultRequestHeaders.Add("User-Agent", "RedInnDynamicPricing/1.0")

            Dim response = Await httpClient.GetAsync(url)

            If response.IsSuccessStatusCode Then
                Dim jsonContent = Await response.Content.ReadAsStringAsync()
                Console.WriteLine("Successfully retrieved API data")
                Console.WriteLine($"Response length: {jsonContent.Length} characters")

                Dim apiDataArray = JsonConvert.DeserializeObject(Of List(Of LittleHotelierResponse))(jsonContent)

                ' Process the first property (Red Inn Court)
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

        ' Get next 3 days from today
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

            ' ---- NEW VERBOSE OUTPUT WITH DISCOUNT INFO ----
            Dim prev = GetPreviousDiscounts(targetDate)
            Dim curr = GetCurrentDiscounts(availability)

            Console.WriteLine($"Date: {targetDate}")

            ' Dorms
            Dim dormLine = $"  Dorms: {availability.DormBedsAvailable}/{totalDormBeds} available " &
                           $"({availability.DormOccupancyPct:F1}% occupied) - "
            If prev.DormDiscount = -1 Or prev.DormDiscount = curr.DormDiscount Then
                dormLine &= $"{curr.DormDiscount}% discount"
            Else
                dormLine &= $"{prev.DormDiscount}% → {curr.DormDiscount}% discount"
            End If
            Console.WriteLine(dormLine)

            ' Private rooms
            Dim privLine = $"  Private: {availability.PrivateRoomsAvailable}/{totalPrivateRooms} available " &
                           $"({availability.PrivateRoomsOccupancyPct:F1}% occupied) - "
            If prev.PrivateDiscount = -1 Or prev.PrivateDiscount = curr.PrivateDiscount Then
                privLine &= $"{curr.PrivateDiscount}% discount"
            Else
                privLine &= $"{prev.PrivateDiscount}% → {curr.PrivateDiscount}% discount"
            End If
            Console.WriteLine(privLine)

            ' Ensuite rooms
            Dim ensLine = $"  Ensuite: {availability.PrivateEnsuitesAvailable}/{totalEnsuiteRooms} available " &
                          $"({availability.PrivateEnsuitesOccupancyPct:F1}% occupied) - "
            If prev.EnsuiteDiscount = -1 Or prev.EnsuiteDiscount = curr.EnsuiteDiscount Then
                ensLine &= $"{curr.EnsuiteDiscount}% discount"
            Else
                ensLine &= $"{prev.EnsuiteDiscount}% → {curr.EnsuiteDiscount}% discount"
            End If
            Console.WriteLine(ensLine)

            ' Twin room
            Dim twinLine = $"  Twin: {availability.TwinRoomsAvailable}/{totalTwinRooms} available " &
                           $"({availability.TwinRoomsOccupancyPct:F1}% occupied) - "
            If prev.TwinDiscount = -1 Or prev.TwinDiscount = curr.TwinDiscount Then
                twinLine &= $"{curr.TwinDiscount}% discount"
            Else
                twinLine &= $"{prev.TwinDiscount}% → {curr.TwinDiscount}% discount"
            End If
            Console.WriteLine(twinLine)
            ' -----------------------------------------------

        Next

        Return availabilityByDate
    End Function

    Private Function CalculateDiscountChanges(availabilityData As Dictionary(Of String, RoomAvailability)) As List(Of DiscountChange)
        Dim changes As New List(Of DiscountChange)

        For Each kvp In availabilityData
            Dim dateStr = kvp.Key
            Dim availability = kvp.Value
            Dim DaysAhead = (DateTime.Parse(dateStr) - DateTime.Now).Days

            ' Only apply last-minute discounts for <15 days
            If DaysAhead < 15 Then
                ' Calculate new discounts based on occupancy thresholds
                Dim newDormDiscount = GetDiscountFromOccupancy(availability.DormOccupancyPct)
                Dim newPrivateDiscount = GetDiscountFromOccupancy(availability.PrivateRoomsOccupancyPct)
                Dim newEnsuiteDiscount = GetDiscountFromOccupancy(availability.PrivateEnsuitesOccupancyPct)
                Dim newTwinDiscount = GetDiscountFromOccupancy(availability.TwinRoomsOccupancyPct)

                ' Get previous discounts from storage
                Dim previousDiscounts = GetPreviousDiscounts(dateStr)

                ' Check for changes and create notifications
                If Math.Abs(newDormDiscount - previousDiscounts.DormDiscount) > 0 Then
                    changes.Add(New DiscountChange With {
                        .CheckDate = dateStr,
                        .RoomType = "Dorm Beds",
                        .OldDiscount = previousDiscounts.DormDiscount,
                        .NewDiscount = newDormDiscount,
                        .OccupancyPct = availability.DormOccupancyPct,
                        .AvailableUnits = availability.DormBedsAvailable,
                        .DaysAhead = DaysAhead
                    })
                End If

                If Math.Abs(newPrivateDiscount - previousDiscounts.PrivateDiscount) > 0 Then
                    changes.Add(New DiscountChange With {
                        .CheckDate = dateStr,
                        .RoomType = "Private Rooms (Shared Bath)",
                        .OldDiscount = previousDiscounts.PrivateDiscount,
                        .NewDiscount = newPrivateDiscount,
                        .OccupancyPct = availability.PrivateRoomsOccupancyPct,
                        .AvailableUnits = availability.PrivateRoomsAvailable,
                        .DaysAhead = DaysAhead
                    })
                End If

                If Math.Abs(newEnsuiteDiscount - previousDiscounts.EnsuiteDiscount) > 0 Then
                    changes.Add(New DiscountChange With {
                        .CheckDate = dateStr,
                        .RoomType = "Queen Ensuite",
                        .OldDiscount = previousDiscounts.EnsuiteDiscount,
                        .NewDiscount = newEnsuiteDiscount,
                        .OccupancyPct = availability.PrivateEnsuitesOccupancyPct,
                        .AvailableUnits = availability.PrivateEnsuitesAvailable,
                        .DaysAhead = DaysAhead
                    })
                End If

                If Math.Abs(newTwinDiscount - previousDiscounts.TwinDiscount) > 0 Then
                    changes.Add(New DiscountChange With {
                        .CheckDate = dateStr,
                        .RoomType = "Twin Room (Shared Bath)",
                        .OldDiscount = previousDiscounts.TwinDiscount,
                        .NewDiscount = newTwinDiscount,
                        .OccupancyPct = availability.TwinRoomsOccupancyPct,
                        .AvailableUnits = availability.TwinRoomsAvailable,
                        .DaysAhead = DaysAhead
                    })
                End If

                ' Update stored discounts
                UpdateStoredDiscounts(dateStr, newDormDiscount, newPrivateDiscount, newEnsuiteDiscount, newTwinDiscount)
            End If
        Next

        Return changes
    End Function

    Private Function GetDiscountFromOccupancy(occupancyPct As Double) As Double
        ' Based on occupancy thresholds: <60%, 60-79%, 80-89%, 90%+
        Select Case occupancyPct
            Case >= 90
                Return 0.0 ' No discount - high occupancy
            Case >= 80
                Return 5.0 ' 5% discount - medium-high occupancy
            Case >= 60
                Return 10.0 ' 10% discount - medium occupancy
            Case Else
                Return 15.0 ' Full last-minute discount - low occupancy
        End Select
    End Function

    ' Returns the discounts that WOULD apply today based on the latest occupancy
    Private Function GetCurrentDiscounts(avail As RoomAvailability) As PreviousDiscount
        Return New PreviousDiscount With {
        .DormDiscount = GetDiscountFromOccupancy(avail.DormOccupancyPct),
        .PrivateDiscount = GetDiscountFromOccupancy(avail.PrivateRoomsOccupancyPct),
        .EnsuiteDiscount = GetDiscountFromOccupancy(avail.PrivateEnsuitesOccupancyPct),
        .TwinDiscount = GetDiscountFromOccupancy(avail.TwinRoomsOccupancyPct)
    }
    End Function


    Private Function GetPreviousDiscounts(dateStr As String) As PreviousDiscount
        Try
            If File.Exists(discountStorageFile) Then
                Dim json = File.ReadAllText(discountStorageFile)
                Dim allDiscounts = JsonConvert.DeserializeObject(Of Dictionary(Of String, PreviousDiscount))(json)

                If allDiscounts.ContainsKey(dateStr) Then
                    Return allDiscounts(dateStr)
                End If
            End If
        Catch ex As Exception
            Console.WriteLine($"Error reading previous discounts: {ex.Message}")
        End Try

        ' Return default discounts if no previous data
        Return New PreviousDiscount With {
            .DormDiscount = -1,
            .PrivateDiscount = -1,
            .EnsuiteDiscount = -1,
            .TwinDiscount = -1
        }
    End Function

    Private Sub UpdateStoredDiscounts(dateStr As String, dormDiscount As Double, privateDiscount As Double, ensuiteDiscount As Double, twinDiscount As Double)
        Try
            Dim allDiscounts As New Dictionary(Of String, PreviousDiscount)

            ' Load existing data
            If File.Exists(discountStorageFile) Then
                Dim json = File.ReadAllText(discountStorageFile)
                allDiscounts = JsonConvert.DeserializeObject(Of Dictionary(Of String, PreviousDiscount))(json)
            End If

            allDiscounts(dateStr) = New PreviousDiscount With {
    .DormDiscount = dormDiscount,
    .PrivateDiscount = privateDiscount,
    .EnsuiteDiscount = ensuiteDiscount,
    .TwinDiscount = twinDiscount,
    .LastUpdated = DateTime.Now
}


            ' Clean up old entries (older than 7 days)
            Dim cutoffDate = DateTime.Now.AddDays(-7).ToString("yyyy-MM-dd")
            Dim keysToRemove = allDiscounts.Keys.Where(Function(k) String.Compare(k, cutoffDate) < 0).ToList()
            For Each key In keysToRemove
                allDiscounts.Remove(key)
            Next

            ' Save back to file
            Dim updatedJson = JsonConvert.SerializeObject(allDiscounts, Formatting.Indented)
            File.WriteAllText(discountStorageFile, updatedJson)

        Catch ex As Exception
            Console.WriteLine($"Error updating stored discounts: {ex.Message}")
        End Try
    End Sub


    Public Async Function SendTwilioWhatsAppAsync(changes As List(Of DiscountChange)) As Task
        Try
            Dim message = BuildDiscountChangeMessage(changes)

            ' Twilio API endpoint
            Dim url = $"https://api.twilio.com/2010-04-01/Accounts/{twilioAccountSid}/Messages.json"

            ' Create credentials for Basic Auth
            Dim credentials = Convert.ToBase64String(Encoding.ASCII.GetBytes($"{twilioAccountSid}:{twilioAuthToken}"))

            ' Prepare form data for Twilio
            Dim formData = New List(Of KeyValuePair(Of String, String)) From {
            New KeyValuePair(Of String, String)("From", twilioFromNumber),
            New KeyValuePair(Of String, String)("To", notificationRecipient),
            New KeyValuePair(Of String, String)("Body", message)
        }

            Dim content = New FormUrlEncodedContent(formData)

            ' Set up HTTP client with Basic Auth
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


    Private Function BuildDiscountChangeMessage(changes As List(Of DiscountChange)) As String
        Dim message As New StringBuilder()
        message.AppendLine("🏨 *RED INN COURT - DISCOUNT UPDATE* 🏨")
        message.AppendLine($"📅 {DateTime.Now:yyyy-MM-dd HH:mm}")
        message.AppendLine("")

        For Each change In changes
            message.AppendLine($"📅 *{change.CheckDate}* ({change.DaysAhead} days ahead)")
            message.AppendLine($"🏠 {change.RoomType}")
            message.AppendLine($"📊 Occupancy: {change.OccupancyPct:F1}%")
            message.AppendLine($"🛏️ Available: {change.AvailableUnits} units")

            If change.OldDiscount = -1 Then
                message.AppendLine($"💰 New Discount: *{change.NewDiscount}%*")
            Else
                message.AppendLine($"💰 Discount: {change.OldDiscount}% → *{change.NewDiscount}%*")
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
