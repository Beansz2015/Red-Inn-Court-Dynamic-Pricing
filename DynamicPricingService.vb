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
    Private ReadOnly whatsappToken As String = ConfigurationManager.AppSettings("WhatsAppToken")
    Private ReadOnly whatsappPhoneId As String = ConfigurationManager.AppSettings("WhatsAppPhoneId")
    Private ReadOnly companyGroupId As String = ConfigurationManager.AppSettings("CompanyGroupId")
    Private ReadOnly apiBaseUrl As String = ConfigurationManager.AppSettings("ApiBaseUrl")

    ' Room capacity configuration - based on your actual JSON data
    Private ReadOnly totalDormBeds As Integer = 16 ' 2+4+4+6 = 16 beds total
    Private ReadOnly totalPrivateRooms As Integer = 3
    Private ReadOnly totalEnsuiteRooms As Integer = 2
    Private ReadOnly totalTwinRooms As Integer = 1

    ' Storage file for tracking previous discounts
    Private ReadOnly discountStorageFile As String = "previous_discounts.json"

    Public Async Function RunDynamicPricingCheck() As Task
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

            ' Send WhatsApp notifications if there are changes
            If discountChanges.Any() Then
                Await SendWhatsAppNotificationAsync(discountChanges)
                Console.WriteLine($"Sent WhatsApp notification for {discountChanges.Count} discount changes")
            Else
                Console.WriteLine("No discount changes detected")
            End If

            Console.WriteLine("Dynamic pricing check completed successfully")

        Catch ex As Exception
            Console.WriteLine($"Error in dynamic pricing check: {ex.Message}")
            Console.WriteLine($"Stack trace: {ex.StackTrace}")
            Await SendErrorNotificationAsync(ex.Message)
        End Try
    End Function

    Public Async Function GetRoomAvailabilityAsync() As Task(Of Dictionary(Of String, RoomAvailability))
        Dim availabilityData As New Dictionary(Of String, RoomAvailability)

        Try
            ' LittleHotelier API endpoint - adjust to your actual endpoint
            Dim url = $"{apiBaseUrl}/rates?property_id={propertyId}"

            httpClient.DefaultRequestHeaders.Clear()
            httpClient.DefaultRequestHeaders.Add("Authorization", $"Bearer {apiKey}")
            httpClient.DefaultRequestHeaders.Add("Accept", "application/json")

            Dim response = Await httpClient.GetAsync(url)

            If response.IsSuccessStatusCode Then
                Dim jsonContent = Await response.Content.ReadAsStringAsync()
                Console.WriteLine("Successfully retrieved API data")

                Dim apiDataArray = JsonConvert.DeserializeObject(Of List(Of LittleHotelierResponse))(jsonContent)

                ' Process the first property (Red Inn Court)
                If apiDataArray.Count > 0 Then
                    Dim property = apiDataArray(0)
                    availabilityData = ParseAvailabilityByDate(property)
                    Console.WriteLine($"Parsed availability for {availabilityData.Count} dates")
                End If
            Else
                Dim errorContent = Await response.Content.ReadAsStringAsync()
                Console.WriteLine($"API Error: {response.StatusCode} - {errorContent}")
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
        Dim targetDates As New List(Of String)
        For i As Integer = 1 To 3
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

            Console.WriteLine($"Date: {targetDate}")
            Console.WriteLine($"  Dorms: {availability.DormBedsAvailable}/{totalDormBeds} available ({availability.DormOccupancyPct:F1}% occupied)")
            Console.WriteLine($"  Private: {availability.PrivateRoomsAvailable}/{totalPrivateRooms} available ({availability.PrivateRoomsOccupancyPct:F1}% occupied)")
            Console.WriteLine($"  Ensuite: {availability.PrivateEnsuitesAvailable}/{totalEnsuiteRooms} available ({availability.PrivateEnsuitesOccupancyPct:F1}% occupied)")
            Console.WriteLine($"  Twin: {availability.TwinRoomsAvailable}/{totalTwinRooms} available ({availability.TwinRoomsOccupancyPct:F1}% occupied)")
        Next

        Return availabilityByDate
    End Function

    Private Function CalculateDiscountChanges(availabilityData As Dictionary(Of String, RoomAvailability)) As List(Of DiscountChange)
        Dim changes As New List(Of DiscountChange)

        For Each kvp In availabilityData
            Dim dateStr = kvp.Key
            Dim availability = kvp.Value
            Dim daysAhead = (DateTime.Parse(dateStr) - DateTime.Now).Days

            ' Only apply last-minute discounts for <15 days
            If daysAhead < 15 Then
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
                        .daysAhead = daysAhead
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
                        .daysAhead = daysAhead
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
                        .daysAhead = daysAhead
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
                        .daysAhead = daysAhead
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

            ' Update for this date
            allDiscounts(dateStr) = New PreviousDiscount With {
                .dormDiscount = dormDiscount,
                .privateDiscount = privateDiscount,
                .ensuiteDiscount = ensuiteDiscount,
                .twinDiscount = twinDiscount,
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

    Public Async Function SendWhatsAppNotificationAsync(changes As List(Of DiscountChange)) As Task
        Try
            Dim message = BuildDiscountChangeMessage(changes)

            Dim whatsappUrl = $"https://graph.facebook.com/v18.0/{whatsappPhoneId}/messages"

            Dim payload = New With {
                .messaging_product = "whatsapp",
                .to = companyGroupId,
                .type = "text",
                .text = New With {
                    .body = message
                }
            }

            Dim jsonPayload = JsonConvert.SerializeObject(payload)
            Dim content = New StringContent(jsonPayload, Encoding.UTF8, "application/json")

            httpClient.DefaultRequestHeaders.Clear()
            httpClient.DefaultRequestHeaders.Add("Authorization", $"Bearer {whatsappToken}")

            Dim response = Await httpClient.PostAsync(whatsappUrl, content)

            If response.IsSuccessStatusCode Then
                Console.WriteLine("WhatsApp notification sent successfully")
            Else
                Dim errorContent = Await response.Content.ReadAsStringAsync()
                Console.WriteLine($"WhatsApp API Error: {response.StatusCode} - {errorContent}")
            End If

        Catch ex As Exception
            Console.WriteLine($"Error sending WhatsApp message: {ex.Message}")
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

    Public Async Function SendErrorNotificationAsync(errorMessage As String) As Task
        Try
            Dim message = $"🚨 *DYNAMIC PRICING ERROR* 🚨{vbCrLf}{vbCrLf}❌ {errorMessage}{vbCrLf}{vbCrLf}⏰ {DateTime.Now:yyyy-MM-dd HH:mm}{vbCrLf}{vbCrLf}Please check the application logs for more details."

            Dim whatsappUrl = $"https://graph.facebook.com/v18.0/{whatsappPhoneId}/messages"

            Dim payload = New With {
                .messaging_product = "whatsapp",
                .to = companyGroupId,
                .type = "text",
                .text = New With {
                    .body = message
                }
            }

            Dim jsonPayload = JsonConvert.SerializeObject(payload)
            Dim content = New StringContent(jsonPayload, Encoding.UTF8, "application/json")

            httpClient.DefaultRequestHeaders.Clear()
            httpClient.DefaultRequestHeaders.Add("Authorization", $"Bearer {whatsappToken}")

            Await httpClient.PostAsync(whatsappUrl, content)

        Catch ex As Exception
            Console.WriteLine($"Error sending error notification: {ex.Message}")
        End Try
    End Function

End Class
