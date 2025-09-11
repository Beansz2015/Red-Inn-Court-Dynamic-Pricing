Imports System.Configuration
Imports System.Globalization
Imports System.IO
Imports Google.Apis.Auth.OAuth2
Imports Google.Apis.Services
Imports Google.Apis.Sheets.v4
Imports Google.Apis.Sheets.v4.Data

Public Class GoogleSheetsService
    Private ReadOnly service As SheetsService
    Private ReadOnly operationalSpreadsheetId As String
    Private ReadOnly analyticsSpreadsheetId As String
    Private ReadOnly enableAnalytics As Boolean

    ' Existing caches for operational data
    Private rateCache As Dictionary(Of String, (RegularRate As Double, WalkInRate As Double))
    Private quietPeriodsCache As List(Of QuietPeriod)
    Private lastRateCacheUpdate As DateTime
    Private lastQuietPeriodsCacheUpdate As DateTime
    Private ReadOnly cacheExpiryMinutes As Integer = 5

    Public Sub New()
        Try
            Dim credentialsPath = ConfigurationManager.AppSettings("GoogleCredentialsPath")
            operationalSpreadsheetId = ConfigurationManager.AppSettings("GoogleSheetId")
            analyticsSpreadsheetId = ConfigurationManager.AppSettings("AnalyticsGoogleSheetId")

            Dim enableAnalyticsStr = ConfigurationManager.AppSettings("EnablePriceHistoryLogging")
            enableAnalytics = String.Equals(enableAnalyticsStr, "true", StringComparison.OrdinalIgnoreCase)

            If String.IsNullOrEmpty(credentialsPath) OrElse String.IsNullOrEmpty(operationalSpreadsheetId) Then
                Throw New Exception("Google Sheets configuration is missing")
            End If

            Dim credential = GoogleCredential.FromFile(credentialsPath).CreateScoped(SheetsService.Scope.Spreadsheets)

            service = New SheetsService(New BaseClientService.Initializer() With {
                .HttpClientInitializer = credential,
                .ApplicationName = "Red Inn Court Dynamic Pricing"
            })

            rateCache = New Dictionary(Of String, (Double, Double))
            quietPeriodsCache = New List(Of QuietPeriod)
            lastRateCacheUpdate = DateTime.MinValue
            lastQuietPeriodsCacheUpdate = DateTime.MinValue

            Console.WriteLine("Google Sheets service initialized successfully")
            If enableAnalytics AndAlso Not String.IsNullOrEmpty(analyticsSpreadsheetId) Then
                Console.WriteLine("Analytics logging enabled")
            End If

        Catch ex As Exception
            Console.WriteLine($"Failed to initialize Google Sheets service: {ex.Message}")
            Throw
        End Try
    End Sub

    ' ============ EXISTING OPERATIONAL METHODS (KEEP UNCHANGED) ============
    Public Function GetRateData() As Dictionary(Of String, (RegularRate As Double, WalkInRate As Double))
        Try
            If rateCache.Count > 0 AndAlso DateTime.Now.Subtract(lastRateCacheUpdate).TotalMinutes < cacheExpiryMinutes Then
                Return rateCache
            End If

            Dim freshData = FetchRateDataFromSheets()
            If freshData.Count > 0 Then
                rateCache = freshData
                lastRateCacheUpdate = DateTime.Now
            End If
            Return rateCache

        Catch ex As Exception
            Console.WriteLine($"Error getting rate data from Google Sheets: {ex.Message}")
            Return rateCache
        End Try
    End Function

    Public Function GetQuietPeriods() As List(Of QuietPeriod)
        Try
            If quietPeriodsCache.Count > 0 AndAlso DateTime.Now.Subtract(lastQuietPeriodsCacheUpdate).TotalMinutes < cacheExpiryMinutes Then
                Return quietPeriodsCache
            End If

            Dim freshData = FetchQuietPeriodsFromSheets()
            If freshData.Count > 0 Then
                quietPeriodsCache = freshData
                lastQuietPeriodsCacheUpdate = DateTime.Now
            End If
            Return quietPeriodsCache

        Catch ex As Exception
            Console.WriteLine($"Error getting quiet periods from Google Sheets: {ex.Message}")
            Return quietPeriodsCache
        End Try
    End Function

    ' ============ NEW ANALYTICS METHODS ============
    Public Function EnsureMonthlyPriceHistoryTab() As String
        If Not enableAnalytics OrElse String.IsNullOrEmpty(analyticsSpreadsheetId) Then
            Console.WriteLine("Analytics logging is disabled or not configured")
            Return Nothing
        End If

        Try
            Dim currentMonth = GetBusinessDateTime().ToString("yyyy_MM")
            Dim tabName = $"PriceHistory_{currentMonth}"

            If Not DoesAnalyticsTabExist(tabName) Then
                Console.WriteLine($"Creating new monthly analytics tab: {tabName}")
                CreateAnalyticsPriceHistoryTab(tabName)
            End If

            Return $"{tabName}!A:P"

        Catch ex As Exception
            Console.WriteLine($"Error ensuring monthly analytics tab: {ex.Message}")
            Return Nothing
        End Try
    End Function

    Public Function LogPriceChangeToAnalytics(entry As PriceHistoryEntry) As Boolean
        If Not enableAnalytics OrElse String.IsNullOrEmpty(analyticsSpreadsheetId) Then
            Return False
        End If

        Try
            Dim monthlyRange = EnsureMonthlyPriceHistoryTab()
            If String.IsNullOrEmpty(monthlyRange) Then
                Return False
            End If

            Dim values As New List(Of IList(Of Object))
            Dim row As IList(Of Object) = New List(Of Object) From {
                entry.Timestamp.ToString("yyyy-MM-dd HH:mm:ss"),
                entry.CheckInDate,
                entry.RoomType,
                entry.RoomCategory,
                entry.AvailableUnits,
                entry.TotalCapacity,
                Math.Round(entry.OccupancyRate, 4),
                entry.DaysAhead,
                entry.OldRate,
                entry.NewRate,
                Math.Round(entry.RateChange, 2),
                Math.Round(entry.PercentChange, 4),
                entry.PriceTier,
                entry.DayType,
                entry.BusinessHour,
                entry.ConfigSource
            }

            values.Add(row)

            Dim body As New ValueRange With {.values = values}
            Dim appendRequest = service.Spreadsheets.Values.Append(body, analyticsSpreadsheetId, monthlyRange)
            appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED
            appendRequest.Execute()

            Return True

        Catch ex As Exception
            Console.WriteLine($"Error logging price change to analytics: {ex.Message}")
            Return False
        End Try
    End Function

    ' ============ PRIVATE HELPER METHODS ============
    Private Function DoesAnalyticsTabExist(tabName As String) As Boolean
        Try
            Dim request = service.Spreadsheets.Get(analyticsSpreadsheetId)
            Dim spreadsheet = request.Execute()

            Return spreadsheet.Sheets.Any(Function(sheet) sheet.Properties.Title.Equals(tabName, StringComparison.OrdinalIgnoreCase))

        Catch ex As Exception
            Console.WriteLine($"Error checking analytics tab existence: {ex.Message}")
            Return False
        End Try
    End Function


    Private Sub CreateAnalyticsPriceHistoryTab(tabName As String)
        Try
            Dim addSheetRequest As New Request()
            addSheetRequest.AddSheet = New AddSheetRequest() With {
                .Properties = New SheetProperties() With {.Title = tabName}
            }

            Dim batchUpdateRequest As New BatchUpdateSpreadsheetRequest()
            batchUpdateRequest.Requests = New List(Of Request) From {addSheetRequest}

            service.Spreadsheets.BatchUpdate(batchUpdateRequest, analyticsSpreadsheetId).Execute()
            AddAnalyticsPriceHistoryHeaders(tabName)

            Console.WriteLine($"✓ Successfully created analytics tab: {tabName}")

        Catch ex As Exception
            Console.WriteLine($"Error creating analytics tab {tabName}: {ex.Message}")
            Throw
        End Try
    End Sub

    Private Sub AddAnalyticsPriceHistoryHeaders(tabName As String)
        Try
            Dim headers As IList(Of Object) = New List(Of Object) From {
                "Timestamp", "CheckInDate", "RoomType", "RoomCategory",
                "AvailableUnits", "TotalCapacity", "OccupancyRate", "DaysAhead",
                "OldRate", "NewRate", "RateChange", "PercentChange",
                "PriceTier", "DayType", "BusinessHour", "ConfigSource"
            }

            Dim values As New List(Of IList(Of Object)) From {headers}
            Dim body As New ValueRange With {.values = values}

            service.Spreadsheets.Values.Update(body, analyticsSpreadsheetId, $"{tabName}!A1:P1") _
                .ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW
            service.Spreadsheets.Values.Update(body, analyticsSpreadsheetId, $"{tabName}!A1:P1").Execute()

        Catch ex As Exception
            Console.WriteLine($"Error adding headers to analytics tab {tabName}: {ex.Message}")
        End Try
    End Sub

    Private Function GetBusinessDateTime() As DateTime
        Dim businessTimeZone = TimeZoneInfo.FindSystemTimeZoneById("Singapore Standard Time")
        Return TimeZoneInfo.ConvertTime(DateTime.UtcNow, businessTimeZone)
    End Function

    ' ============ EXISTING PRIVATE METHODS (KEEP UNCHANGED) ============
    Private Function FetchRateDataFromSheets() As Dictionary(Of String, (RegularRate As Double, WalkInRate As Double))
        Dim rateData As New Dictionary(Of String, (Double, Double))
        Try
            Dim range = ConfigurationManager.AppSettings("GoogleSheetRange")
            Dim request = service.Spreadsheets.Values.Get(operationalSpreadsheetId, range)
            Dim response = request.Execute()

            If response.Values IsNot Nothing AndAlso response.Values.Count > 1 Then
                For i As Integer = 1 To response.Values.Count - 1
                    Dim row = response.Values(i)
                    If row.Count >= 5 Then
                        Try
                            Dim configKey = row(0).ToString().Trim()
                            Dim regularRate = ParseNumericValue(row(3).ToString())
                            Dim walkInRate = ParseNumericValue(row(4).ToString())

                            If regularRate.HasValue AndAlso walkInRate.HasValue Then
                                rateData(configKey) = (regularRate.Value, walkInRate.Value)
                            End If
                        Catch ex As Exception
                            Console.WriteLine($"✗ Error parsing rate row {i + 1}: {ex.Message}")
                        End Try
                    End If
                Next
            End If
        Catch ex As Exception
            Console.WriteLine($"Error fetching rate data from Google Sheets: {ex.Message}")
            Throw
        End Try
        Return rateData
    End Function

    ' ============ EXISTING PRIVATE METHODS (COMPLETE IMPLEMENTATIONS) ============
    Private Function FetchQuietPeriodsFromSheets() As List(Of QuietPeriod)
        Dim quietPeriods As New List(Of QuietPeriod)

        Try
            ' Read from QuietPeriods tab
            Dim quietPeriodsRange = ConfigurationManager.AppSettings("GoogleSheetQuietPeriodsRange")
            If String.IsNullOrEmpty(quietPeriodsRange) Then
                quietPeriodsRange = "QuietPeriods!A:E" ' Default range
            End If

            Dim request = service.Spreadsheets.Values.Get(operationalSpreadsheetId, quietPeriodsRange)
            Dim response = request.Execute()

            If response.Values IsNot Nothing AndAlso response.Values.Count > 1 Then
                ' Skip header row
                For i As Integer = 1 To response.Values.Count - 1
                    Dim row = response.Values(i)

                    If row.Count >= 5 Then ' Need: Name, Day, StartTime, EndTime, Enabled
                        Try
                            Dim name = row(0).ToString().Trim()
                            Dim dayOfWeekStr = row(1).ToString().Trim()
                            Dim startTimeStr = row(2).ToString().Trim()
                            Dim endTimeStr = row(3).ToString().Trim()
                            Dim enabledStr = row(4).ToString().Trim()

                            ' Parse day of week
                            Dim dayOfWeek As DayOfWeek? = Nothing
                            If Not String.IsNullOrEmpty(dayOfWeekStr) Then
                                Dim parsedDay As DayOfWeek
                                If [Enum].TryParse(dayOfWeekStr, True, parsedDay) Then
                                    dayOfWeek = parsedDay
                                Else
                                    ' Try parsing common variations
                                    Select Case dayOfWeekStr.ToLower()
                                        Case "daily", "everyday", "all"
                                            dayOfWeek = Nothing ' Special case for daily
                                        Case Else
                                            Console.WriteLine($"✗ Invalid day of week in row {i + 1}: '{dayOfWeekStr}'")
                                            Continue For
                                    End Select
                                End If
                            End If

                            ' Parse times (assuming format like "00:00" or "12:30")
                            Dim startTime As TimeSpan
                            Dim endTime As TimeSpan
                            If Not TimeSpan.TryParse(startTimeStr, startTime) Then
                                Console.WriteLine($"✗ Invalid start time in row {i + 1}: '{startTimeStr}'")
                                Continue For
                            End If
                            If Not TimeSpan.TryParse(endTimeStr, endTime) Then
                                Console.WriteLine($"✗ Invalid end time in row {i + 1}: '{endTimeStr}'")
                                Continue For
                            End If

                            ' Parse enabled flag
                            Dim enabled As Boolean = String.Equals(enabledStr, "true", StringComparison.OrdinalIgnoreCase) OrElse
                                               String.Equals(enabledStr, "yes", StringComparison.OrdinalIgnoreCase) OrElse
                                               String.Equals(enabledStr, "1", StringComparison.OrdinalIgnoreCase)

                            ' Handle daily periods (apply to all days)
                            If dayOfWeek Is Nothing OrElse dayOfWeekStr.ToLower() = "daily" OrElse dayOfWeekStr.ToLower() = "everyday" OrElse dayOfWeekStr.ToLower() = "all" Then
                                For Each day As DayOfWeek In [Enum].GetValues(GetType(DayOfWeek))
                                    quietPeriods.Add(New QuietPeriod With {
                                    .Name = name,
                                    .DayOfWeek = day,
                                    .StartTime = startTime,
                                    .EndTime = endTime,
                                    .Enabled = enabled
                                })
                                Next
                            Else
                                quietPeriods.Add(New QuietPeriod With {
                                .Name = name,
                                .DayOfWeek = dayOfWeek.Value,
                                .StartTime = startTime,
                                .EndTime = endTime,
                                .Enabled = enabled
                            })
                            End If

                        Catch ex As Exception
                            Console.WriteLine($"✗ Error parsing quiet period row {i + 1}: {ex.Message}")
                        End Try
                    Else
                        Console.WriteLine($"✗ Skipping quiet period row {i + 1}: Not enough columns (has {row.Count}, needs 5)")
                    End If
                Next
            Else
                Console.WriteLine("No quiet periods data found in Google Sheets")
            End If

        Catch ex As Exception
            Console.WriteLine($"Error fetching quiet periods from Google Sheets: {ex.Message}")
            ' Don't throw - we can fall back to hardcoded periods
        End Try

        Return quietPeriods
    End Function

    Private Function ParseNumericValue(value As String) As Double?
        If String.IsNullOrWhiteSpace(value) Then
            Return Nothing
        End If

        Try
            Dim cleanValue = value.Trim()
            cleanValue = cleanValue.Replace(",", ".")
            cleanValue = cleanValue.Replace("RM", "").Replace("$", "").Trim()

            Dim result As Double
            If Double.TryParse(cleanValue, NumberStyles.Float, CultureInfo.InvariantCulture, result) Then
                Return result
            End If

            If Double.TryParse(cleanValue, result) Then
                Return result
            End If

            Return Nothing

        Catch ex As Exception
            Return Nothing
        End Try
    End Function


    Public Function IsGoogleSheetsEnabled() As Boolean
        Dim enabledSetting = ConfigurationManager.AppSettings("EnableGoogleSheets")
        Return String.Equals(enabledSetting, "true", StringComparison.OrdinalIgnoreCase)
    End Function
End Class
