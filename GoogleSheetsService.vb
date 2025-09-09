Imports Google.Apis.Auth.OAuth2
Imports Google.Apis.Sheets.v4
Imports Google.Apis.Services
Imports System.IO
Imports System.Configuration
Imports System.Globalization

Public Class GoogleSheetsService
    Private ReadOnly service As SheetsService
    Private ReadOnly spreadsheetId As String
    Private rateCache As Dictionary(Of String, (RegularRate As Double, WalkInRate As Double))
    Private quietPeriodsCache As List(Of QuietPeriod)
    Private lastRateCacheUpdate As DateTime
    Private lastQuietPeriodsCacheUpdate As DateTime
    Private ReadOnly cacheExpiryMinutes As Integer = 5

    Public Sub New()
        Try
            Dim credentialsPath = ConfigurationManager.AppSettings("GoogleCredentialsPath")
            spreadsheetId = ConfigurationManager.AppSettings("GoogleSheetId")

            If String.IsNullOrEmpty(credentialsPath) OrElse String.IsNullOrEmpty(spreadsheetId) Then
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

        Catch ex As Exception
            Console.WriteLine($"Failed to initialize Google Sheets service: {ex.Message}")
            Throw
        End Try
    End Sub

    Public Function GetRateData() As Dictionary(Of String, (RegularRate As Double, WalkInRate As Double))
        Try
            ' Return cached data if still valid
            If rateCache.Count > 0 AndAlso DateTime.Now.Subtract(lastRateCacheUpdate).TotalMinutes < cacheExpiryMinutes Then
                Return rateCache
            End If

            ' Fetch fresh data from Google Sheets
            Dim freshData = FetchRateDataFromSheets()

            If freshData.Count > 0 Then
                rateCache = freshData
                lastRateCacheUpdate = DateTime.Now
            Else
                Console.WriteLine("No rate data retrieved from Google Sheets, using existing cache if available")
            End If

            Return rateCache

        Catch ex As Exception
            Console.WriteLine($"Error getting rate data from Google Sheets: {ex.Message}")
            Return rateCache ' Return existing cache on error
        End Try
    End Function

    ' NEW: Get quiet periods configuration from Google Sheets
    Public Function GetQuietPeriods() As List(Of QuietPeriod)
        Try
            ' Return cached data if still valid
            If quietPeriodsCache.Count > 0 AndAlso DateTime.Now.Subtract(lastQuietPeriodsCacheUpdate).TotalMinutes < cacheExpiryMinutes Then
                Return quietPeriodsCache
            End If

            ' Fetch fresh data from Google Sheets
            Dim freshData = FetchQuietPeriodsFromSheets()

            If freshData.Count > 0 Then
                quietPeriodsCache = freshData
                lastQuietPeriodsCacheUpdate = DateTime.Now
                'Console.WriteLine($"📋 Loaded {quietPeriodsCache.Count} quiet periods from Google Sheets")
            Else
                Console.WriteLine("No quiet periods data retrieved from Google Sheets, using existing cache if available")
            End If

            Return quietPeriodsCache

        Catch ex As Exception
            Console.WriteLine($"Error getting quiet periods from Google Sheets: {ex.Message}")
            Return quietPeriodsCache ' Return existing cache on error
        End Try
    End Function

    Private Function FetchRateDataFromSheets() As Dictionary(Of String, (RegularRate As Double, WalkInRate As Double))
        Dim rateData As New Dictionary(Of String, (Double, Double))

        Try
            Dim range = ConfigurationManager.AppSettings("GoogleSheetRange")
            Dim request = service.Spreadsheets.Values.Get(spreadsheetId, range)
            Dim response = request.Execute()

            If response.Values IsNot Nothing AndAlso response.Values.Count > 1 Then
                ' Skip header row
                For i As Integer = 1 To response.Values.Count - 1
                    Dim row = response.Values(i)

                    If row.Count >= 5 Then ' Need at least 5 columns: A, B, C, D, E
                        Try
                            Dim configKey = row(0).ToString().Trim()
                            Dim regularRate = ParseNumericValue(row(3).ToString()) ' Column D
                            Dim walkInRate = ParseNumericValue(row(4).ToString())   ' Column E

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

    ' NEW: Fetch quiet periods from Google Sheets
    Private Function FetchQuietPeriodsFromSheets() As List(Of QuietPeriod)
        Dim quietPeriods As New List(Of QuietPeriod)

        Try
            ' Read from QuietPeriods tab
            Dim quietPeriodsRange = ConfigurationManager.AppSettings("GoogleSheetQuietPeriodsRange")
            If String.IsNullOrEmpty(quietPeriodsRange) Then
                quietPeriodsRange = "QuietPeriods!A:E" ' Default range
            End If

            Dim request = service.Spreadsheets.Values.Get(spreadsheetId, quietPeriodsRange)
            Dim response = request.Execute()

            If response.Values IsNot Nothing AndAlso response.Values.Count > 1 Then
                'Console.WriteLine($"📋 Found {response.Values.Count - 1} quiet period rows in Google Sheets")

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
                                        .name = name,
                                        .dayOfWeek = day,
                                        .startTime = startTime,
                                        .endTime = endTime,
                                        .enabled = enabled
                                    })
                                Next
                            Else
                                quietPeriods.Add(New QuietPeriod With {
                                    .name = name,
                                    .dayOfWeek = dayOfWeek.Value,
                                    .startTime = startTime,
                                    .endTime = endTime,
                                    .enabled = enabled
                                })
                            End If

                            'Console.WriteLine($"✓ Loaded quiet period: {name} - {dayOfWeekStr} {startTime:hh\:mm}-{endTime:hh\:mm} (Enabled: {enabled})")

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

    ' Culture-aware numeric parsing (unchanged)
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
