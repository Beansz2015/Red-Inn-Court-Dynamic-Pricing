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
    Private lastCacheUpdate As DateTime
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
            lastCacheUpdate = DateTime.MinValue

        Catch ex As Exception
            Console.WriteLine($"Failed to initialize Google Sheets service: {ex.Message}")
            Throw
        End Try
    End Sub

    Public Function GetRateData() As Dictionary(Of String, (RegularRate As Double, WalkInRate As Double))
        Try
            ' Return cached data if still valid
            If rateCache.Count > 0 AndAlso DateTime.Now.Subtract(lastCacheUpdate).TotalMinutes < cacheExpiryMinutes Then
                'Console.WriteLine("Using cached Google Sheets rate data")
                Return rateCache
            End If

            ' Fetch fresh data from Google Sheets
            'Console.WriteLine("Fetching fresh rate data from Google Sheets")
            Dim freshData = FetchRateDataFromSheets()

            If freshData.Count > 0 Then
                rateCache = freshData
                lastCacheUpdate = DateTime.Now
                'Console.WriteLine($"Successfully cached {rateCache.Count} rate configurations from Google Sheets")
            Else
                Console.WriteLine("No data retrieved from Google Sheets, using existing cache if available")
            End If

            Return rateCache

        Catch ex As Exception
            Console.WriteLine($"Error getting rate data from Google Sheets: {ex.Message}")
            Return rateCache ' Return existing cache on error
        End Try
    End Function

    Private Function FetchRateDataFromSheets() As Dictionary(Of String, (RegularRate As Double, WalkInRate As Double))
        Dim rateData As New Dictionary(Of String, (Double, Double))

        Try
            Dim range = ConfigurationManager.AppSettings("GoogleSheetRange")
            Dim request = service.Spreadsheets.Values.Get(spreadsheetId, range)
            Dim response = request.Execute()

            If response.Values IsNot Nothing AndAlso response.Values.Count > 1 Then
                'Console.WriteLine($"Found {response.Values.Count} rows in Google Sheets (including header)")

                ' Skip header row
                For i As Integer = 1 To response.Values.Count - 1
                    Dim row = response.Values(i)

                    ' FIXED: Updated to match your actual sheet structure
                    If row.Count >= 5 Then ' Need at least 5 columns: A, B, C, D, E
                        Try
                            Dim configKey = row(0).ToString().Trim()

                            ' FIXED: Read from correct columns based on your sheet structure
                            ' Column D = Regular Rate, Column E = Walk-In Rate
                            Dim regularRate = ParseNumericValue(row(3).ToString()) ' Column D (index 3)
                            Dim walkInRate = ParseNumericValue(row(4).ToString())   ' Column E (index 4)

                            If regularRate.HasValue AndAlso walkInRate.HasValue Then
                                rateData(configKey) = (regularRate.Value, walkInRate.Value)
                                'Console.WriteLine($"✓ Loaded rate: {configKey} = Regular:{regularRate.Value}, Walk-in:{walkInRate.Value}")
                            Else
                                Console.WriteLine($"✗ Skipping row {i + 1}: Invalid numeric values - Regular:'{row(3)}', Walk-in:'{row(4)}'")
                            End If

                        Catch ex As Exception
                            Console.WriteLine($"✗ Error parsing row {i + 1}: {ex.Message}")
                            Console.WriteLine($"   Row data: [{String.Join(", ", row.Select(Function(x) $"'{x}'").ToArray())}]")
                        End Try
                    Else
                        Console.WriteLine($"✗ Skipping row {i + 1}: Not enough columns (has {row.Count}, needs 5)")
                        If row.Count > 0 Then
                            Console.WriteLine($"   Row data: [{String.Join(", ", row.Select(Function(x) $"'{x}'").ToArray())}]")
                        End If
                    End If
                Next
            Else
                Console.WriteLine("No data found in Google Sheets or sheet is empty")
            End If

        Catch ex As Exception
            Console.WriteLine($"Error fetching from Google Sheets: {ex.Message}")
            Throw
        End Try

        'Console.WriteLine($"Successfully parsed {rateData.Count} rate configurations from Google Sheets")
        Return rateData
    End Function

    ' Culture-aware numeric parsing
    Private Function ParseNumericValue(value As String) As Double?
        If String.IsNullOrWhiteSpace(value) Then
            Return Nothing
        End If

        Try
            ' Clean up the value
            Dim cleanValue = value.Trim()

            ' Replace comma with dot for decimal separator
            cleanValue = cleanValue.Replace(",", ".")

            ' Remove any currency symbols or extra spaces
            cleanValue = cleanValue.Replace("RM", "").Replace("$", "").Trim()

            ' Try parsing with invariant culture (uses dot as decimal separator)
            Dim result As Double
            If Double.TryParse(cleanValue, NumberStyles.Float, CultureInfo.InvariantCulture, result) Then
                Return result
            End If

            ' If that fails, try with current culture
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
