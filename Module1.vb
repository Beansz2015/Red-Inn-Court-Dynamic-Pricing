Imports System

Module Module1
    Sub Main()
        Try
            Console.WriteLine("=== Red Inn Court Dynamic Pricing Service ===")
            Console.WriteLine($"Started at: {DateTime.Now}")
            Console.WriteLine()

            ' Create and run the service
            Dim service As New DynamicPricingService()
            Dim task = service.RunDynamicPricingCheck()
            task.Wait()

            Console.WriteLine()
            Console.WriteLine($"Completed at: {DateTime.Now}")
            Console.WriteLine("Press any key to exit...")

            ' Only wait for key in interactive mode (not when scheduled)
            If Environment.UserInteractive Then
                Console.ReadKey()
            End If

        Catch ex As Exception
            Console.WriteLine($"Fatal error: {ex.Message}")
            Console.WriteLine($"Stack trace: {ex.StackTrace}")

            ' Log to Windows Event Log for production monitoring
            Try
                Dim log As New System.Diagnostics.EventLog("Application")
                log.Source = "RedInnDynamicPricing"
                log.WriteEntry($"Fatal error: {ex.Message}", System.Diagnostics.EventLogEntryType.Error)
            Catch
                ' Ignore logging errors
            End Try

            Environment.Exit(1) ' Exit with error code
        End Try
    End Sub
End Module
