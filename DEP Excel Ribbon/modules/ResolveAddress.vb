Imports System.Device.Location
Imports System.Web.Script.Serialization

Module ResolveAddressSync

    Public Sub ResolveAddress()
        Dim watcher As GeoCoordinateWatcher
        watcher = New System.Device.Location.GeoCoordinateWatcher(GeoPositionAccuracy.Default)
        Dim started As Boolean = False
        watcher.MovementThreshold = 1.0     'set to one meter
        started = watcher.TryStart(False, TimeSpan.FromMilliseconds(500))
        Dim i As Integer = 0
        While watcher.Status <> GeoPositionStatus.Ready
            Threading.Thread.Sleep(TimeSpan.FromMilliseconds(10))
            i += 1
            If i > 1000 Then Exit While
        End While

        Dim resolver As CivicAddressResolver = New CivicAddressResolver()
        If started Then
            If Not watcher.Position.Location.IsUnknown Then
                Dim City As String = GetCity(watcher.Position.Location.Latitude, watcher.Position.Location.Longitude)
            Else
                Dim City = "Unknown"

            End If
        Else
            Dim City = "Unknown"

        End If

        'return city
    End Sub

    Private Function GetCity(Lat As Double, Lon As Double) As String

        Dim webClient As New System.Net.WebClient
        Dim url As String = "https://eu1.locationiq.com/v1/reverse.php?key=3a1215da390148&lat=" & Lat & "&lon=" & Lon & "&format=json&zoom=12"
        Dim result As String = webClient.DownloadString(url)
        Dim city As String = Mid(result.ToLower, InStr(result.ToLower, """address"":{"))
        city = Mid(city, InStr(city, "city"":""") + 7)
        city = Mid(city, 1, InStr(city, """,") - 1)
        Return city

    End Function




End Module