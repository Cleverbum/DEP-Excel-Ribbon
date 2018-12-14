Module TimeSpanPrettyString
    Function PrettyString(span As TimeSpan) As String

        Dim d As Integer = span.Duration.Days()
        Dim h As Integer = span.Duration.Hours()
        Dim m As Integer = span.Duration.Minutes()
        Dim s As Integer = span.Duration.Seconds()

        If span.Duration().Days > 0 Then
            PrettyString = d & " Day" & WithS(d) & ", " &
                           h & " Hour" & WithS(h) & ", " &
                           m & " Minute" & WithS(m) & " & " &
                           s & " Second" & WithS(s)

        ElseIf span.Duration().Hours > 0 Then
            PrettyString = h & " Hour" & WithS(h) & ", " &
                           m & " Minute" & WithS(m) & " & " &
                           s & " Second" & WithS(s)

        ElseIf span.Duration().Minutes > 0 Then
            PrettyString = m & " Minute" & WithS(m) & " & " &
                           s & " Second" & WithS(s)

        Else
            PrettyString = s & " Second" & WithS(s)

        End If







    End Function

    Function WithS(num As Integer) As String
        If num > 1 Then
            WithS = "s"
        Else
            WithS = ""
        End If
    End Function
End Module
