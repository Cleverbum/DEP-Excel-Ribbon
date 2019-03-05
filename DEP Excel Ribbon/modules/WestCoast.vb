Imports OpenQA.Selenium


Partial Class CreateNew

    Function DoWCLogin() As Chrome.ChromeDriver

        Dim wd As Chrome.ChromeDriver
        Dim ndt As New clsNextDeskTicket.ClsNextDeskTicket
        wd = ndt.GiveMeChrome(True, True) ' visible and maximised.
        wd.Navigate.GoToUrl("https://www.westcoast.co.uk/login")

        Try
            wd.FindElementByClassName("introjs-skipbutton").Click()
            wd.FindElementByName("loginname").SendKeys("martin.klefas@insight.com")
            wd.FindElementByName("password").SendKeys("Mg3oOjM535")
            Threading.Thread.Sleep(TimeSpan.FromSeconds(4))
            For Each tElement As IWebElement In wd.FindElementsByClassName("btn")
                If tElement.Text.ToLower.Contains("login") Then
                    tElement.Click()
                    Threading.Thread.Sleep(TimeSpan.FromSeconds(4))
                    Exit For
                End If
            Next

            wd.FindElementByClassName("introjs-skipbutton").Click()

        Catch ex As Exception
            UpdateDebugMessage("Westcoast Login Failed")
        End Try



        Return wd

    End Function

    Function DoOneWC_DEP(line As ClsDepLine, wd As Chrome.ChromeDriver) As Boolean
        Try
            wd.Navigate.GoToUrl("https://www.westcoast.co.uk/StartApp.aspx?name=dep&update=0&delete=0&mixbasket=0&pricecheck=1")

            wd.Navigate.GoToUrl("http://dep.westcoast.co.uk/createOrder.php")


            wd.FindElementByName("po").SendKeys(line.Customer_PO)
            wd.FindElementByName("enduserid").SendKeys(line.Customer_DEP_ID)

            Dim serials As String = ""

            For Each serial In line.Serials
                If serial IsNot Nothing AndAlso serial <> "" Then
                    serials &= serial & vbCrLf
                End If
            Next

            If serials <> "" Then
                Dim serialTypeAreas = wd.FindElementsByName("serialupload")
                For Each SerialArea As IWebElement In serialTypeAreas
                    If SerialArea.TagName = "textarea" Then
                        SerialArea.SendKeys(serials)
                        Exit For
                    End If
                Next

            Else
                Return False
            End If

            wd.FindElementByName("uploadsubmitlist").Click()

            Threading.Thread.Sleep(TimeSpan.FromSeconds(4)) ' it seems to prefer if we go slow

            wd.FindElementByName("po").SendKeys(line.Customer_PO)
            wd.FindElementByName("enduserid").SendKeys(line.Customer_DEP_ID)

            Threading.Thread.Sleep(TimeSpan.FromSeconds(6)) ' it seems to prefer if we go slow

            wd.FindElementByClassName("orddetail").Click()

            Threading.Thread.Sleep(TimeSpan.FromSeconds(4)) ' it seems to prefer if we go slow

            Dim OrderDetailButtons = wd.FindElementsByClassName("orddetail")
            For Each OrderButton As IWebElement In OrderDetailButtons
                If OrderButton.Text = "Submit to Apple" Then
                    OrderButton.Click()
                    Exit For
                End If
            Next

            Threading.Thread.Sleep(TimeSpan.FromSeconds(4)) ' it seems to prefer if we go slow
            wd.FindElementByName("reselleremail").SendKeys("jenni.ings@insight.com")
            wd.FindElementByName("poordswap").Click()
            Dim SubmitButtons = wd.FindElementsByClassName("orddetail")
            For Each SubmitButton As IWebElement In SubmitButtons
                If SubmitButton.Text = "" Then
                    SubmitButton.Click()
                    Exit For
                End If
            Next


            Return wd.FindElementByClassName("error").Text.ToLower.Contains("have been submitted")
        Catch
            Return False
        End Try
    End Function

End Class
