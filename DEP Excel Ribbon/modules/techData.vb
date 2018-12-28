Imports OpenQA.Selenium

Partial Class CreateNew

    Function DoTDLogin() As Chrome.ChromeDriver

        Dim wd As Chrome.ChromeDriver
        Dim ndt As New clsNextDeskTicket.ClsNextDeskTicket
        wd = ndt.GiveMeChrome(True)
        wd.Navigate.GoToUrl("https://intouch.techdata.com/Intouch/MiscFE/SSO/ServiceLogin?service=IntouchClient&ContinueUrl=http%3A%2F%2Fintouch.techdata.com%2Fintouch%2FHome.aspx&SessForm=1&Lang=en-GB")

        Try

            wd.FindElementByName("customerId").SendKeys("316133")
            wd.FindElementByName("loginUserName").SendKeys("duncanjc")
            wd.FindElementById("password").SendKeys("Fraser123")
            wd.FindElementByClassName("logINbtn").Click()

            Return wd
        Catch
            MsgBox("Could not login, please do so manually", vbCritical)
            Return wd

        End Try


    End Function


    Function RegisterTechdata(line As ClsDepLine, wd As Chrome.ChromeDriver) As Boolean

        wd.FindElementByLinkText("Apple DEP").Click()

        wd.SwitchTo.Frame("ctl00_CPH_iframeCat")
        wd.FindElementById("txtEndCustId").SendKeys(line.Customer_DEP_ID)

        wd.FindElementById("txtEndCustRetNr").SendKeys(line.Customer_PO)



        wd.FindElementById("txtEndCustName").SendKeys(line.Company)

        wd.FindElementById("txtMyReference").SendKeys(line.Sales_ID)

        ' generate a list inside the clipboard
        ' as the user to paste it.

        Dim serials As String = ""

        For Each serial In line.Serials
            If serial <> "" Then
                serials &= serial & vbCrLf
            End If
        Next

        Try
            SetClipText(serials)
        Catch
            MsgBox("well that didn't work")
        End Try

        RegisterTechdata = (MsgBox("The serials are now in the clipboard and ready to be pasted into the box. Did this work?", vbYesNo) = vbYes)




    End Function
End Class
