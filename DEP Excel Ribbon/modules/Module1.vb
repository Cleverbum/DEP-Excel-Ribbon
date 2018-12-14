Imports OpenQA.Selenium

Module TechDataRegistration

    Function RegisterTechdata(line As ClsDepLine) As Boolean
        Dim wd As Chrome.ChromeDriver
        Dim ndt As New clsNextDeskTicket.ClsNextDeskTicket
        wd = ndt.GiveMeChrome(True)
        wd.Navigate.GoToUrl("https://intouch.techdata.com/Intouch/MiscFE/SSO/ServiceLogin?service=IntouchClient&ContinueUrl=http%3A%2F%2Fintouch.techdata.com%2Fintouch%2FHome.aspx&SessForm=1&Lang=en-GB")


        wd.FindElementByName("customerId").SendKeys("316133")
        wd.FindElementByName("loginUserName").SendKeys("duncanjc")
        wd.FindElementById("password").SendKeys("Fraser123")

        wd.FindElementByClassName("logINbtn").Click()


        wd.FindElementByLinkText("Apple DEP").Click()

        wd.SwitchTo.Frame("ctl00_CPH_iframeCat")
        wd.FindElementById("txtEndCustId").SendKeys(line.Customer_DEP_ID)

        wd.FindElementById("txtEndCustRetNr").SendKeys(line.Customer_PO)

        wd.FindElementById("txtEndCustRetNr").SendKeys(line.Customer_PO)

        wd.FindElementById("txtEndCustName").SendKeys(line.Customer_PO)

        wd.FindElementById("txtMyReference").SendKeys(line.Sales_ID)

        ' generate a list inside the clipboard
        ' as the user to paste it.

        Dim serials As String = ""

        For Each serial In line.Serials
            If serial <> "" Then
                serials = serial & vbCrLf
            End If
        Next

        My.Computer.Clipboard.SetText(serials)

        RegisterTechdata = (MsgBox("serials are now ready to be pasted into the box. Did this work?", vbYesNo) = vbYes)

        wd.Quit()


    End Function
End Module
