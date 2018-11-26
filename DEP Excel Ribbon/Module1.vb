Imports OpenQA.Selenium

Module TechDataRegistration

    Function RegisterTechdata(line As ClsDepLine) As Boolean
        Dim wd As Chrome.ChromeDriver
        Dim ndt As New clsNextDeskTicket.ClsNextDeskTicket
        wd = ndt.GiveMeChrome(True)
        wd.Navigate.GoToUrl("http://uk.techdata.com")

        wd.FindElementByClassName("login").Click()

        wd.FindElementByName("customerId").SendKeys("316133")
        wd.FindElementByName("loginUserName").SendKeys("duncanjc")
        wd.FindElementById("password").SendKeys("Fraser123")

        wd.FindElementByClassName("logINbtn").Click()


        wd.FindElementByLinkText("Apple DEP").Click()

        wd.SwitchTo.Frame("ctl00_CPH_iframeCat")
        wd.FindElementById("txtEndCustId").SendKeys("1") 'line.Customer_DEP_ID)

        wd.FindElementById("txtEndCustRetNr").SendKeys("2") 'line.Customer_PO)

        wd.FindElementById("txtEndCustRetNr").SendKeys("3") 'line.Customer_PO)

        wd.FindElementById("txtEndCustName").SendKeys("4") 'line.Customer_PO)

        wd.FindElementById("txtMyReference").SendKeys("5") 'line.Sales_ID)

        '  wd.FindElementByXPath("//*[@id='divEnroll']/div[1]/div[1]/div/div[1]/table/tbody/tr/td[4]").SendKeys("123")
        ' generate a list inside the clipboard
        ' as the user to paste it.

        wd.Quit()

        RegisterTechdata = False


    End Function
End Module
