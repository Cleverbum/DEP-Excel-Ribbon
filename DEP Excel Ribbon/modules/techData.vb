Imports OpenQA.Selenium

Partial Class CreateNew

    Function DoTDLogin() As Chrome.ChromeDriver

        Dim wd As Chrome.ChromeDriver
        Dim ndt As New clsNextDeskTicket.ClsNextDeskTicket
        wd = ndt.GiveMeChrome(True, True) ' visible and maximised.
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

        wd.Navigate.GoToUrl("https://intouch.techdata.com/intouch/Home.aspx")
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
            If serial IsNot Nothing AndAlso serial <> "" Then
                serials &= serial & vbCrLf
            End If
        Next

        If serials <> "" Then
            Try

                SetClipText(serials)

            Catch
                MsgBox("error setting clipboard")
            End Try

            RegisterTechdata = (MsgBox("The serials are now in the clipboard and ready to be pasted into the box. " & vbCrLf & "Please paste the serial lines in, click submit, and wait for the success/failure message." & vbCrLf & "Did the registration complete successfullly?", vbYesNo) = vbYes)
        Else

            MsgBox("The serial numbers for so " & line.Sales_ID & " are blank, TD registration has been skipped")
            wd.Navigate.GoToUrl("https://intouch.techdata.com/Intouch/MiscFE/SSO/ServiceLogin?service=IntouchClient&ContinueUrl=http%3A%2F%2Fintouch.techdata.com%2Fintouch%2FHome.aspx&SessForm=1&Lang=en-GB")

            Return False
        End If



    End Function
End Class
