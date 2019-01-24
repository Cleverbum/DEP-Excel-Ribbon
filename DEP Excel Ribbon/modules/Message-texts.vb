Module Message_texts
    Public Const distiEmailMessage As String = "As attached, an email was sent to distribution to ask them to register these devices to the appropriate DEP ID. They will confirm by email when this is done, at which point this ticket will be updated again."
    Public Const OnlyMessage As String = "DEP Team: There is an 'Only' condition in this customer's registration preferences, and so this registration will need to be completed manually. Thanks."
    Public Const FakeMessage As String = "DEP Team: It seems that some of the serial numbers recorded in iCare do not match normal Apple patterns - please can you investigate this prior to submitting these for DEP."
    Public Const DepQuestion As String = "Hi, this shipped yesterday, would the client like this to be added to DEP? If so, please provide their DEP ID.  Would the customer also like all Apple devices adding to DEP when shipped? Thanks"
    Public Const NoEmailSent As String = "No mail was sent for this as the distributor is %SupplierName%. DEP Team: Please complete their process manually."
    Public Const tdFail As String = "Techdata registration failed via the assisted tool: DEP Team, please complete this manually for this order."
    Public Const tdSuccess As String = "Techdata registration was submitted successfully using the assisted tool. DEP Team: Please close this ticket once confirmation has been recieved from TechData"
    Public Const wcFail As String = "Westcoast registration failed via the assisted tool: DEP Team, please complete this manually for this order."
    Public Const wcSuccess As String = "Westcoast registration was submitted successfully using the assisted tool. DEP Team: Please close this ticket once confirmation has been recieved from TechData"


    Public Const FailedRegMsg As String = "An automated check of the distributor website suggests that the registration process failed for this order. DEP Team: Please can you manually verify that there has been an error and take steps to resolve it."

    Public Const SuccessfulRegMsg As String = "An automated check of the distribor's website confirms that the registration process for this order was successful. This ticket will now be closed."

    Public Const CloseMsg As String = "All of the tasks required for this order have now been completed, and as such this ticket has been closed. Please let us know if we are mistaken and something else needs doing."

End Module
