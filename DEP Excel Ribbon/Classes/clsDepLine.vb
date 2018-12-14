Public Class ClsDepLine
    Public Property ID As Integer
    Public Property Entity As String
    Public Property Account_Number As Long
    Public Property Company As String
    Public Property DEP As String
    Public Property Post_Code As String
    Public Property Customer_PO As String
    Public Property Sales_ID As Long
    Public Property Order_Date As Date
    Public Property Invoice_Date As Date
    Public Property Order_Type_Desc As String
    Public Property Invoice_ID As Long
    Public Property Item_ID As String
    Public Property Manufacturer_Part_Number As String
    Public Property Item_Name As String
    Public Property Sub_Cat As String
    Public Property Sub_Cat_Description As String
    Public Property Manufacturer_Name As String
    Public Property Units As String
    Public Property POto_Supplier As Long
    Public Property Suppliername As String
    Public Property Account_Manager As String
    Public Property Account_Manager_Email As String
    Public Property POType As String
    Public Property POCreated_Date As Date
    Public Property NDT_Number As Long
    Public Property Action As String
    Public Property Customer_DEP_ID As Long
    Private pSerials() As String


    Public Property Serials() As String()

        Set(ByVal strSerials() As String)
            pSerials = strSerials
        End Set

        Get
            Return pSerials
        End Get
    End Property

    Public Function ToDEPBody() As String
        Dim tmp As String, tmpDEP As String

        On Error Resume Next
        tmpDEP = Split(DEP, " ")(0)
        If tmpDEP = "" Then tmpDEP = DEP
        On Error GoTo 0

        tmp = ":::Short description|||Register DEP:::Client Number||| " &
            Account_Number & ":::Client Name||| " & Company & ":::DEP ID||| " &
            tmpDEP & ":::SO for devices to be added||| " & Sales_ID &
            ":::Description|||Account Manager: " & Account_Manager_Email &
            ", Client PO: " & Customer_PO & ", " & Suppliername & ", "

        tmp = tmp & Join(pSerials, ",") & ", End"
        ToDEPBody = tmp

    End Function
    Public Function ToTicket() As Dictionary(Of String, String)
        ToTicket = New Dictionary(Of String, String)
        Dim tmpDEP As String

        If Order_Type_Desc.ToLower.Contains("return") Then
            ToTicket.Add("Short Description", "Return that may need DEP processing")
        Else
            ToTicket.Add("Short Description", "Register DEP")
        End If

        ToTicket.Add("Client Number", Account_Number)
        ToTicket.Add("Client PO Number", Customer_PO)
        ToTicket.Add("Client Name", Company)
        ToTicket.Add("Units", Units)

        On Error Resume Next
        If DEP Is Nothing Then
            tmpDEP = "TBD"
        Else
            tmpDEP = Split(DEP, " ")(0)
            If tmpDEP = "" Then tmpDEP = DEP
        End If
        On Error GoTo 0

        ToTicket.Add("DEP ID", tmpDEP)
        ToTicket.Add("SO", Sales_ID)
        ToTicket.Add("Long Description", "Account Manager: " & Account_Manager_Email &
            ", Client PO " & Customer_PO & ", " & Suppliername & ", " & Join(pSerials, ",") & ", End")


    End Function

End Class
