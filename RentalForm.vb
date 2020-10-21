Option Explicit On
Option Strict On
Option Compare Binary
Public Class RentalForm
    Dim GoodData As Boolean
    Dim totalCustomers As Integer
    Dim totalMileDriven As Double
    Dim Charge As Double
    Dim totalCharges As Double
    Dim milesTotal As Double

    Function Validate() As Boolean
        'Dim Eval As Button
        Dim zipCheck, beginOdeCheck, endOdeCheck, daysCheck As Integer

        If NameTextBox.Text = Nothing Then
            ActiveControl = NameTextBox
            MsgBox("Please fill out all fields before calculating.")
            Exit Function

        ElseIf AddressTextBox.Text = Nothing Then
            ActiveControl = AddressTextBox
            MsgBox("Please fill out all fields before calculating.")
            Exit Function

        ElseIf CityTextBox.Text = Nothing Then
            ActiveControl = CityTextBox
            MsgBox("Please fill out all fields before calculating.")
            Exit Function

        ElseIf StateTextBox.Text = Nothing Then
            ActiveControl = StateTextBox
            MsgBox("Please fill out all fields before calculating.")
            Exit Function


        ElseIf ZipCodeTextBox.Text = Nothing Then
            ActiveControl = ZipCodeTextBox
            MsgBox("Please fill out all fields before calculating.")
            Exit Function

        ElseIf BeginOdometerTextBox.Text = Nothing Then
            ActiveControl = BeginOdometerTextBox
            MsgBox("Please fill out all fields before calculating.")
            Exit Function

        ElseIf EndOdometerTextBox.Text = Nothing Then
            ActiveControl = EndOdometerTextBox
            MsgBox("Please fill out all fields before calculating.")
            Exit Function

        ElseIf DaysTextBox.Text = Nothing Then
            ActiveControl = DaysTextBox
            MsgBox("Please fill out all fields before calculating.")
            Exit Function
        End If

        Try
            zipCheck = CInt(ZipCodeTextBox.Text)
        Catch ex As Exception
            MsgBox("Zipcode has to be a number.")
            ActiveControl = ZipCodeTextBox
            Exit Function
        End Try

        Try
            beginOdeCheck = CInt(BeginOdometerTextBox.Text)
        Catch ex As Exception
            MsgBox("Beginning odometer reading has to be a number.")
            ActiveControl = BeginOdometerTextBox
            Exit Function

        End Try

        Try
            endOdeCheck = CInt(EndOdometerTextBox.Text)
        Catch ex As Exception
            MsgBox("Ending odometer reading has to be a number.")
            ActiveControl = EndOdometerTextBox
            Exit Function
        End Try

        Try
            daysCheck = CInt(DaysTextBox.Text)
        Catch ex As Exception
            MsgBox("Number of days renting has to be a number.")
            ActiveControl = DaysTextBox
            Exit Function
        End Try

        If daysCheck > 45 Or daysCheck <= 0 Then
            MsgBox("Vehicles can only be rented between 1 to 45 days.")
            ActiveControl = DaysTextBox
            Exit Function

        ElseIf endOdeCheck <= beginOdeCheck Then
            MsgBox("Ending odometer can not be less than beginng odemter ")
            BeginOdometerTextBox.Text = Nothing
            EndOdometerTextBox.Text = Nothing
            ActiveControl = BeginOdometerTextBox
            Exit Function
        End If

        GoodData = True
    End Function

    Sub Calculate()
        Dim numDays As Integer
        Dim daysTotalSum As Integer
        Dim milesPrice As Double
        Dim Subtotal As Double
        Dim discountAmount As Double
        Dim AAAdiscount As Double
        Dim seniorDiscount As Double
        Dim roundDiscount As Double
        numDays = CInt(DaysTextBox.Text)
        daysTotalSum = (numDays * 15)
        DayChargeTextBox.Text = ("$" & Str(daysTotalSum))


        If KilometersradioButton.Checked = True Then
            milesTotal = CInt((0.62 * CDbl(EndOdometerTextBox.Text)) - (0.62 * CDbl(BeginOdometerTextBox.Text)))
            TotalMilesTextBox.Text = Str(milesTotal)
        Else
            milesTotal = (CInt(EndOdometerTextBox.Text) - CDbl(BeginOdometerTextBox.Text))
            TotalMilesTextBox.Text = (Str(milesTotal) & " Mi")
        End If

        totalMileDriven += milesTotal


        If milesTotal <= 200 Then
            milesPrice = 0
        ElseIf milesTotal <= 500 Then
            milesPrice = ((milesTotal - 200) * 0.12)
        ElseIf milesTotal > 500 Then
            milesPrice = ((300 * 0.12) + ((milesTotal - 500) * 0.1))

        End If
        MileageChargeTextBox.Text = ("$" & Str(milesPrice))

        Subtotal = milesPrice + daysTotalSum

        If AAAcheckbox.Checked = True Then
            AAAdiscount = Subtotal * 0.05
        Else
            AAAdiscount = 0
        End If

        If Seniorcheckbox.Checked = True Then
            seniorDiscount = Subtotal * 0.03
        Else
            seniorDiscount = 0
        End If

        discountAmount = (AAAdiscount + seniorDiscount)
        roundDiscount = Math.Round(discountAmount, 2, MidpointRounding.AwayFromZero)


        TotalDiscountTextBox.Text = ("$" & Str(roundDiscount))
        Charge = ((Subtotal - roundDiscount))
        TotalChargeTextBox.Text = ("$" & Str(Charge))
        totalCharges += Charge
        totalCustomers += 1
        If totalCustomers <> 0 Then
            SummaryButton.Enabled = True
        End If

    End Sub

    Public Sub CalculateButton_Click(sender As Object, e As EventArgs) Handles CalculateButton.Click

        GoodData = False

        Validate()
        If GoodData = False Then
            Exit Sub
        End If
        Calculate()

    End Sub

    Private Sub ExitButton_Click(sender As Object, e As EventArgs) Handles ExitButton.Click
        'Create message box with YES NO option to exit
        Dim result As MsgBoxResult
        result = MsgBox("Do you really want to exit?", MsgBoxStyle.YesNo)

        If result = 6 Then
            Me.Close()
        Else


        End If
    End Sub

    Private Sub ClearButton_Click(sender As Object, e As EventArgs) Handles ClearButton.Click
        Clear()

    End Sub

    Sub Clear()
        NameTextBox.Text = ""
        AddressTextBox.Text = ""
        CityTextBox.Text = ""
        StateTextBox.Text = ""
        ZipCodeTextBox.Text = ""
        BeginOdometerTextBox.Text = ""
        EndOdometerTextBox.Text = ""
        DaysTextBox.Text = ""
        MilesradioButton.Checked = True
        AAAcheckbox.Checked = False
        Seniorcheckbox.Checked = False
    End Sub

    Private Sub RentalForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        SummaryButton.Enabled = False
    End Sub

    Private Sub SummaryButton_Click(sender As Object, e As EventArgs) Handles SummaryButton.Click
        Clear()
        MsgBox("Customers: " & totalCustomers & vbNewLine & "Total distance driven: " & totalMileDriven & " Miles" & vbNewLine & "Total Charges: $" & totalCharges)
    End Sub
End Class
