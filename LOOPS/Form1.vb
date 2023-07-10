Option Strict On
Option Infer Off
Option Explicit On
Imports System.IO.Packaging
Imports System.Windows.Forms.VisualStyles
'Program: Lab 2---Structure and Sequential Access File
'Purpose: Program for Lab 2 Chapter 9
'Name:Mason Merritt
'Date: Feb 18th,2022

Public Class Form1

    'Page 209 has all the financial methods that can be used in Visual Basic

    'Add 2 Class Level Variables for the Ice Rink judges for the total times a skater was scored and iceTotal for
    'a skater's overall score

    Private judges As Double
    Private icetotal As Double

    'Add Public Variavbles for the GPA Application- Need a gpaTotal(combined total), maleTotal(combined total),
    'FemaleTotal(combined total), maleCount(total times a score was added while male selected),
    'femaleCount(total times a score was added while female selected) and a 
    'gpaCount(total times a score was added overall)

    Private gpaTotal As Double
    Private maleTotal As Double
    Private femaleTotal As Double
    Private maleCount As Double
    Private femaleCount As Double
    Private gpaCount As Double

    'Add Public variables for General Dollar Application- Need a total to accumulate and then access to clear for
    'new order in the Next button

    Private gdTotal As Double

    Private Sub btnAddTestScores_Click(sender As Object, e As EventArgs)

        'Take in User Test Scores add to accumulator and Display average in a label

        Dim score As Double
        Static total As Double
        Static entires As Double
        Dim average As Double
        entires += 1

        Double.TryParse(txtTestScores.Text, score)
        total += score

        average = total / entires

        lblTestScoreAverage.Text = "You scored a " & score.ToString("N2") & vbCrLf &
            "the total scores so far are " & total.ToString("N2") & vbCrLf & " and the average score is " &
            average.ToString("N2")
    End Sub

    Private Sub btnForNext_Click(sender As Object, e As EventArgs) Handles btnForNext.Click

        'Display numbers 1-5 using a For...Next Loop

        For number As Integer = 1 To 5
            lblForNext.Text = lblForNext.Text & number.ToString & "  "
        Next
    End Sub

    Private Sub btnForNextStep_Click(sender As Object, e As EventArgs) Handles btnForNextStep.Click

        'Display the numbers 1 3 5 7 using a For...Next Statement Loop

        For number As Integer = 1 To 7 Step 2
            lblForNext.Text = lblForNext.Text & number.ToString & "  "
        Next
    End Sub

    Private Sub btnYouCanDoIt3_Click(sender As Object, e As EventArgs) Handles btnYouCanDoIt3.Click

        'App should display the numbers 14 - 23 in one label in the second label the
        'total of those numbers should be displayed

        Dim total As Integer
        Dim count As Integer

        For number As Integer = 14 To 23
            count += 1
            total += number
            lblYouCanDoit3.Text = count.ToString
            lblYouCanDoItThree.Text = total.ToString
        Next number
    End Sub

    Private Sub btnGrowthCalc_Click(sender As Object, e As EventArgs) Handles btnGrowthCalc.Click

        'Build App the will take in Current Sales total and a Projected Growth percentage then see how many Years
        'until they reach a $150,000 total.

        Dim growthRate As Double
        Dim totalYears As Integer
        Dim sales As Double
        Dim saleIncrease As Double

        Double.TryParse(txtCurrentSales.Text, sales)
        Double.TryParse(txtGrowthRate.Text, growthRate)

        growthRate /= 100

        Do While sales < 150000
            saleIncrease = sales * growthRate
            sales += saleIncrease
            totalYears += 1
        Loop

        lblA1Results.Text = "It will take " & totalYears.ToString & " years." & vbCrLf & "To reach a sales total of " & sales.ToString("C2")
    End Sub

    Private Sub btnA1CalcRepeat_Click(sender As Object, e As EventArgs) Handles btnA1CalcRepeat.Click

        'Build App the will take in Current Sales total and a Projected Growth percentage then see how many Years
        'until they reach a $150,000 total. Use posttest loop

        Dim growthRate As Double
        Dim totalYears As Integer = 2019
        Dim sales As Double
        Dim saleIncrease As Double

        Double.TryParse(txtCurrentSales2.Text, sales)
        Double.TryParse(txtGrowthRate2.Text, growthRate)

        growthRate /= 100

        Do
            saleIncrease = sales * growthRate
            sales += saleIncrease
            totalYears += 1
            lblA1RepeatResults.Text = lblA1RepeatResults.Text & totalYears.ToString & "          " & sales.ToString("C2") & vbCrLf

        Loop Until sales > 150000
    End Sub

    Private Sub btnA1ForNext_Click(sender As Object, e As EventArgs) Handles btnA1ForNext.Click

        'Build App the will take in Current Sales total and a Projected Growth percentage then see how many Years
        'until they reach a $150,000 total. Use For Next statement

        Dim growthRate As Double
        Dim sales As Double
        Dim saleIncrease As Double

        Double.TryParse(txtGrowthRate3.Text, growthRate)
        Double.TryParse(txtCurrentSales3.Text, sales)

        growthRate /= 100

        For years As Integer = 2023 To 2060
            saleIncrease = sales * growthRate
            sales += saleIncrease
            lstForNext.Items.Add(years.ToString & "          " & sales.ToString("C2") & vbCrLf)
        Next
    End Sub

    Private Sub btnAddDoLoop_Click(sender As Object, e As EventArgs) Handles btnAddDoLoop.Click

        'Build App that will Take in a Number then repeat until 1000 in a ListBox

        Dim number As Double

        Double.TryParse(txtDoLoopAdd.Text, number)

        For count As Double = number To 1000 'could use the STEP method here to skip a certain amount like 5 10 15 20
            lstDoLoopAdd.Items.Add(count)
        Next count

        lstDoLoopAdd.SelectedIndex = 0
    End Sub

    Private Sub btnAddtoListBox_Click(sender As Object, e As EventArgs) Handles btnAddtoListBox.Click, btnAddtoListBox.Click

        'Add List of numbers to List box from Input to `1000

        Dim number As Double

        Double.TryParse(txtCountDisplay.Text, number)

        For count As Double = number To 1000 'could use the STEP method here to skip a certain amount like 5 10 15 20
            lstCountDisplay.Items.Add(count)
        Next count

        lstCountDisplay.SelectedIndex = 0
    End Sub

    Private Sub btnCount_Click(sender As Object, e As EventArgs) Handles btnCount.Click

        'Count How many items are within the listbox 

        lblCountDisplay.Text = lstCountDisplay.Items.Count.ToString
    End Sub

    Private Sub btnUserAdd_Click(sender As Object, e As EventArgs) Handles btnUserAdd.Click

        'User can Input item into List box

        Dim userInput As String = txtUserAdd.Text
        lstUserAdd.Items.Add(userInput.ToString)

        If IsNumeric(userInput) Then
            Dim userNumber As Double
            Double.TryParse(userInput, userNumber)
            Dim userNumber2 As Double
            Double.TryParse(InputBox("To What Number"), userNumber2)
            lstUserAdd.Items.Clear()
            For count As Double = userNumber To userNumber2 Step 5
                lstUserAdd.Items.Add(count)
            Next count
        End If
    End Sub

    Private Sub btn5Loan_Click(sender As Object, e As EventArgs) Handles btn5Loan.Click

        'Calculate the Monthly Payments of a Loan Amount at 5% interest rate for three years

        Dim monthlyPayment As Double
        Dim loanAmount As Double
        Double.TryParse(txt5LoanInput.Text, loanAmount)

        monthlyPayment = -Financial.Pmt(0.05 / 12, 3 * 12, loanAmount)

        lbl5LoanResults.Text = monthlyPayment.ToString("C2")
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        'Load Rates into the Mortgage Rate Listbox

        For rate As Double = 2.0 To 7.0 Step 0.5
            lstMortRate.Items.Add(rate)
        Next rate

        lstMortRate.SelectedIndex = 3

        'Add heading to Savings Textbox

        txtDepositResults.Text = "Rate:" & ControlChars.Tab & "Year:" & ControlChars.Tab & "Balance:" & vbCrLf

        'Add State to the List box Practice

        lstListBoxPractice.Items.Add("California")
        lstListBoxPractice.Items.Add("North Carolina")
        lstListBoxPractice.Items.Add("Virginia")
        lstListBoxPractice.Items.Add("Louisiana")
        lstListBoxPractice.Items.Add("Mississippi")

        'Add percents to the Warehouse Discount App from 5% to 100% in increments of 5

        For percent As Double = 5 To 100 Step 5
            lstWarehousePercents.Items.Add(percent)
        Next percent

        'add numbers 1- 10 to Ice Rink Score List Box

        For number As Integer = 1 To 10
            lstIceRink.Items.Add(number)
        Next number

        'Add Gpa Increments to GPA List box

        Dim gpa As Decimal = 1D

        While gpa < 4.01
            lstGPA.Items.Add(gpa)
            gpa += 0.1D
        End While

        'Add numbers 3-20 to the lstUsefulLife listbox and also ad the heading to the lblDouble and lblSumofYear labels

        Dim years As Double = 3
        While years < 21
            lstUsefulLife.Items.Add(years)
            years += 1
        End While

        lstUsefulLife.SelectedIndex = 2

        lblDoubleDecline.Text = "Year:" & "      " & "Depreciation:" & vbCrLf
        lblSumofYear.Text = "Year:" & "      " & "Depreciation:" & vbCrLf
    End Sub

    Private Sub btnMortCalculate_Click(sender As Object, e As EventArgs) Handles btnMortCalculate.Click

        'App that will take in Loan Amount and use a rate from the listbox to calculate the monthly payment on a loan
        'with terms of 15, 20, 25, and 30 years.

        Dim loanAmount As Double
        Dim rate As Double
        Dim monthlyPayment As Double

        Double.TryParse(txtMortPrincipal.Text, loanAmount)
        Double.TryParse(lstMortRate.SelectedItem.ToString, rate)
        rate /= 100

        'For term As Double = 15 To 30 Step 5
        'monthlyPayment = -Financial.Pmt(rate / 12, term * 12, loanAmount)
        'lblMortResults.Text = lblMortResults.Text & term.ToString & " years:" & "      " & monthlyPayment.ToString("C2") & vbCrLf
        'Next term

        'This code serves as the practice in the Excercise Portion(1st Introductory) It asks to code the same program but using a Do loop

        Dim term As Double = 15

        Do Until term > 30
            monthlyPayment = -Financial.Pmt(rate / 12, term * 12, loanAmount)
            lblMortResults.Text = lblMortResults.Text & term.ToString & " years:" & "      " & monthlyPayment.ToString("C2") & vbCrLf
            term += 5
        Loop

        ' Do While term < 30
        'monthlyPayment = -Financial.Pmt(rate / 12, term * 12, loanAmount)
        'lblMortResults.Text = lblMortResults.Text & term.ToString & " years:" & "      " & monthlyPayment.ToString("C2") & vbCrLf
        'term += 5
        ' Loop
    End Sub
    Private Sub txtMortPrincipal_Click(sender As Object, e As EventArgs) Handles txtMortPrincipal.Click

        'clear label for mortage app

        lblMortResults.Text = ""
    End Sub

    Private Sub lstMortRate_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lstMortRate.SelectedIndexChanged

        'clear label for mortage app

        lblMortResults.Text = ""
    End Sub

    Private Sub btnSavings_Click(sender As Object, e As EventArgs) Handles btnSavings.Click, btnSavings.Click

        'Create an App that will take in a deposit and then calculate the savings over 5 years for 3% to 7% step 1% increments
        'Display results in txtDepositResults. Deposit fomrula is deposit * (1 + rate) ^ year

        Dim deposit As Double
        Dim total As Double
        Dim results As String = ""

        Double.TryParse(txtDeposit.Text, deposit)

        'For rate As Double = 0.03 To 0.08 Step 0.01
        'txtDepositResults.Text = txtDepositResults.Text & rate.ToString("P0") & vbCrLf
        'For year As Integer = 1 To 5
        'total = deposit * (1 + rate) ^ year
        'itResults.Text = txtDepositResults.Text & ControlChars.Tab & year.ToString & ControlChars.Tab & total.ToString("C2") & vbCrLf
        'Next year
        ' Next rate

        'Excercise 3 asks to modify this code by changing the FOR YEAR loop to a do loop

        For rate As Decimal = 0.03D To 0.07D Step 0.01D
            results += rate.ToString("P0") & vbCrLf
            Dim year As Integer = 1
            Do While year <= 5
                total = deposit * (1 + rate) ^ year
                results += ControlChars.Tab & year.ToString & ControlChars.Tab & total.ToString("C2") & vbCrLf
                year += 1
                txtDepositResults.Text = results
            Loop
        Next rate
    End Sub
    Private Sub btnSavingsTwo_Click(sender As Object, e As EventArgs) Handles btnSavingsTwo.Click

        'Create an App that will take in a deposit and then calculate the savings over 5 years for 3% to 7% step 1% increments
        'Display results in txtDepositResults. Deposit fomrula is deposit * (1 + rate) ^ year

        Dim deposit As Double
        Dim total As Double
        Dim results As String = ""
        Dim year As Integer = 1

        Double.TryParse(txtDeposit.Text, deposit)

        While year <= 5
            results += year.ToString & vbCrLf
            Dim rate As Decimal = 0.03D
            While rate <= 0.07
                total = deposit * (1 + rate) ^ year
                results += ControlChars.Tab & rate.ToString("P0") & ControlChars.Tab & total.ToString("C2") & vbCrLf
                rate += 0.01D
                txtDepositResults.Text = results
            End While
            year += 1
        End While
    End Sub

    Private Sub lstListBoxPractice_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lstListBoxPractice.SelectedIndexChanged

        'Make selecteditem from Listbox Practice display State Capital in Label

        Select Case True
            Case lstListBoxPractice.SelectedIndex = 0
                lblListBoxPractice.Text = "Sacramento"
            Case lstListBoxPractice.SelectedIndex = 1
                lblListBoxPractice.Text = "Raleigh"
            Case lstListBoxPractice.SelectedIndex = 2
                lblListBoxPractice.Text = "Richmond"
            Case lstListBoxPractice.SelectedIndex = 3
                lblListBoxPractice.Text = "Baton Rouge"
            Case lstListBoxPractice.SelectedIndex = 4
                lblListBoxPractice.Text = "Jackson"
        End Select
    End Sub

    Private Sub btnMultiply_Click(sender As Object, e As EventArgs) Handles btnMultiply.Click

        'Use a For Loop to Take a a UserNumber and then display that number being multiplied by the numbers 1 - 12

        Dim userNumber As Double
        Dim total As Double

        Double.TryParse(txtMulNumber.Text, userNumber)

        For rate As Double = 1 To 12
            total = userNumber * rate
            txtTimesTable.Text = txtTimesTable.Text & userNumber.ToString & " * " & rate & "   = " & "   " & total.ToString & vbCrLf
        Next rate
    End Sub

    Private Sub btnMultiply2_Click(sender As Object, e As EventArgs) Handles btnMultiply2.Click

        'Use a Do Loop to Take a a UserNumber and then display that number being multiplied by the numbers 1 - 12

        Dim userNumber As Double
        Dim total As Double
        Dim rate As Double = 1

        Double.TryParse(txtMulNumber.Text, userNumber)

        Do While rate <= 12
            total = userNumber * rate
            txtTimesTable.Text = txtTimesTable.Text & userNumber.ToString & " * " & rate &
                "   = " & "   " & total.ToString & vbCrLf
            rate += 1
        Loop
    End Sub

    Private Sub btnWarehouseCalc_Click(sender As Object, e As EventArgs) Handles btnWarehouseCalc.Click

        'Build an App that will take in a price and a discount from a listbox from 5% tp 100% in increments of
        '5 and display the total savings as well as the total price

        Dim userPrice As Double
        Dim userRate As Decimal
        Dim savings As Double

        Double.TryParse(txtWarehousePrice.Text, userPrice)
        Decimal.TryParse(lstWarehousePercents.SelectedItem.ToString, userRate)

        savings = userPrice * (userRate / 100)
        userPrice -= savings

        lblWarehousrResults.Text = " You saved a total of " & savings.ToString("C2") & vbCrLf &
            "The amount due is: " & userPrice.ToString("C2")
    End Sub

    Private Sub btnIceScores_Click(sender As Object, e As EventArgs) Handles btnIceScores.Click

        'Build an App that will Allow a Judge to score a Skater between 1 and 10. Keep the total score for the skater
        ' in the iceTotal label Keep the amount of judges who have scored in the iceJudges label and
        ' calculate the score average and display the total in iceAverage label

        Dim score As Double
        Dim average As Double


        Double.TryParse(lstIceRink.SelectedItem.ToString, score)

        icetotal += score
        judges += 1

        average = icetotal / judges

        lblIceTotal.Text = icetotal.ToString
        lblIceJudges.Text = judges.ToString
        lblIceAverage.Text = average.ToString
    End Sub

    Private Sub btnNextSkater_Click(sender As Object, e As EventArgs) Handles btnNextSkater.Click

        'Clear out labels for next Skater

        lblIceTotal.Text = ""
        lblIceJudges.Text = ""
        lblIceAverage.Text = ""

        judges = 0
        icetotal = 0
    End Sub

    Private Sub btnGPACalc_Click(sender As Object, e As EventArgs) Handles btnGPACalc.Click

        'Build an App that will take in a GPA from list box either Male or Demale and Display the average Gpa
        'for All Student , for Male Student, and For Female Students lblTotalGps, lblMaleGPA, lblFemaleGPA

        Dim score As Double
        Dim gpaAverage As Double
        Dim maleAverage As Double
        Dim femaleAverage As Double

        Double.TryParse(lstGPA.SelectedItem.ToString, score)

        If radMale.Checked = True Then
            maleTotal += score
            gpaTotal += score
            maleCount += 1
            gpaCount += 1
        ElseIf radFemale.Checked = True Then
            femaleTotal += score
            gpaTotal += score
            femaleCount += 1
            gpaCount += 1
        Else
            gpaTotal += score
            gpaCount += 1
        End If

        gpaaverage = gpaTotal / gpaCount
        maleAverage = maleTotal / maleCount
        FemaleAverage = femaleTotal / femaleCount

        lblTotalGPA.Text = gpaAverage.ToString
        lblMaleGPA.Text = maleAverage.ToString
        lblFemaleGPA.Text = femaleAverage.ToString
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        'Clear Accumulators for GPA App and Labels

        gpaTotal = 0
        maleTotal = 0
        femaleTotal = 0
        maleCount = 0
        femaleCount = 0
        gpaCount = 0

        lblTotalGPA.Text = ""
        lblMaleGPA.Text = ""
        lblFemaleGPA.Text = ""
    End Sub

    Private Sub btnUntil_Click_1(sender As Object, e As EventArgs) Handles btnUntil.Click

        '1st PRactice with Loops. Use a Loop to Display the numbers 1-5 in the Label

        lblNumberDisplay.Text = ""

        Dim number As Integer

        Do Until number > 5
            lblNumberDisplay.Text = lblNumberDisplay.Text & number.ToString & "  "
            number += 1
        Loop
    End Sub

    Private Sub btnWhile_Click_1(sender As Object, e As EventArgs) Handles btnWhile.Click

        '1st PRactice with Loops. Use a Loop to Display the numbers 1-5 in the Label

        lblNumberDisplay.Text = ""

        Dim number As Integer = 1

        Do While number <= 5
            lblNumberDisplay.Text = lblNumberDisplay.Text & number.ToString & "  "
            number += 1
        Loop
    End Sub

    Private Sub btnYouDoItOne_Click_1(sender As Object, e As EventArgs) Handles btnYouDoItOne.Click

        'Display the number 1 3 5 and 7 using a pretest loop

        Dim number As Integer = 1

        While number < 8
            lblNumberDisplay.Text = lblNumberDisplay.Text & number.ToString & "  "
            number += 2
        End While
    End Sub

    Private Sub btnPostUntil_Click_1(sender As Object, e As EventArgs) Handles btnPostUntil.Click

        'Dssplay numbers 1-5 unising a post Until Loop
        'Post Loop is one that has the exit condition at the end of the loop instead of the end

        lblPostNumberDisplay.Text = ""

        Dim number As Integer = 1

        Do
            lblPostNumberDisplay.Text = lblPostNumberDisplay.Text & number.ToString & "  "
            number += 1
        Loop While number <= 5
    End Sub

    Private Sub btnPostWhile_Click_1(sender As Object, e As EventArgs) Handles btnPostWhile.Click

        'Display numbers 1-5 unising a post Until Loop
        'Post Loop is one that has the exit condition at the end of the loop instead of the end

        lblPostNumberDisplay.Text = ""

        Dim number As Integer

        Do
            lblPostNumberDisplay.Text = lblPostNumberDisplay.Text & number.ToString & "  "
            number += 1
        Loop Until number > 5
    End Sub

    Private Sub btnYouDoItTwo_Click_1(sender As Object, e As EventArgs) Handles btnYouDoItTwo.Click

        'Display the number 1 3 5 and 7 using a pretest loop

        lblPostNumberDisplay.Text = ""

        Dim number As Integer = 1

        Do
            lblPostNumberDisplay.Text = lblPostNumberDisplay.Text & number.ToString & "  "
            number += 2
        Loop Until number > 7
    End Sub

    Private Sub btnGDCalc_Click(sender As Object, e As EventArgs) Handles btnGDCalc.Click

        'Build an App that will take in a Price add 3% tax to the total and then accumalte for a total Order amount
        'in the label. The textbox should display the prices entered and the Next button should clear the order and
        'Allow for a new order to be built

        Dim itemPrice As Double
        Dim tax As Double = 0.03
        Dim taxAdd As Double
        Double.TryParse(txtGDPrice.Text, itemPrice)

        txtGDPricesEntered.Text = txtGDPricesEntered.Text & itemPrice.ToString & vbCrLf

        taxAdd = itemPrice * tax
        itemPrice += taxAdd
        gdTotal += itemPrice

        lblGDTotal.Text = gdTotal.ToString("C2")
    End Sub

    Private Sub btnGDNext_Click(sender As Object, e As EventArgs) Handles btnGDNext.Click

        'Clear out old data from label and textbox reset gdTotal variable to zero and iniaitate the code to start again

        gdTotal = 0
        txtGDPricesEntered.Text = ""
        lblGDTotal.Text = ""
        txtGDPrice.Text = ""

        Dim itemPrice As Double
        Dim tax As Double = 0.03
        Dim taxAdd As Double
        Double.TryParse(txtGDPrice.Text, itemPrice)

        If itemPrice = 0 Then
            txtGDPricesEntered.Text = ""
        Else
            txtGDPricesEntered.Text = txtGDPricesEntered.Text & itemPrice.ToString & vbCrLf
        End If

        taxAdd = itemPrice * tax
        itemPrice += taxAdd
        gdTotal += itemPrice

        lblGDTotal.Text = gdTotal.ToString("C2")
    End Sub

    Private Sub btnCantonCalc_Click(sender As Object, e As EventArgs) Handles btnCantonCalc.Click

        'Build an app that will take in an Assets Cost and Salvage value as well as its Useful Life. Using both the Declining
        'Balance method and the Sum of Years Digits method calculate the depreciation value for the given period
        'Formula for Declining Balance = Financial.DDB(cost, salavge, life, period)
        'Formula for Sum-of-the-Years method = Financial.SYD(cost, salavge, life, period)

        Dim assetCost As Double
        Dim salavgeValue As Double
        Dim usefulYears As Double
        Dim period As Double
        Dim doubleDecline As Double
        Dim sumofYears As Double
        Dim count As Double = 1
        Dim doubleResults As String
        Dim sumofyearsResults As String

        Double.TryParse(txtAssestCost.Text, assetCost)
        Double.TryParse(txtSalvageValue.Text, salavgeValue)
        Double.TryParse(lstUsefulLife.SelectedItem.ToString, usefulYears)
        Double.TryParse(lstUsefulLife.SelectedItem.ToString, period)

        While count <= usefulYears
            doubleDecline = Financial.DDB(assetCost, salavgeValue, usefulYears, count)
            sumofYears = Financial.SYD(assetCost, salavgeValue, usefulYears, count)
            doubleResults = lblDoubleDecline.Text & count.ToString & "            " & doubleDecline.ToString("N2") & vbCrLf
            sumofyearsResults = lblSumofYear.Text & count.ToString & "            " & sumofYears.ToString("N2") & vbCrLf
            lblDoubleDecline.Text = doubleResults
            lblSumofYear.Text = sumofyearsResults
            count += 1
        End While
    End Sub

    Private Sub btnFibonacci_Click(sender As Object, e As EventArgs) Handles btnFibonacci.Click

        'Display the Fibonacci Sequence to 50th place

        Dim numberOne As Double = 0
        Dim numberTwo As Double = 1
        Dim total As Double
        Dim count As Double = 1

        lstFibonacci.Items.Add("1")

        While count <= 50
            total = numberOne + numberTwo
            numberOne = numberTwo
            numberTwo = total
            lstFibonacci.Items.Add(total)
            count += 1
        End While
    End Sub

    Private Sub btnInsertAt_Click(sender As Object, e As EventArgs) Handles btnInsertAt.Click

        'Allow User to enter a Name and a Number and then insert the item from Listbox at that index

        Dim insertNumber As Integer
        Dim insertName As String = txtInsertName.Text

        Integer.TryParse(txtInsertAtNumber.Text, insertNumber)

        lstAddName.Items.Insert(insertNumber, insertName)
    End Sub

    Private Sub btnRemoveAt_Click(sender As Object, e As EventArgs) Handles btnRemoveAt.Click

        'Allow User to enter a number and then remove the Item from Listbox at that Index

        Dim removeNumber As Integer

        Integer.TryParse(txtRemoveAt.Text, removeNumber)

        lstAddName.Items.RemoveAt(removeNumber)
    End Sub

    Private Sub btnAddName_Click(sender As Object, e As EventArgs) Handles btnAddName.Click

        'Add Name to ListBox AddName

        Dim name As String = txtAddName.Text

        lstAddName.Items.Add(name)
        txtAddName.Text = ""
    End Sub
    Private Sub btnSalaryCalc_Click(sender As Object, e As EventArgs) Handles btnSalaryCalc.Click

        'Build App that will take in an User's Salary and then show the pay increase for raises betwen 1.5% and 3.0 increasing by 0.5% each time
        'can use the formula used for the savings app--- deposit * (1 + rate) ^ year

        Dim salary As Double
        Dim year As Integer = 1
        Dim total As Double
        Dim results As String = ""

        Double.TryParse(txtSalary.Text, salary)

        While year <= 5
            results += vbCrLf & "Year: " & year.ToString & vbCrLf & vbCrLf
            Dim rate As Decimal = 0.015D
            While rate <= 0.03D
                total = salary * (1 + rate) ^ year
                results += rate.ToString("P1") & ControlChars.Tab & total.ToString("C2") & vbCrLf
                txtSalaryResults.Text = results
                rate += 0.005D
            End While
            year += 1
        End While
    End Sub
End Class
