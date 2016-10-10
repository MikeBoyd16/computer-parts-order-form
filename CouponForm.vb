Option Strict On
Option Explicit On
Public Class CouponForm
    'declare class-level variables
    Dim intOrderTotal As Integer
    Dim decSavings As Decimal
    Dim decOrderTotalWithSavings As Decimal

    Private Sub btnSubmit_Click(sender As Object, e As EventArgs) Handles btnNext.Click
        'declare local variables
        Dim intCouponCode As Integer
        Dim intCouponChoice As Integer

        'check that the text box isn't empty
        If txtCouponCode.Text = String.Empty Then
            'check if the user has a coupon code to use
            intCouponChoice = MsgBox(Prompt:="Do you have a coupon code?", Buttons:=vbYesNo)
            If intCouponChoice = vbYes Then
                MessageBox.Show("You must enter your coupon code to continue.")
                txtCouponCode.Focus()
                Exit Sub
            ElseIf intCouponChoice = vbNo Then
                CalcOrderTotal()
                MainForm.frmSummary.lblSavings.Text = "none"
                Me.Hide()
                MainForm.frmSummary.Show()
                Exit Sub
            Else
                MessageBox.Show("You must answer with either a yes or a no.")
            End If
        Else
            'check that the length of the string is 8 characters
            If Len(txtCouponCode.Text) = 8 Then
                'check that the string contents are all numeric
                If IsNumeric(txtCouponCode.Text) Then
                    intCouponCode = CInt(txtCouponCode.Text)
                    CalcDiscount()
                    CalcOrderTotalWithSavings()
                    decSavings = 0
                    If txtCouponCode.Enabled = True Then
                        MessageBox.Show("A 20% discount has been applied to your order.")
                    End If
                    txtCouponCode.Enabled = False
                    'hide the coupon form and show the summary form
                    Me.Hide()
                    MainForm.frmSummary.Show()
                    Exit Sub
                Else
                    MessageBox.Show("Only numbers are allowed. Try again.")
                    txtCouponCode.SelectAll()
                    txtCouponCode.Focus()
                End If
            Else
                MessageBox.Show("The coupon code must be 8 digits. Try again.")
                txtCouponCode.SelectAll()
                txtCouponCode.Focus()
            End If
        End If
    End Sub

    Sub CalcOrderTotal()
        'calculate the order total
        intOrderTotal = MainForm.intProcessor + MainForm.intOperatingSystem + MainForm.intMemory + MainForm.intHardDrive + MainForm.intVideoCard +
            MainForm.intOpticalDrive + MainForm.intMonitor

        'display the order total and reset
        MainForm.frmSummary.lblOrderTotal.Text = intOrderTotal.ToString("c")
        intOrderTotal = 0
    End Sub

    Sub CalcOrderTotalWithSavings()
        'calculate the order total with savings
        decOrderTotalWithSavings = CDec(intOrderTotal) - decSavings

        'display the order total with savings and reset
        MainForm.frmSummary.lblOrderTotal.Text = decOrderTotalWithSavings.ToString("c")
        decOrderTotalWithSavings = 0
    End Sub

    Sub CalcDiscount()
        'declare local variables
        Const decDISCOUNT As Decimal = 0.2D

        'calculate the order total
        intOrderTotal = MainForm.intProcessor + MainForm.intOperatingSystem + MainForm.intMemory + MainForm.intHardDrive + MainForm.intVideoCard +
            MainForm.intOpticalDrive + MainForm.intMonitor

        'calculate the savings
        decSavings = CDec(intOrderTotal) * decDISCOUNT

        'display the savings and reset
        MainForm.frmSummary.lblSavings.Text = decSavings.ToString("c")
    End Sub

    Private Sub btnBack_Click(sender As Object, e As EventArgs) Handles btnBack.Click
        'redirects the user back to the main form
        Me.Hide()
        MainForm.Show()
    End Sub

    Private Sub ClearToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ClearToolStripMenuItem.Click
        txtCouponCode.Clear()
    End Sub

    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        Application.Exit()
    End Sub

    Private Sub CouponForm_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        Application.Exit()
    End Sub

End Class