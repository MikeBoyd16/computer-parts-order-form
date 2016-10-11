Option Strict On
Option Explicit On
Public Class SummaryForm

    Private Sub SummaryForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'display the customer's name
        lblName.Text = MainForm.strName
    End Sub

    Private Sub OrderReceipt_SelectedIndexChanged(sender As Object, e As EventArgs) Handles OrderReceipt.SelectedIndexChanged
        'display the order receipt
        OrderReceipt.ToString()
    End Sub

    Private Sub btnSubmitOrder_Click(sender As Object, e As EventArgs) Handles btnSubmitOrder.Click
        NewOrderPrompt()
    End Sub

    Private Sub btnBack_Click(sender As Object, e As EventArgs) Handles btnBack.Click
        'redirects the user back to the coupon form
        Me.Hide()
        MainForm.frmCoupon.Show()
    End Sub

    Sub CalcOrderTotal()
        'declare local variables
        Dim intOrderTotal As Integer

        'calculate the order total
        intOrderTotal = MainForm.intProcessor + MainForm.intOperatingSystem + MainForm.intMemory + MainForm.intHardDrive + MainForm.intVideoCard +
            MainForm.intOpticalDrive + MainForm.intMonitor

        'display the order total and reset
        lblOrderTotal.Text = "$" & CStr(intOrderTotal)
        intOrderTotal = 0
    End Sub

    Sub NewOrderPrompt()
        Dim intUserAnswer As Integer

        ' Displays a message box with the yes and no options.
        intUserAnswer = MsgBox(Prompt:="Success! Would you like to place a new order?", Buttons:=vbYesNo)

        'reset the main form controls for a new order
        If intUserAnswer = vbYes Then
            MainForm.cboProcessor.SelectedIndex = -1
            MainForm.cboOperatingSystem.SelectedIndex = -1
            MainForm.cboMemory.SelectedIndex = -1
            MainForm.cboHardDrive.SelectedIndex = -1
            MainForm.cboVideoCard.SelectedIndex = -1
            MainForm.cboOpticalDrive.SelectedIndex = -1
            MainForm.cboMonitor.SelectedIndex = -1
            MainForm.frmCoupon.txtCouponCode.Clear()
            MainForm.frmCoupon.txtCouponCode.Enabled = True
            MainForm.strName = String.Empty
            While MainForm.strName = String.Empty Or MainForm.strName = " "
                MainForm.strName = InputBox("Customer name: ")
                If MainForm.strName = String.Empty Or MainForm.strName = " " Then
                    MessageBox.Show("You must enter your name. Try again.")
                End If
            End While
            MainForm.Show()
            Me.Hide()
        Else
            'close the program
            Application.Exit()
        End If

    End Sub

    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        Application.Exit()
    End Sub

    Private Sub SummaryForm_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        Application.Exit()
    End Sub
End Class
