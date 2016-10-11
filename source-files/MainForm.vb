Option Strict On
Option Explicit On
Imports System.Threading

Public Class MainForm
    'create new form instances
    Public frmCoupon As New CouponForm
    Public frmSummary As New SummaryForm

    'declare public variables
    Public intProcessor, intOperatingSystem, intMemory, intHardDrive, intVideoCard, intOpticalDrive, intMonitor As Integer
    Public strName As String = String.Empty

    Private Sub MainForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        ComputerPartsSplash.Show()

        'grab the customer's name when the form loads
        While strName = String.Empty Or strName = " "
            strName = InputBox("Customer name: ")

            If strName = String.Empty Or strName = " " Then
                MessageBox.Show("You must enter your name. Try again.")
            End If
        End While
        MessageBox.Show("Select the computer parts you wish to use for your very own personal desktop computer!")

        ComputerPartsSplash.Hide()

        'disable the back button on the main form
        btnBack.Enabled = False
    End Sub

    Private Sub btnNext_Click(sender As Object, e As EventArgs) Handles btnNext.Click
        'validate that a selection was made from each combo box
        If cboProcessor.Text = String.Empty Or cboOperatingSystem.Text = String.Empty Or cboMemory.Text = String.Empty Or cboHardDrive.Text = String.Empty Or
            cboVideoCard.Text = String.Empty Or cboOpticalDrive.Text = String.Empty Or cboMonitor.Text = String.Empty Then
            MessageBox.Show("You must make a selection from each category.")
        Else
            ComputerPartsCalc()
            AddToReceipt()

            'hide the main form and open the coupon form
            Me.Hide()
            frmCoupon.Show()
        End If
    End Sub

    Sub ComputerPartsCalc()
        'place the correct value based on the user's selection for each category
        Select Case cboProcessor.SelectedIndex
            Case 0 : intProcessor = 85
            Case 1 : intProcessor = 150
            Case 2 : intProcessor = 240
            Case 3 : intProcessor = 585
        End Select
        Select Case cboOperatingSystem.SelectedIndex
            Case 0 : intOperatingSystem = 120
            Case 1 : intOperatingSystem = 200
        End Select
        Select Case cboMemory.SelectedIndex
            Case 0 : intMemory = 85
            Case 1 : intMemory = 145
            Case 2 : intMemory = 180
        End Select
        Select Case cboHardDrive.SelectedIndex
            Case 0 : intHardDrive = 55
            Case 1 : intHardDrive = 70
            Case 2 : intHardDrive = 100
        End Select
        Select Case cboVideoCard.SelectedIndex
            Case 0 : intVideoCard = 160
            Case 1 : intVideoCard = 360
            Case 2 : intVideoCard = 580
            Case 3 : intVideoCard = 3000
        End Select
        Select Case cboOpticalDrive.SelectedIndex
            Case 0 : intOpticalDrive = 25
            Case 1 : intOpticalDrive = 55
        End Select
        Select Case cboMonitor.SelectedIndex
            Case 0 : intMonitor = 140
            Case 1 : intMonitor = 235
            Case 2 : intMonitor = 680
            Case 3 : intMonitor = 3300
        End Select
    End Sub

    Sub AddToReceipt()
        'clear the list to avoid repeat information
        frmSummary.OrderReceipt.Items.Clear()

        'add all selected computer parts to the receipt on the summary form
        frmSummary.OrderReceipt.Items.Add(cboProcessor.SelectedItem.ToString)
        frmSummary.OrderReceipt.Items.Add(cboOperatingSystem.SelectedItem.ToString)
        frmSummary.OrderReceipt.Items.Add(cboMemory.SelectedItem.ToString)
        frmSummary.OrderReceipt.Items.Add(cboHardDrive.SelectedItem.ToString)
        frmSummary.OrderReceipt.Items.Add(cboVideoCard.SelectedItem.ToString)
        frmSummary.OrderReceipt.Items.Add(cboOpticalDrive.SelectedItem.ToString)
        frmSummary.OrderReceipt.Items.Add(cboMonitor.SelectedItem.ToString)
    End Sub

    Private Sub ClearToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ClearToolStripMenuItem.Click
        'clear all controls
        cboProcessor.SelectedIndex = -1
        cboOperatingSystem.SelectedIndex = -1
        cboMemory.SelectedIndex = -1
        cboHardDrive.SelectedIndex = -1
        cboVideoCard.SelectedIndex = -1
        cboOpticalDrive.SelectedIndex = -1
        cboMonitor.SelectedIndex = -1
    End Sub

    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        Application.Exit()
    End Sub

    Private Sub MainForm_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        Application.Exit()
    End Sub
End Class
