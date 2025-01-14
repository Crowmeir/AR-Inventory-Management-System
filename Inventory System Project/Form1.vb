Imports Guna.UI2.WinForms
Imports Microsoft.Data.SqlClient

Public Class Form1

    Private drag As Boolean
    Private dragCursorPoint As Point
    Private dragFormPoint As Point



    Private categories As New List(Of String)
    Private quantities As New List(Of Integer)
    Private totalQuantity As Integer = 0
    Private percentages As New List(Of Integer)


    Private receivedOrders As Integer = 0
    Private pendingOrders As Integer = 0
    Private cancelledOrders As Integer = 0
    Private totalOrders As Integer = 0


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        LoadInventoryData()
        LoadOrderData()
    End Sub


    Private Sub Form1_MouseDown(sender As Object, e As MouseEventArgs) Handles MyBase.MouseDown
        If e.Button = MouseButtons.Left Then
            drag = True
            dragCursorPoint = Cursor.Position
            dragFormPoint = Me.Location
        End If
    End Sub

    Private Sub Form1_MouseMove(sender As Object, e As MouseEventArgs) Handles MyBase.MouseMove
        If drag Then
            Dim dif As Point = Point.Subtract(Cursor.Position, New Size(dragCursorPoint))
            Me.Location = Point.Add(dragFormPoint, New Size(dif))
        End If
    End Sub

    Private Sub Form1_MouseUp(sender As Object, e As MouseEventArgs) Handles MyBase.MouseUp
        If e.Button = MouseButtons.Left Then
            drag = False
        End If
    End Sub
    Public Sub LoadInventoryData()

        categories.Clear()
        quantities.Clear()
        percentages.Clear()
        totalQuantity = 0


        Dim connString As String = "Server=DESKTOP-3LP8SBD\SQLEXPRESS;Database=InventoryProjectDB;User Id=sa;Password=B1Admin;TrustServerCertificate=True;"
        Using conn As New SqlConnection(connString)
            Dim query As String = "SELECT ItemName, Category, Quantity FROM Inventory"
            Dim cmd As New SqlCommand(query, conn)
            conn.Open()
            Dim reader As SqlDataReader = cmd.ExecuteReader()

            While reader.Read()
                Dim category As String = reader("Category").ToString()
                Dim itemName As String = reader("ItemName").ToString()
                Dim quantity As Integer = Convert.ToInt32(reader("Quantity"))

                Dim index As Integer = categories.IndexOf(category)
                If index >= 0 Then
                    quantities(index) += quantity
                Else
                    categories.Add(category)
                    quantities.Add(quantity)
                End If

                totalQuantity += quantity
            End While
        End Using

        For Each quantity In quantities
            percentages.Add(CInt((quantity / totalQuantity) * 100))
        Next


        UpdateProgressBars()


        CheckLowStock()
    End Sub





    Public Sub LoadOrderData()

        receivedOrders = 0
        pendingOrders = 0
        cancelledOrders = 0
        totalOrders = 0

        Dim connString As String = "Server=DESKTOP-3LP8SBD\SQLEXPRESS;Database=InventoryProjectDB;User Id=sa;Password=B1Admin;TrustServerCertificate=True;"
        Using conn As New SqlConnection(connString)
            Dim query As String = "SELECT Status FROM Orders"
            Dim cmd As New SqlCommand(query, conn)
            conn.Open()
            Dim reader As SqlDataReader = cmd.ExecuteReader()

            While reader.Read()
                If reader("Status").ToString() = "Received" Then
                    receivedOrders += 1
                ElseIf reader("Status").ToString() = "Pending" Then
                    pendingOrders += 1
                ElseIf reader("Status").ToString() = "Cancelled" Then
                    cancelledOrders += 1
                End If
                totalOrders += 1
            End While
        End Using


        pendingCountLabel.Text = $"{pendingOrders}"
        receivedCountLabel.Text = $"{receivedOrders}"
        cancelledCountLabel.Text = $"{cancelledOrders}"


        If totalOrders > 0 Then
            Dim receivedPercentage As Integer = CInt((receivedOrders / totalOrders) * 100)
            Dim pendingPercentage As Integer = CInt((pendingOrders / totalOrders) * 100)
            Dim cancelledPercentage As Integer = CInt((cancelledOrders / totalOrders) * 100)


            UpdateOrderProgressBars(receivedPercentage, pendingPercentage, cancelledPercentage)
        End If
    End Sub



    Private Sub UpdateProgressBars()

        If categories.Count >= 1 Then
            Guna2CircleProgressBar1.Value = percentages(0)
            Guna2CircleProgressBar1.ProgressColor = Color.Black
            Guna2CircleProgressBar1.ProgressColor2 = Color.Black
            Guna2CircleProgressBar1.ForeColor = Color.Black
            progressLabel1.Text = $"{percentages(0)}%"
        End If

        If categories.Count >= 2 Then
            Guna2CircleProgressBar2.Value = percentages(1)
            Guna2CircleProgressBar2.ProgressColor = Color.Black
            Guna2CircleProgressBar2.ProgressColor2 = Color.Black
            Guna2CircleProgressBar2.ForeColor = Color.Black
            progressLabel2.Text = $"{percentages(1)}%"
        End If

        If categories.Count >= 3 Then
            Guna2CircleProgressBar3.Value = percentages(2)
            Guna2CircleProgressBar3.ProgressColor = Color.Black
            Guna2CircleProgressBar3.ProgressColor2 = Color.Black
            Guna2CircleProgressBar3.ForeColor = Color.Black
            progressLabel3.Text = $"{percentages(2)}%"
        End If
    End Sub



    Private Sub UpdateOrderProgressBars(receivedPercentage As Integer, pendingPercentage As Integer, cancelledPercentage As Integer)

        receivedProgressBar.Value = receivedPercentage
        receivedProgressBar.ProgressColor = Color.Black
        receivedProgressBar.ProgressColor2 = Color.Black
        receivedProgressBar.ForeColor = Color.Black
        receivedPercentageLabel.Text = $"{receivedPercentage}%"

        pendingProgressbar.Value = pendingPercentage
        pendingProgressbar.ProgressColor = Color.Black
        pendingProgressbar.ProgressColor2 = Color.Black
        pendingProgressbar.ForeColor = Color.Black
        pendingPercentageLabel.Text = $"{pendingPercentage}%"


        cancelledProgressBar.Value = cancelledPercentage
        cancelledProgressBar.ProgressColor = Color.Black
        cancelledProgressBar.ProgressColor2 = Color.Black
        cancelledProgressBar.ForeColor = Color.Black
        cancelledPercentageLabel.Text = $"{cancelledPercentage}%"
    End Sub



    Private Sub btnViewInventory_Click(sender As Object, e As EventArgs) Handles btnViewInventory.Click
        Form2.Show()
        Form2.RefreshInventoryData()
        Me.Hide()
    End Sub

    Private Sub btnAddItem_Click(sender As Object, e As EventArgs) Handles btnAddItem.Click
        Form3.Show()
        Me.Hide()
    End Sub

    Private Sub btnlogo_Click(sender As Object, e As EventArgs) Handles btnlogo.Click

    End Sub

    Private Sub btnManageOrders_Click(sender As Object, e As EventArgs) Handles btnManageOrders.Click
        Form5.Show()
        Me.Hide()
    End Sub

    Private Sub btnReports_Click(sender As Object, e As EventArgs) Handles btnReports.Click
        Hide
        Form4.Show
    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub

    Private Sub pendingCountLabel_Click(sender As Object, e As EventArgs) Handles pendingCountLabel.Click

    End Sub

    Private Sub receivedCountLabel_Click(sender As Object, e As EventArgs) Handles receivedCountLabel.Click

    End Sub

    Private Sub cancelledCountLabel_Click(sender As Object, e As EventArgs) Handles cancelledCountLabel.Click

    End Sub

    Public Sub CheckLowStock()
        Dim lowStockCount As Integer = 0
        Dim lowStockThreshold As Integer = 10
        Dim quantities As New List(Of Integer)()


        Dim connString As String = "Server=DESKTOP-3LP8SBD\SQLEXPRESS;Database=InventoryProjectDB;User Id=sa;Password=B1Admin;TrustServerCertificate=True;"
        Using conn As New SqlConnection(connString)
            Dim query As String = "SELECT Quantity FROM Inventory"
            Dim cmd As New SqlCommand(query, conn)
            conn.Open()
            Dim reader As SqlDataReader = cmd.ExecuteReader()


            While reader.Read()
                quantities.Add(Convert.ToInt32(reader("Quantity")))
            End While
        End Using


        For Each quantity In quantities
            If quantity < lowStockThreshold Then
                lowStockCount += 1
            End If
        Next


        lowStockLabel.Text = $"{lowStockCount}"
    End Sub

    Private Sub Guna2Button4_Click(sender As Object, e As EventArgs) Handles Guna2Button4.Click
        Me.Hide()
        Form7.Show()
    End Sub

    Private Sub cancelledProgressBar_ValueChanged(sender As Object, e As EventArgs) Handles cancelledProgressBar.ValueChanged

    End Sub

    Private Sub cancelledPercentageLabel_Click(sender As Object, e As EventArgs) Handles cancelledPercentageLabel.Click

    End Sub
End Class
