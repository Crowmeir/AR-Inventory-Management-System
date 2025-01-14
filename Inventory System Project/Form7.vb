Imports System.Data.SqlClient
Imports Microsoft.Data.SqlClient

Public Class Form7

    Private TotalSale As Decimal = 0
    Private AvailableStock As Integer = 0
    Private CurrentPrice As Decimal = 0

    Private drag As Boolean
    Private dragCursorPoint As Point
    Private dragFormPoint As Point

    Private Sub Form7_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        If DataGridViewSales.Columns.Count = 0 Then
            DataGridViewSales.Columns.Add("ItemID", "Item ID")
            DataGridViewSales.Columns.Add("ItemName", "Item Name")
            DataGridViewSales.Columns.Add("QuantitySold", "Quantity Sold")
            DataGridViewSales.Columns.Add("Price", "Price")
            DataGridViewSales.Columns.Add("TotalPrice", "Total Price")
        End If

        'Pang populate sa datagrid' 
        Dim query As String = "SELECT ItemID, ItemName FROM Inventory"
        Dim dt As DataTable = GetData(query)
        cmbItemName.DataSource = dt
        cmbItemName.DisplayMember = "ItemName"
        cmbItemName.ValueMember = "ItemID"
    End Sub

    Private Sub Form7_MouseDown(sender As Object, e As MouseEventArgs) Handles MyBase.MouseDown
        If e.Button = MouseButtons.Left Then
            drag = True
            dragCursorPoint = Cursor.Position
            dragFormPoint = Me.Location
        End If
    End Sub

    Private Sub Form7_MouseMove(sender As Object, e As MouseEventArgs) Handles MyBase.MouseMove
        If drag Then
            Dim dif As Point = Point.Subtract(Cursor.Position, New Size(dragCursorPoint))
            Me.Location = Point.Add(dragFormPoint, New Size(dif))
        End If
    End Sub

    Private Sub Form7_MouseUp(sender As Object, e As MouseEventArgs) Handles MyBase.MouseUp
        If e.Button = MouseButtons.Left Then
            drag = False
        End If
    End Sub


    Private Function GetData(query As String) As DataTable
        Dim dt As New DataTable()
        Dim connectionString As String = "Server=DESKTOP-3LP8SBD\SQLEXPRESS;Database=InventoryProjectDB;User Id=sa;Password=B1Admin;TrustServerCertificate=True;"

        Using conn As New SqlConnection(connectionString)
            Using cmd As New SqlCommand(query, conn)
                conn.Open()
                Dim adapter As New SqlDataAdapter(cmd)
                adapter.Fill(dt)
            End Using
        End Using

        Return dt
    End Function

    Private Sub ExecuteQuery(query As String)
        Dim connectionString As String = "Server=DESKTOP-3LP8SBD\SQLEXPRESS;Database=InventoryProjectDB;User Id=sa;Password=B1Admin;TrustServerCertificate=True;"

        Using conn As New SqlConnection(connectionString)
            Using cmd As New SqlCommand(query, conn)
                conn.Open()
                cmd.ExecuteNonQuery()
            End Using
        End Using
    End Sub


    Private Sub cmbItemName_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbItemName.SelectedIndexChanged

        If cmbItemName.SelectedValue IsNot Nothing AndAlso TypeOf cmbItemName.SelectedValue IsNot DBNull Then
            Dim itemId As Integer
            If Integer.TryParse(cmbItemName.SelectedValue.ToString(), itemId) Then
                ' Kukuha ng price sa inventory table'
                Dim query As String = $"SELECT Price, Quantity FROM Inventory WHERE ItemID = {itemId}"
                Dim dt As DataTable = GetData(query)

                If dt.Rows.Count > 0 Then
                    CurrentPrice = Convert.ToDecimal(dt.Rows(0)("Price"))
                    lblPrice.Text = $"Price: {CurrentPrice:C}"
                    lblStockAvailable.Text = $"Stock Available: {dt.Rows(0)("Quantity")}"
                    AvailableStock = Convert.ToInt32(dt.Rows(0)("Quantity"))
                End If
            End If
        End If
    End Sub


    Private Sub btnAddToSale_Click(sender As Object, e As EventArgs) Handles btnAddToSale.Click
        Dim quantitySold As Integer
        If Integer.TryParse(txtQuantitySold.Text, quantitySold) AndAlso quantitySold > 0 Then
            If quantitySold <= AvailableStock Then
                Dim itemId As Integer = Convert.ToInt32(cmbItemName.SelectedValue)
                Dim itemName As String = cmbItemName.Text
                Dim totalPrice As Decimal = CurrentPrice * quantitySold

                DataGridViewSales.Rows.Add(itemId, itemName, quantitySold, CurrentPrice, totalPrice)

                TotalSale += totalPrice
                lblTotal.Text = $"Total Sale: {TotalSale:C}"

                AvailableStock -= quantitySold
                lblStockAvailable.Text = $"Stock Available: {AvailableStock}"
            Else
                MessageBox.Show("Not enough stock available.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        Else
            MessageBox.Show("Please enter a valid quantity.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub


    Private Sub btnRemoveItem_Click(sender As Object, e As EventArgs) Handles btnRemoveItem.Click
        If DataGridViewSales.SelectedRows.Count > 0 Then
            Dim selectedRow As DataGridViewRow = DataGridViewSales.SelectedRows(0)
            Dim totalPrice As Decimal = Convert.ToDecimal(selectedRow.Cells("TotalPrice").Value)
            Dim quantity As Integer = Convert.ToInt32(selectedRow.Cells("QuantitySold").Value)


            TotalSale -= totalPrice
            lblTotal.Text = $"Total Sale: {TotalSale:C}"

            AvailableStock += quantity
            lblStockAvailable.Text = $"Stock Available: {AvailableStock}"

            DataGridViewSales.Rows.Remove(selectedRow)
        Else
            MessageBox.Show("Please select an item to remove.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub


    Private Sub btnFinalizeSale_Click(sender As Object, e As EventArgs) Handles btnFinalizeSale.Click
        For Each row As DataGridViewRow In DataGridViewSales.Rows

            If row.IsNewRow Then Continue For

            Dim itemId As Integer = Convert.ToInt32(row.Cells("ItemID").Value)
            Dim quantitySold As Integer = Convert.ToInt32(row.Cells("QuantitySold").Value)
            Dim totalPrice As Decimal = Convert.ToDecimal(row.Cells("TotalPrice").Value)


            If itemId <= 0 Then
                MessageBox.Show($"Invalid ItemID: {itemId}. Please check the inventory.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If

            Dim checkQuery As String = $"SELECT COUNT(*) FROM Inventory WHERE ItemID = {itemId}"
            Dim count As Integer = Convert.ToInt32(GetData(checkQuery).Rows(0)(0))

            If count > 0 Then

                Dim query As String = $"INSERT INTO Transactions (ItemID, QuantitySold, TotalPrice) VALUES ({itemId}, {quantitySold}, {totalPrice})"
                ExecuteQuery(query)

                Dim updateQuery As String = $"UPDATE Inventory SET Quantity = Quantity - {quantitySold} WHERE ItemID = {itemId}"
                ExecuteQuery(updateQuery)
            Else
                MessageBox.Show($"Item with ID {itemId} does not exist in the Inventory.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If
        Next

        MessageBox.Show("Sale finalized successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)

        DataGridViewSales.Rows.Clear()
        lblTotal.Text = "Total Sale: $0.00"
        TotalSale = 0
    End Sub




    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        If MessageBox.Show("Are you sure you want to cancel the sale?", "Cancel Sale", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) = DialogResult.Yes Then
            DataGridViewSales.Rows.Clear()
            lblTotal.Text = "Total Sale: $0.00"
            TotalSale = 0
            lblStockAvailable.Text = "Stock Available: N/A"
        End If
    End Sub

    Private Sub Guna2Button1_Click(sender As Object, e As EventArgs) Handles Guna2Button1.Click
        Me.Close()
        Form1.Show()
    End Sub
End Class
