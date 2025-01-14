Imports Microsoft.Data.SqlClient

Public Class Form5
    ' Connection string for the database
    Private connectionString As String = "Server=DESKTOP-3LP8SBD\SQLEXPRESS;Database=InventoryProjectDB;User Id=sa;Password=B1Admin;TrustServerCertificate=True;"

    Private drag As Boolean
    Private dragCursorPoint As Point
    Private dragFormPoint As Point


    Private Sub Form5_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Load item names into the ComboBox and populate the DataGridView
        LoadItems()
        LoadOrders()

        ' Populate cmbFilter with status options
        cmbFilter.Items.Clear()
        cmbFilter.Items.Add("All") ' Add an option for all statuses
        cmbFilter.Items.Add("Pending")
        cmbFilter.Items.Add("Cancelled")
        cmbFilter.Items.Add("Received")
        cmbFilter.SelectedItem = "All"
    End Sub


    Private Sub Form5_MouseDown(sender As Object, e As MouseEventArgs) Handles MyBase.MouseDown
        If e.Button = MouseButtons.Left Then
            drag = True
            dragCursorPoint = Cursor.Position
            dragFormPoint = Me.Location
        End If
    End Sub

    Private Sub Form5_MouseMove(sender As Object, e As MouseEventArgs) Handles MyBase.MouseMove
        If drag Then
            Dim dif As Point = Point.Subtract(Cursor.Position, New Size(dragCursorPoint))
            Me.Location = Point.Add(dragFormPoint, New Size(dif))
        End If
    End Sub

    Private Sub Form5_MouseUp(sender As Object, e As MouseEventArgs) Handles MyBase.MouseUp
        If e.Button = MouseButtons.Left Then
            drag = False
        End If
    End Sub

    Private Sub LoadItems()

        Try
            Using connection As New SqlConnection(connectionString)
                connection.Open()

                Dim query As String = "SELECT ItemName FROM Inventory"
                Using command As New SqlCommand(query, connection)
                    Using reader As SqlDataReader = command.ExecuteReader()
                        cmbItemName.Items.Clear()
                        While reader.Read()
                            cmbItemName.Items.Add(reader("ItemName").ToString())
                        End While
                    End Using
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("Error loading items: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub LoadOrders()

        Try
            Using connection As New SqlConnection(connectionString)
                connection.Open()


                Dim query As String = "SELECT O.OrderID, O.QuantityOrdered, O.Status, O.OrderDate, I.ItemName FROM Orders O JOIN Inventory I ON O.ItemID = I.ItemID"

                If cmbFilter.SelectedItem IsNot Nothing AndAlso cmbFilter.SelectedItem.ToString() <> "All" Then
                    query &= " WHERE O.Status = @Status"
                End If

                Dim adapter As New SqlDataAdapter(query, connection)

                If cmbFilter.SelectedItem IsNot Nothing AndAlso cmbFilter.SelectedItem.ToString() <> "All" Then
                    adapter.SelectCommand.Parameters.AddWithValue("@Status", cmbFilter.SelectedItem.ToString())
                End If

                Dim table As New DataTable()
                adapter.Fill(table)
                Guna2DataGridView1.DataSource = table
            End Using
        Catch ex As Exception
            MessageBox.Show("Error loading orders: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub


    Private Sub btnPlaceOrder_Click(sender As Object, e As EventArgs) Handles btnPlaceOrder.Click

        If cmbItemName.SelectedItem Is Nothing OrElse String.IsNullOrWhiteSpace(txtQuantity.Text) Then
            MessageBox.Show("Please fill in all fields.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        Dim itemName As String = cmbItemName.SelectedItem.ToString()
        Dim itemID As Integer

        Using connection As New SqlConnection(connectionString)
            connection.Open()
            Dim query As String = "SELECT ItemID FROM Inventory WHERE ItemName = @ItemName"
            Using command As New SqlCommand(query, connection)
                command.Parameters.AddWithValue("@ItemName", itemName)
                itemID = Convert.ToInt32(command.ExecuteScalar())
            End Using
        End Using

        Try
            Using connection As New SqlConnection(connectionString)
                connection.Open()

                Dim query As String = "INSERT INTO Orders (ItemID, QuantityOrdered, Status) VALUES (@ItemID, @QuantityOrdered, 'Pending')"
                Using command As New SqlCommand(query, connection)
                    command.Parameters.AddWithValue("@ItemID", itemID)
                    command.Parameters.AddWithValue("@QuantityOrdered", Convert.ToInt32(txtQuantity.Text))

                    command.ExecuteNonQuery()
                End Using
            End Using

            MessageBox.Show("Order placed successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
            LoadOrders()


            Form2.RefreshInventoryData()
        Catch ex As Exception
            MessageBox.Show("Error placing order: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnMarkReceived_Click(sender As Object, e As EventArgs) Handles btnMarkReceived.Click

        If Guna2DataGridView1.SelectedRows.Count = 0 Then
            MessageBox.Show("Please select an order to mark as received.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        Dim orderId As Integer = Convert.ToInt32(Guna2DataGridView1.SelectedRows(0).Cells("OrderID").Value)
        Dim itemName As String = Guna2DataGridView1.SelectedRows(0).Cells("ItemName").Value.ToString()
        Dim quantity As Integer = Convert.ToInt32(Guna2DataGridView1.SelectedRows(0).Cells("QuantityOrdered").Value)

        Try
            Using connection As New SqlConnection(connectionString)
                connection.Open()


                Dim updateOrderQuery As String = "UPDATE Orders SET Status = 'Received' WHERE OrderID = @OrderID"
                Using updateOrderCmd As New SqlCommand(updateOrderQuery, connection)
                    updateOrderCmd.Parameters.AddWithValue("@OrderID", orderId)
                    updateOrderCmd.ExecuteNonQuery()
                End Using


                Dim updateInventoryQuery As String = "UPDATE Inventory SET Quantity = Quantity + @Quantity WHERE ItemName = @ItemName"
                Using updateInventoryCmd As New SqlCommand(updateInventoryQuery, connection)
                    updateInventoryCmd.Parameters.AddWithValue("@Quantity", quantity)
                    updateInventoryCmd.Parameters.AddWithValue("@ItemName", itemName)
                    updateInventoryCmd.ExecuteNonQuery()
                End Using
            End Using

            MessageBox.Show("Order marked as received and inventory updated.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
            LoadOrders()

            Form2.RefreshInventoryData()
        Catch ex As Exception
            MessageBox.Show("Error marking order as received: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnCancelOrder_Click(sender As Object, e As EventArgs) Handles btnCancelOrder.Click

        If Guna2DataGridView1.SelectedRows.Count = 0 Then
            MessageBox.Show("Please select an order to cancel.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        Dim orderId As Integer = Convert.ToInt32(Guna2DataGridView1.SelectedRows(0).Cells("OrderID").Value)

        Try
            Using connection As New SqlConnection(connectionString)
                connection.Open()

                Dim query As String = "UPDATE Orders SET Status = 'Cancelled' WHERE OrderID = @OrderID"
                Using command As New SqlCommand(query, connection)
                    command.Parameters.AddWithValue("@OrderID", orderId)
                    command.ExecuteNonQuery()
                End Using
            End Using

            MessageBox.Show("Order cancelled successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
            LoadOrders()


            Form2.RefreshInventoryData()
        Catch ex As Exception
            MessageBox.Show("Error cancelling order: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Guna2Button1_Click(sender As Object, e As EventArgs) Handles Guna2Button1.Click

        Form2.RefreshInventoryData()


        Me.Close()
        Form1.Show()
        Form1.LoadInventoryData()
        Form1.LoadOrderData()
        Form1.CheckLowStock()
    End Sub

    Private Sub Guna2DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles Guna2DataGridView1.CellContentClick

    End Sub

    Private Sub cmbFilter_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbFilter.SelectedIndexChanged

        LoadOrders()
    End Sub
End Class
