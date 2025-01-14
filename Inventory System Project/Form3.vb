Imports System.Data.SqlClient
Imports Microsoft.Data.SqlClient

Public Class Form3
    ' Database connection string
    Private connectionString As String = "Server=DESKTOP-3LP8SBD\SQLEXPRESS;Database=InventoryProjectDB;User Id=sa;Password=B1Admin;TrustServerCertificate=True;"

    Private drag As Boolean
    Private dragCursorPoint As Point
    Private dragFormPoint As Point

    Private Sub Form3_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        ClearFields()


        txtCategory.Items.Clear()
        txtCategory.Items.Add("Electronics")
        txtCategory.Items.Add("Furniture")
        txtCategory.Items.Add("Accessories")
        txtCategory.Items.Add("Stationery")
    End Sub

    Private Sub Form3_MouseDown(sender As Object, e As MouseEventArgs) Handles MyBase.MouseDown
        If e.Button = MouseButtons.Left Then
            drag = True
            dragCursorPoint = Cursor.Position
            dragFormPoint = Me.Location
        End If
    End Sub

    Private Sub Form3_MouseMove(sender As Object, e As MouseEventArgs) Handles MyBase.MouseMove
        If drag Then
            Dim dif As Point = Point.Subtract(Cursor.Position, New Size(dragCursorPoint))
            Me.Location = Point.Add(dragFormPoint, New Size(dif))
        End If
    End Sub

    Private Sub Form3_MouseUp(sender As Object, e As MouseEventArgs) Handles MyBase.MouseUp
        If e.Button = MouseButtons.Left Then
            drag = False
        End If
    End Sub

    Private Sub Guna2HtmlLabel6_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Guna2Button1_Click(sender As Object, e As EventArgs) Handles Guna2Button1.Click
        Me.Hide()
        Form1.Show()
        Form1.LoadInventoryData()
        Form1.LoadOrderData()
        Form1.CheckLowStock()
    End Sub

    Private Sub txtItemName_TextChanged(sender As Object, e As EventArgs) Handles txtItemName.TextChanged

    End Sub

    Private Sub txtQuantity_TextChanged(sender As Object, e As EventArgs) Handles txtQuantity.TextChanged

    End Sub

    Private Sub txtPrice_TextChanged(sender As Object, e As EventArgs) Handles txtPrice.TextChanged

    End Sub

    Private Sub txtDescription_TextChanged(sender As Object, e As EventArgs) Handles txtDescription.TextChanged

    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click

        If ValidateInputs() Then

            SaveToDatabase()
        End If
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click

        ClearFields()
    End Sub

    Private Sub SaveToDatabase()
        Try
            Using connection As New SqlConnection(connectionString)
                connection.Open()

                Dim query As String = "INSERT INTO Inventory (ItemName, Category, Quantity, Price, Description) VALUES (@ItemName, @Category, @Quantity, @Price, @Description)"
                Using command As New SqlCommand(query, connection)
                    command.Parameters.AddWithValue("@ItemName", txtItemName.Text)
                    command.Parameters.AddWithValue("@Category", txtCategory.SelectedItem.ToString())
                    command.Parameters.AddWithValue("@Quantity", Convert.ToInt32(txtQuantity.Text))
                    command.Parameters.AddWithValue("@Price", Convert.ToDecimal(txtPrice.Text))
                    command.Parameters.AddWithValue("@Description", txtDescription.Text)

                    command.ExecuteNonQuery()
                End Using
            End Using

            MessageBox.Show("Item added successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
            ClearFields()
        Catch ex As Exception
            MessageBox.Show($"Error saving item: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Function ValidateInputs() As Boolean
        If String.IsNullOrWhiteSpace(txtItemName.Text) Then
            MessageBox.Show("Item name is required.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtItemName.Focus()
            Return False
        ElseIf txtCategory.SelectedItem Is Nothing Then
            MessageBox.Show("Category is required.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtCategory.Focus()
            Return False
        ElseIf Not Integer.TryParse(txtQuantity.Text, Nothing) OrElse Convert.ToInt32(txtQuantity.Text) <= 0 Then
            MessageBox.Show("Quantity must be a positive integer.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtQuantity.Focus()
            Return False
        ElseIf Not Decimal.TryParse(txtPrice.Text, Nothing) OrElse Convert.ToDecimal(txtPrice.Text) <= 0 Then
            MessageBox.Show("Price must be a positive number.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtPrice.Focus()
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtDescription.Text) Then
            MessageBox.Show("Description is required.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtDescription.Focus()
            Return False
        End If

        Return True
    End Function

    Private Sub ClearFields()
        txtItemName.Clear()
        txtCategory.SelectedIndex = -1
        txtQuantity.Clear()
        txtPrice.Clear()
        txtDescription.Clear()
    End Sub

    Private Sub txtCategory_SelectedIndexChanged(sender As Object, e As EventArgs) Handles txtCategory.SelectedIndexChanged

        If txtCategory.SelectedItem IsNot Nothing Then
            Dim selectedCategory As String = txtCategory.SelectedItem.ToString()
            MessageBox.Show($"You selected: {selectedCategory}", "Category Selected", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub
End Class
