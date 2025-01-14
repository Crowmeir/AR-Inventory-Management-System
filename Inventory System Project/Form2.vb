Imports System.Data.SqlClient
Imports Microsoft.Data.SqlClient
Imports System.Drawing
Imports System.Drawing.Printing

Public Class Form2

    Private drag As Boolean
    Private dragCursorPoint As Point
    Private dragFormPoint As Point


    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LoadInventoryData()
    End Sub

    Private Sub Form2_MouseDown(sender As Object, e As MouseEventArgs) Handles MyBase.MouseDown
        If e.Button = MouseButtons.Left Then
            drag = True
            dragCursorPoint = Cursor.Position
            dragFormPoint = Me.Location
        End If
    End Sub

    Private Sub Form2_MouseMove(sender As Object, e As MouseEventArgs) Handles MyBase.MouseMove
        If drag Then
            Dim dif As Point = Point.Subtract(Cursor.Position, New Size(dragCursorPoint))
            Me.Location = Point.Add(dragFormPoint, New Size(dif))
        End If
    End Sub

    Private Sub Form2_MouseUp(sender As Object, e As MouseEventArgs) Handles MyBase.MouseUp
        If e.Button = MouseButtons.Left Then
            drag = False
        End If
    End Sub

    Private Function GetConnection() As SqlConnection
        Dim connString As String = "Server=DESKTOP-3LP8SBD\SQLEXPRESS;Database=InventoryProjectDB;User Id=sa;Password=B1Admin;TrustServerCertificate=True;"
        Return New SqlConnection(connString)
    End Function

    Private Sub LoadInventoryData()
        Using conn As SqlConnection = GetConnection()
            conn.Open()
            Dim query As String = "SELECT ItemID, ItemName, Category, Quantity, Price, Description, LastUpdated FROM Inventory"
            Dim cmd As New SqlCommand(query, conn)
            Dim da As New SqlDataAdapter(cmd)
            Dim dt As New DataTable()
            da.Fill(dt)
            Guna2DataGridView1.DataSource = dt
        End Using
    End Sub

    Private Sub txtSearch_TextChanged(sender As Object, e As EventArgs) Handles txtSearch.TextChanged
        If txtSearch.Text.Trim() <> String.Empty Then
            SearchInventory(txtSearch.Text)
        Else
            LoadInventoryData()
        End If
    End Sub

    Private Sub SearchInventory(query As String)
        Using conn As SqlConnection = GetConnection()
            conn.Open()
            Dim searchQuery As String = "SELECT ItemID, ItemName, Category, Quantity, Price, Description, LastUpdated FROM Inventory WHERE ItemName LIKE @query OR Category LIKE @query"
            Dim cmd As New SqlCommand(searchQuery, conn)
            cmd.Parameters.AddWithValue("@query", "%" & query & "%")
            Dim da As New SqlDataAdapter(cmd)
            Dim dt As New DataTable()
            da.Fill(dt)
            Guna2DataGridView1.DataSource = dt
        End Using
    End Sub

    Private Sub btnDelete_Click(sender As Object, e As EventArgs) Handles btnDelete.Click

        If Guna2DataGridView1.SelectedRows.Count > 0 Then
            Dim itemID As Integer = Convert.ToInt32(Guna2DataGridView1.SelectedRows(0).Cells("ItemID").Value)


            If MessageBox.Show("Are you sure you want to delete this item and its associated orders?", "Confirm Delete", MessageBoxButtons.YesNo) = DialogResult.Yes Then
                Using conn As SqlConnection = GetConnection()
                    Try
                        conn.Open()


                        Dim transaction As SqlTransaction = conn.BeginTransaction()


                        Dim deleteOrdersQuery As String = "DELETE FROM Orders WHERE ItemID = @ItemID"
                        Dim deleteOrdersCmd As New SqlCommand(deleteOrdersQuery, conn, transaction)
                        deleteOrdersCmd.Parameters.AddWithValue("@ItemID", itemID)
                        deleteOrdersCmd.ExecuteNonQuery()


                        Dim deleteItemQuery As String = "DELETE FROM Inventory WHERE ItemID = @ItemID"
                        Dim deleteItemCmd As New SqlCommand(deleteItemQuery, conn, transaction)
                        deleteItemCmd.Parameters.AddWithValue("@ItemID", itemID)
                        deleteItemCmd.ExecuteNonQuery()


                        transaction.Commit()

                        MessageBox.Show("Item deleted successfully.")

                        LoadInventoryData()

                    Catch ex As Exception
                        MessageBox.Show("An error occurred: " & ex.Message)
                    End Try
                End Using
            End If
        Else
            MessageBox.Show("Please select an item to delete.")
        End If
    End Sub

    Private Sub DeleteItem(itemID As Integer)
        Using conn As SqlConnection = GetConnection()
            conn.Open()
            Dim query As String = "DELETE FROM Inventory WHERE ItemID = @itemID"
            Dim cmd As New SqlCommand(query, conn)
            cmd.Parameters.AddWithValue("@itemID", itemID)
            cmd.ExecuteNonQuery()
        End Using
    End Sub

    Private Sub Guna2Button1_Click(sender As Object, e As EventArgs) Handles Guna2Button1.Click
        Me.Hide()
        Form1.Show()
        Form1.LoadInventoryData()
        Form1.LoadOrderData()
        Form1.CheckLowStock()
    End Sub

    Private Sub Guna2DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles Guna2DataGridView1.CellContentClick
    End Sub

    Public Sub RefreshInventoryData()
        LoadInventoryData()
    End Sub


    Private Sub btnPrint_Click(sender As Object, e As EventArgs) Handles btnPrint.Click

        Dim printPreview As New PrintPreviewDialog()


        Dim printDoc As New PrintDocument()


        AddHandler printDoc.PrintPage, AddressOf Me.PrintDocument_PrintPage


        printPreview.Document = printDoc

        printPreview.ShowDialog()
    End Sub


    Private Sub PrintDocument_PrintPage(sender As Object, e As PrintPageEventArgs)

        Dim companyFont As New Font("Arial", 16, FontStyle.Bold)
        Dim headerFont As New Font("Arial", 12, FontStyle.Bold)
        Dim bodyFont As New Font("Arial", 10)
        Dim brush As New SolidBrush(Color.Black)

        Dim leftMargin As Single = e.MarginBounds.Left
        Dim topMargin As Single = e.MarginBounds.Top
        Dim rowHeight As Single = 20
        Dim currentY As Single = topMargin


        e.Graphics.DrawString("A&R Company Inventory System", companyFont, brush, leftMargin, currentY)
        currentY += 30


        e.Graphics.DrawString("Inventory Record - " & DateTime.Now.ToString("MMMM dd, yyyy"), headerFont, brush, leftMargin, currentY)
        currentY += 30


        Dim columnWidths As New List(Of Single)
        For Each column As DataGridViewColumn In Guna2DataGridView1.Columns

            Dim maxColumnWidth As Single = column.HeaderText.Length * 8
            For Each row As DataGridViewRow In Guna2DataGridView1.Rows
                If row.Cells(column.Index).Value IsNot Nothing Then
                    maxColumnWidth = Math.Max(maxColumnWidth, row.Cells(column.Index).Value.ToString().Length * 8)
                End If
            Next
            columnWidths.Add(maxColumnWidth)
        Next


        Dim headerX As Single = leftMargin
        For i As Integer = 0 To Guna2DataGridView1.Columns.Count - 1
            e.Graphics.DrawString(Guna2DataGridView1.Columns(i).HeaderText, headerFont, brush, headerX, currentY)
            headerX += columnWidths(i)
        Next
        currentY += rowHeight


        headerX = leftMargin
        For Each row As DataGridViewRow In Guna2DataGridView1.Rows
            If row.IsNewRow Then Continue For

            Dim rowX As Single = headerX
            For i As Integer = 0 To Guna2DataGridView1.Columns.Count - 1
                e.Graphics.DrawString(row.Cells(i).Value.ToString(), bodyFont, brush, rowX, currentY)
                rowX += columnWidths(i)
            Next
            currentY += rowHeight
            headerX = leftMargin

            If currentY + rowHeight > e.MarginBounds.Bottom Then
                e.HasMorePages = True
                Return
            End If
        Next


        e.HasMorePages = False
    End Sub
End Class
