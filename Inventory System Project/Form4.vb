Imports Microsoft.Data.SqlClient
Imports System.Drawing
Imports System.Drawing.Printing

Public Class Form4

    Dim connString As String = "Server=DESKTOP-3LP8SBD\SQLEXPRESS;Database=InventoryProjectDB;User Id=sa;Password=B1Admin;TrustServerCertificate=True;"


    Private drag As Boolean
    Private dragCursorPoint As Point
    Private dragFormPoint As Point


    Private Sub Form4_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        LoadLowStockItems()
        LoadTransactionData()
    End Sub

    Private Sub Form4_MouseDown(sender As Object, e As MouseEventArgs) Handles MyBase.MouseDown
        If e.Button = MouseButtons.Left Then
            drag = True
            dragCursorPoint = Cursor.Position
            dragFormPoint = Me.Location
        End If
    End Sub

    Private Sub Form4_MouseMove(sender As Object, e As MouseEventArgs) Handles MyBase.MouseMove
        If drag Then
            Dim dif As Point = Point.Subtract(Cursor.Position, New Size(dragCursorPoint))
            Me.Location = Point.Add(dragFormPoint, New Size(dif))
        End If
    End Sub

    Private Sub Form4_MouseUp(sender As Object, e As MouseEventArgs) Handles MyBase.MouseUp
        If e.Button = MouseButtons.Left Then
            drag = False
        End If
    End Sub




    Private Sub LoadTransactionData()
        Try

            Using connection As New SqlConnection(connString)
                connection.Open()


                Dim query As String = "SELECT TransactionID, TransactionDate, ItemID, QuantitySold, TotalPrice FROM Transactions"
                Dim adapter As New SqlDataAdapter(query, connection)
                Dim table As New DataTable()


                adapter.Fill(table)
                Guna2DataGridView2.DataSource = table
            End Using
        Catch ex As Exception
            MessageBox.Show("Error loading transaction data: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub LoadLowStockItems()
        Try

            Using connection As New SqlConnection(connString)
                connection.Open()

                Dim query As String = "SELECT ItemName, Quantity FROM Inventory WHERE Quantity < 10"
                Dim adapter As New SqlDataAdapter(query, connection)
                Dim table As New DataTable()


                adapter.Fill(table)
                Guna2DataGridView1.DataSource = table
            End Using
        Catch ex As Exception
            MessageBox.Show("Error loading low stock items: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
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
        Dim reportFont As New Font("Arial", 12, FontStyle.Bold)
        Dim font As New Font("Arial", 10)
        Dim headerFont As New Font("Arial", 12, FontStyle.Bold)
        Dim brush As New SolidBrush(Color.Black)

        Dim leftMargin As Single = e.MarginBounds.Left
        Dim topMargin As Single = e.MarginBounds.Top
        Dim rowHeight As Single = 20
        Dim currentY As Single = topMargin


        e.Graphics.DrawString("A&R Company Inventory System", companyFont, brush, leftMargin, currentY)
        currentY += 30


        e.Graphics.DrawString("Low Stock Report", reportFont, brush, leftMargin, currentY)
        currentY += 20


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


        For i As Integer = 0 To Guna2DataGridView1.Columns.Count - 1
            e.Graphics.DrawString(Guna2DataGridView1.Columns(i).HeaderText, headerFont, brush, leftMargin, currentY)
            leftMargin += columnWidths(i)
        Next


        currentY += rowHeight
        leftMargin = e.MarginBounds.Left

        For Each row As DataGridViewRow In Guna2DataGridView1.Rows
            If row.IsNewRow Then Continue For

            For i As Integer = 0 To Guna2DataGridView1.Columns.Count - 1
                e.Graphics.DrawString(row.Cells(i).Value.ToString(), font, brush, leftMargin, currentY)
                leftMargin += columnWidths(i)
            Next
            currentY += rowHeight
            leftMargin = e.MarginBounds.Left
        Next


        e.HasMorePages = False
    End Sub

    Private Sub Guna2Button1_Click(sender As Object, e As EventArgs) Handles Guna2Button1.Click
        Me.Close()
        Form1.Show()
        Form1.LoadInventoryData()
        Form1.LoadOrderData()
        Form1.CheckLowStock()
    End Sub

    Private Sub Guna2HtmlLabel1_Click(sender As Object, e As EventArgs) Handles Guna2HtmlLabel1.Click

    End Sub

    Private Sub Guna2DataGridView2_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles Guna2DataGridView2.CellContentClick

    End Sub
End Class
