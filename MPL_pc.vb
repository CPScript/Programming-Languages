Imports MySql.Data.MySqlClient
Imports Microsoft.Office.Interop
Imports System.Runtime.InteropServices
Imports System.Reflection
Imports System.Net.Mail

Public Class MPL_pc


    'Private Sub PR_grid_ColumnHeaderMouseDoubleClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles PR_grid.ColumnHeaderMouseDoubleClick

    '    'attn_l.Text = "Attn:  " & PR_grid.Columns.Item(e.ColumnIndex).HeaderText
    '    'whereiam = e.ColumnIndex
    'End Sub

    'Private Sub CreatePackingSlipToolStripMenuItem_Click(sender As Object, e As EventArgs)
    '    '--- add a column shipment
    '    'shipping_wiz.ShowDialog()
    '    'PR_grid.Columns.Add("Shipping 1", "Shipping 1")

    'End Sub

    Private Sub MPL_pc_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try

            p_number2.Text = Procurement_Overview.job_label.Text

            Dim cmd51 As New MySqlCommand
            cmd51.Parameters.Clear()
            cmd51.Parameters.AddWithValue("@mr_name", Procurement_Overview.Text)
            cmd51.CommandText = "SELECT shipping_ad from Material_Request.mr where mr_name = @mr_name"
            cmd51.Connection = Login.Connection
            Dim reader51 As MySqlDataReader
            reader51 = cmd51.ExecuteReader

            If reader51.HasRows Then
                While reader51.Read
                    If reader51.IsDBNull(0) = False Then
                        address_ship.Text = reader51(0).ToString
                    End If
                End While
            End If

            reader51.Close()

            '----------------------------
            Dim cmd511 As New MySqlCommand
            cmd511.Parameters.Clear()
            cmd511.Parameters.AddWithValue("@job", Procurement_Overview.job_label.Text)
            cmd511.CommandText = "SELECT Job_description from management.projects where Job_number = @job"
            cmd511.Connection = Login.Connection
            Dim reader511 As MySqlDataReader
            reader511 = cmd511.ExecuteReader

            If reader511.HasRows Then
                While reader511.Read
                    p_name2.Text = reader511(0).ToString
                End While
            End If

            reader511.Close()

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub


    Private Sub RemoveRowToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RemoveRowToolStripMenuItem.Click
        If PR_grid.SelectedRows.Count > 0 Then
            For Each r As DataGridViewRow In PR_grid.SelectedRows
                Try
                    PR_grid.Rows.Remove(r)
                Catch ex As Exception
                    MessageBox.Show("This row cannot be deleted")
                End Try
            Next
        Else
            MessageBox.Show("Select the row you want to delete.")
        End If
    End Sub

    Private Sub RemoveShippingToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RemoveShippingToolStripMenuItem.Click
        '-- export to excel
        Dim appPath As String = Application.StartupPath()

        Dim xlApp As Excel.Application = New Microsoft.Office.Interop.Excel.Application()
        Dim xlWorkSheet As Excel.Worksheet

        xlApp.DisplayAlerts = False

        If xlApp Is Nothing Then
            MessageBox.Show("Excel is not properly installed!!")
        Else

            Try
                Dim wb As Excel.Workbook = xlApp.Workbooks.Open("O:\atlanta\APL\Slip.xlsx")   ' Dim wb As Excel.Workbook = xlApp.Workbooks.Open(appPath & "\Template.xlsx")
                xlWorkSheet = wb.Sheets("Sheet1")

                xlWorkSheet.Cells(7, 2) = p_name2.Text  'job name
                xlWorkSheet.Cells(10, 2) = address_ship.Text  'shipping

                xlWorkSheet.Cells(8, 2) = p_number2.Text  'job number
                xlWorkSheet.Cells(11, 2) = "Phone: " & phone.Text
                xlWorkSheet.Cells(11, 3) = "Attn: " & attn.Text
                xlWorkSheet.Cells(14, 2) = z_number.Text
                xlWorkSheet.Cells(12, 3) = "Date: " & DateTimePicker1.Value.Date

                Dim z As Integer : z = 17

                For i = 0 To PR_grid.Rows.Count - 1
                    xlWorkSheet.Cells(z, 1) = PR_grid.Rows(i).Cells(0).Value
                    xlWorkSheet.Cells(z, 2) = PR_grid.Rows(i).Cells(1).Value
                    xlWorkSheet.Cells(z, 3) = PR_grid.Rows(i).Cells(2).Value
                    xlWorkSheet.Cells(z, 4) = PR_grid.Rows(i).Cells(3).Value

                    z = z + 1
                Next


                SaveFileDialog1.Filter = "Excel Files|*.xlsx;*.xlsm"

                If SaveFileDialog1.ShowDialog = DialogResult.OK Then
                    wb.SaveCopyAs(SaveFileDialog1.FileName.ToString)
                End If

                wb.Close(False)

                Marshal.ReleaseComObject(xlApp)

                MessageBox.Show("Packing List exported successfully!")

            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try

        End If

    End Sub
End Class
