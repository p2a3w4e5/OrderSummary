Imports System.IO
Imports Microsoft.Office.Interop

Public Class frmMain

    Friend excel As Excel.Application
    Friend wb As Excel.Workbook
    Friend ws As Object
    Dim itemList As New List(Of OrderItem)
    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CurrentOrdersFolder = My.Settings.CurrentOrders
        Me.txtCurrentOrders.Text = CurrentOrdersFolder


        Try
            excel = New Excel.Application
        Catch ex As Exception
            MsgBox("Excel is not open.", MsgBoxStyle.Critical, "ERROR")

        End Try
    End Sub
    Private Sub frmMain_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        excel = Nothing
        GC.Collect()
    End Sub
    Private Sub btnCurrentOrders_Click(sender As Object, e As EventArgs) Handles btnCurrentOrders.Click

        If Me.FolderBrowserDialog1.ShowDialog = DialogResult.OK Then
            CurrentOrdersFolder = FolderBrowserDialog1.SelectedPath
            My.Settings.CurrentOrders = CurrentOrdersFolder
            Me.txtCurrentOrders.Text = CurrentOrdersFolder
            My.Settings.Save()
        End If

        ''--------------------------------------------------------------------------------
        ''create Excel object
        ''--------------------------------------------------------------------------------
        'Try
        '    'excel = CreateObject("Excel.Application")
        '    excel = New Excel.Application
        '    'wb = excel.ActiveWorkbook
        '    'ws = wb.WorkSheets(1)
        'Catch ex As Exception
        '    MsgBox("Excel is not open.", MsgBoxStyle.Critical, "ERROR")
        '    'End
        'End Try

    End Sub


    Private Sub btnCalculate_Click(sender As Object, e As EventArgs) Handles btnCalculate.Click
        Try
            excel.Visible = True
            If Me.txtCurrentOrders.Text.Trim = "" Then
                MessageBox.Show("Current orders path is required")
                Exit Sub
            End If

            Dim di As New DirectoryInfo(Me.txtCurrentOrders.Text)
            Dim oi As OrderItem
            Dim row As Integer = 1
            Dim startRead As Boolean = False

            itemList.Clear()

            For Each fi As FileInfo In di.GetFiles
                Application.DoEvents()

                If fi.Extension.Contains("xlsx") OrElse fi.Extension.Contains("xls") Then
                    Me.lblCurrentFile.Text = fi.FullName
                    wb = excel.Workbooks.Open(fi.FullName)

                    For i As Integer = 1 To 300
                        Application.DoEvents()
                        Dim colA As String = Trim(wb.Worksheets(1).Cells(i, 1).Value)

                        If colA.ToUpper = "ITEM" Then
                            startRead = True
                        End If

                        If startRead = True AndAlso colA = "" Then
                            startRead = False
                            'Close workbook
                            wb.Close()
                            Exit For
                        End If

                        If startRead = True AndAlso colA.ToUpper <> "ITEM" Then
                            oi = New OrderItem
                            oi.ItemName = Trim(wb.Worksheets(1).Cells(i, 1).Value)
                            oi.ItemCount = Val(Trim(wb.Worksheets(1).Cells(i, 3).Value))

                            Dim findItem As OrderItem = itemList.Find(Function(x) x.ItemName = oi.ItemName)

                            If Not IsNothing(findItem) Then
                                findItem.ItemCount = findItem.ItemCount + oi.ItemCount
                            Else
                                itemList.Add(oi)
                            End If
                            Application.DoEvents()
                        End If

                    Next

                End If

            Next

            'Create excel sheet

            wb = excel.Workbooks.Add()
            wb.Worksheets(1).Cells(2, 1).Value = "Product and Materials Needed for All Current Orders"
            wb.Worksheets(1).Cells(4, 1).Value = "Item"
            wb.Worksheets(1).Cells(4, 2).Value = "Oz or Count"

            row = 5

            For Each item As OrderItem In itemList
                wb.Worksheets(1).Cells(row, 1).Value = item.ItemName
                wb.Worksheets(1).Cells(row, 2).Value = item.ItemCount
                row += 1
            Next

            Application.DoEvents()

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub


End Class
