
Imports iText.Kernel.Colors
Imports iText.Kernel.Font
Imports iText.Kernel.Pdf
Imports iText.Layout
Imports iText.Layout.Element
Imports iText.Layout.Properties
Imports System.Threading
Imports ClosedXML.Excel
Imports PLANDAY.OvertimeForm

Public Class ExportPDF

    Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hWnd As IntPtr, ByVal wMsg As Int32, ByVal wParam As Int32, ByVal lParam As IntPtr) As IntPtr

    Private Sub CLOSE_button_Click(sender As Object, e As EventArgs) Handles CLOSE_button.Click
        Me.Close()
    End Sub

    Private Sub Columns_list_KeyDown(sender As Object, e As KeyEventArgs) Handles Columns_list.KeyDown
        If e.KeyCode = Keys.Enter Then
            CREATE_button_Click(sender, e)
        End If
        If e.KeyData = (Keys.A Or Keys.Control) Then
            SendMessage(Columns_list.Handle, &H185, 1, -1)
        End If
    End Sub

    Public Sub WorkingProgress(x As Integer, y As Integer, width As Integer, height As Integer, normalDPI As Boolean)
        myWorkingForm = New WorkingForm() ' Must be created on this thread!
        myWorkingForm.StartPosition = FormStartPosition.Manual
        myWorkingForm.Location = New Point(x + width / 2 - myWorkingForm.Width / 2, y + height / 2 - myWorkingForm.Height / 2 + If(normalDPI, 52, 80))
        System.Windows.Forms.Application.Run(myWorkingForm)
    End Sub

    Private Sub CREATE_button_Click(sender As Object, e As EventArgs) Handles CREATE_button.Click

        Try
            If Columns_list.SelectedItems.Count > 0 Then
                Columns_list.DrawMode = DrawMode.OwnerDrawFixed
                Dim kw As String
                Dim file_name As String

                kw = OvertimeForm.loaded.Text

                If Me.Text = "Export PDF" Then
                    file_name = kw & "_Überstunden_Export_am_" & GetDateTimeStamp() & ".pdf"
                Else
                    file_name = kw & "_Überstunden_Export_am_" & GetDateTimeStamp() & ".xlsx"
                End If

                If Me.Text = "Export PDF" Then
                    Dim th As System.Threading.Thread = New Threading.Thread(Sub() WorkingProgress(Me.Location.X, Me.Location.Y, Me.Width, Me.Height, normalDPI))
                    th.SetApartmentState(ApartmentState.STA)
                    th.Start()

                    Dim summe As Double
                    Dim writer As PdfWriter = New PdfWriter(export_folder & "\" & file_name)
                    Dim pdf As PdfDocument = New PdfDocument(writer)
                    Dim document As Document = New Document(pdf)

                    Dim FONT As String = "C:\Windows\Fonts\Calibri.ttf"
                    Dim FONT_BOLD As String = "C:\Windows\Fonts\Calibrib.ttf"
                    Dim defaultFont = PdfFontFactory.CreateFont(FONT, "Cp1250", True)
                    Dim defaultFont_BOLD = PdfFontFactory.CreateFont(FONT_BOLD, "Cp1250", True)
                    document.SetFont(defaultFont).SetFontSize(11)

                    Dim table As Table = New Table(Columns_list.SelectedItems.Count).UseAllAvailableWidth

                    For i = 0 To Columns_list.SelectedItems.Count - 1
                        If i = 0 Then
                            table.AddCell(New Cell().SetPadding(0).SetBackgroundColor(ColorConstants.BLACK).SetTextAlignment(TextAlignment.LEFT).Add(New Paragraph(Columns_list.SelectedItems(i).ToString).SetMultipliedLeading(1.1F).SetPadding(1)).SetFontColor(ColorConstants.WHITE).SetFont(defaultFont_BOLD))
                        Else
                            table.AddCell(New Cell().SetPadding(0).SetBackgroundColor(ColorConstants.BLACK).SetTextAlignment(TextAlignment.RIGHT).Add(New Paragraph(Columns_list.SelectedItems(i).ToString).SetMultipliedLeading(1.1F).SetPadding(1)).SetFontColor(ColorConstants.WHITE).SetFont(defaultFont_BOLD))
                        End If
                    Next

                    Dim cell As Cell
                    For i = 0 To OvertimeForm.select_employees.CheckedItems.Count - 1
                        For j = 0 To Columns_list.SelectedIndices.Count - 1
                            Dim clr = OvertimeForm.select_employees.CheckedItems(i).SubItems(Columns_list.SelectedIndices(j)).ForeColor
                            If j = 0 Then
                                cell = New Cell().SetPadding(0).SetTextAlignment(TextAlignment.LEFT).Add(New Paragraph(OvertimeForm.select_employees.CheckedItems(i).SubItems(Columns_list.SelectedIndices(j)).Text).SetMultipliedLeading(1.1F))
                            Else
                                cell = New Cell().SetPadding(0).SetTextAlignment(TextAlignment.RIGHT).Add(New Paragraph(OvertimeForm.select_employees.CheckedItems(i).SubItems(Columns_list.SelectedIndices(j)).Text).SetMultipliedLeading(1.1F))
                            End If
                            cell.SetFontColor(New DeviceRgb(clr))
                            table.AddCell(cell)
                        Next
                    Next

                    document.Add(table)

                    For i = 0 To OvertimeForm.select_employees.CheckedItems.Count - 1
                        Dim sheetName As String
                        If Len(OvertimeForm.select_employees.CheckedItems(i).Text) > 31 Then
                            sheetName = Strings.Left(OvertimeForm.select_employees.CheckedItems(i).Text, 31)
                        Else
                            sheetName = OvertimeForm.select_employees.CheckedItems(i).Text
                        End If

                        document.Add(New AreaBreak(AreaBreakType.NEXT_PAGE))
                        table = New Table(10).UseAllAvailableWidth
                        table.AddCell(New Cell().SetPadding(0).SetBackgroundColor(ColorConstants.BLACK).SetTextAlignment(TextAlignment.LEFT).Add(New Paragraph("Datum").SetMultipliedLeading(1.1F)).SetFontColor(ColorConstants.WHITE).SetFont(defaultFont_BOLD))
                        table.AddCell(New Cell().SetPadding(0).SetBackgroundColor(ColorConstants.BLACK).SetTextAlignment(TextAlignment.LEFT).Add(New Paragraph("Schichtart").SetMultipliedLeading(1.1F)).SetFontColor(ColorConstants.WHITE).SetFont(defaultFont_BOLD))

                        table.AddCell(New Cell().SetPadding(0).SetBackgroundColor(ColorConstants.BLACK).SetTextAlignment(TextAlignment.RIGHT).Add(New Paragraph("Start").SetMultipliedLeading(1.1F)).SetFontColor(ColorConstants.WHITE).SetFont(defaultFont_BOLD))
                        table.AddCell(New Cell().SetPadding(0).SetBackgroundColor(ColorConstants.BLACK).SetTextAlignment(TextAlignment.RIGHT).Add(New Paragraph("Ende").SetMultipliedLeading(1.1F)).SetFontColor(ColorConstants.WHITE).SetFont(defaultFont_BOLD))
                        table.AddCell(New Cell().SetPadding(0).SetBackgroundColor(ColorConstants.BLACK).SetTextAlignment(TextAlignment.RIGHT).Add(New Paragraph("Pause von").SetMultipliedLeading(1.1F)).SetFontColor(ColorConstants.WHITE).SetFont(defaultFont_BOLD))
                        table.AddCell(New Cell().SetPadding(0).SetBackgroundColor(ColorConstants.BLACK).SetTextAlignment(TextAlignment.RIGHT).Add(New Paragraph("Pause bis").SetMultipliedLeading(1.1F)).SetFontColor(ColorConstants.WHITE).SetFont(defaultFont_BOLD))
                        table.AddCell(New Cell().SetPadding(0).SetBackgroundColor(ColorConstants.BLACK).SetTextAlignment(TextAlignment.RIGHT).Add(New Paragraph("Pause von").SetMultipliedLeading(1.1F)).SetFontColor(ColorConstants.WHITE).SetFont(defaultFont_BOLD))
                        table.AddCell(New Cell().SetPadding(0).SetBackgroundColor(ColorConstants.BLACK).SetTextAlignment(TextAlignment.RIGHT).Add(New Paragraph("Pause bis").SetMultipliedLeading(1.1F)).SetFontColor(ColorConstants.WHITE).SetFont(defaultFont_BOLD))
                        table.AddCell(New Cell().SetPadding(0).SetBackgroundColor(ColorConstants.BLACK).SetTextAlignment(TextAlignment.RIGHT).Add(New Paragraph("IST").SetMultipliedLeading(1.1F)).SetFontColor(ColorConstants.WHITE).SetFont(defaultFont_BOLD))
                        table.AddCell(New Cell().SetPadding(0).SetBackgroundColor(ColorConstants.BLACK).SetTextAlignment(TextAlignment.RIGHT).Add(New Paragraph("Urlaubstage").SetMultipliedLeading(1.1F)).SetFontColor(ColorConstants.WHITE).SetFont(defaultFont_BOLD))

                        summe = 0
                        If OvertimeForm.employee_kw_hours.ContainsKey(sheetName) Then
                            For j = 0 To OvertimeForm.employee_kw_hours(sheetName).Count - 1
                                If j <> OvertimeForm.employee_kw_hours(sheetName).Count - 1 Then
                                    summe = summe + OvertimeForm.employee_kw_hours(sheetName)(j).Shift_length
                                    table.AddCell(New Cell().SetPadding(0).SetTextAlignment(TextAlignment.LEFT).Add(New Paragraph(OvertimeForm.employee_kw_hours(sheetName)(j).Shift_date).SetMultipliedLeading(1.1F)))
                                    table.AddCell(New Cell().SetPadding(0).SetTextAlignment(TextAlignment.LEFT).Add(New Paragraph(OvertimeForm.employee_kw_hours(sheetName)(j).Shift_type).SetMultipliedLeading(1.1F)))
                                    table.AddCell(New Cell().SetPadding(0).SetTextAlignment(TextAlignment.RIGHT).Add(New Paragraph(OvertimeForm.employee_kw_hours(sheetName)(j).Shift_from).SetMultipliedLeading(1.1F)))
                                    table.AddCell(New Cell().SetPadding(0).SetTextAlignment(TextAlignment.RIGHT).Add(New Paragraph(OvertimeForm.employee_kw_hours(sheetName)(j).Shift_to).SetMultipliedLeading(1.1F)))
                                    table.AddCell(New Cell().SetPadding(0).SetTextAlignment(TextAlignment.RIGHT).Add(New Paragraph(OvertimeForm.employee_kw_hours(sheetName)(j).Shift_break_1_from).SetMultipliedLeading(1.1F)))
                                    table.AddCell(New Cell().SetPadding(0).SetTextAlignment(TextAlignment.RIGHT).Add(New Paragraph(OvertimeForm.employee_kw_hours(sheetName)(j).Shift_break_1_to).SetMultipliedLeading(1.1F)))
                                    table.AddCell(New Cell().SetPadding(0).SetTextAlignment(TextAlignment.RIGHT).Add(New Paragraph(OvertimeForm.employee_kw_hours(sheetName)(j).Shift_break_2_from).SetMultipliedLeading(1.1F)))
                                    table.AddCell(New Cell().SetPadding(0).SetTextAlignment(TextAlignment.RIGHT).Add(New Paragraph(OvertimeForm.employee_kw_hours(sheetName)(j).Shift_break_2_to).SetMultipliedLeading(1.1F)))
                                    table.AddCell(New Cell().SetPadding(0).SetTextAlignment(TextAlignment.RIGHT).Add(New Paragraph(OvertimeForm.employee_kw_hours(sheetName)(j).Shift_length).SetMultipliedLeading(1.1F)).SetFontColor(ColorConstants.BLUE))
                                    table.AddCell(New Cell().SetPadding(0).SetTextAlignment(TextAlignment.RIGHT).Add(New Paragraph("").SetMultipliedLeading(1.1F)))
                                Else
                                    summe = summe + OvertimeForm.employee_kw_hours(sheetName)(j).Shift_length
                                    table.AddCell(New Cell().SetPadding(0).SetTextAlignment(TextAlignment.LEFT).Add(New Paragraph("").SetMultipliedLeading(1.1F)))
                                    table.AddCell(New Cell().SetPadding(0).SetTextAlignment(TextAlignment.LEFT).Add(New Paragraph("").SetMultipliedLeading(1.1F)))
                                    table.AddCell(New Cell().SetPadding(0).SetTextAlignment(TextAlignment.RIGHT).Add(New Paragraph("").SetMultipliedLeading(1.1F)))
                                    table.AddCell(New Cell().SetPadding(0).SetTextAlignment(TextAlignment.RIGHT).Add(New Paragraph("").SetMultipliedLeading(1.1F)))
                                    table.AddCell(New Cell().SetPadding(0).SetTextAlignment(TextAlignment.RIGHT).Add(New Paragraph("").SetMultipliedLeading(1.1F)))
                                    table.AddCell(New Cell().SetPadding(0).SetTextAlignment(TextAlignment.RIGHT).Add(New Paragraph("").SetMultipliedLeading(1.1F)))
                                    table.AddCell(New Cell().SetPadding(0).SetTextAlignment(TextAlignment.RIGHT).Add(New Paragraph("").SetMultipliedLeading(1.1F)))
                                    table.AddCell(New Cell().SetPadding(0).SetTextAlignment(TextAlignment.RIGHT).Add(New Paragraph("Urlaub:").SetMultipliedLeading(1.1F)))
                                    table.AddCell(New Cell().SetPadding(0).SetTextAlignment(TextAlignment.RIGHT).Add(New Paragraph(OvertimeForm.employee_kw_hours(sheetName)(j).Shift_length).SetMultipliedLeading(1.1F)).SetFontColor(ColorConstants.BLUE))
                                    table.AddCell(New Cell().SetPadding(0).SetTextAlignment(TextAlignment.RIGHT).Add(New Paragraph(OvertimeForm.employee_kw_hours(sheetName)(j).Holiday_days).SetMultipliedLeading(1.1F)))
                                End If
                            Next

                            table.AddCell(New Cell(1, 7).SetPadding(0).SetTextAlignment(TextAlignment.LEFT).Add(New Paragraph(OvertimeForm.select_employees.CheckedItems(i).Text).SetMultipliedLeading(1.1F)).SetBorder(Borders.Border.NO_BORDER).SetFont(defaultFont_BOLD))
                            table.AddCell(New Cell().SetPadding(0).SetTextAlignment(TextAlignment.RIGHT).Add(New Paragraph("TOTAL:").SetMultipliedLeading(1.1F)).SetBorder(Borders.Border.NO_BORDER).SetFont(defaultFont_BOLD))
                            table.AddCell(New Cell().SetPadding(0).SetTextAlignment(TextAlignment.RIGHT).Add(New Paragraph(summe).SetMultipliedLeading(1.1F)).SetBorder(Borders.Border.NO_BORDER).SetFont(defaultFont_BOLD))
                        End If

                        document.Add(table)
                    Next

                    document.Close()
                Else

                    Dim bkWorkBook As XLWorkbook
                    Dim shWorkSheet As IXLWorksheet
                    Dim i As Integer
                    Dim j As Integer

                    Dim th As System.Threading.Thread = New Threading.Thread(Sub() WorkingProgress(Me.Location.X, Me.Location.Y, Me.Width, Me.Height, normalDPI))
                    th.SetApartmentState(ApartmentState.STA)
                    th.Start()

                    bkWorkBook = New XLWorkbook()
                    shWorkSheet = bkWorkBook.Worksheets.Add("Überstunden")

                    For i = 0 To Columns_list.SelectedItems.Count - 1
                        shWorkSheet.Cell(1, i + 1).Value = Columns_list.SelectedItems(i)
                    Next
                    shWorkSheet.Range(shWorkSheet.Cell(1, 1), shWorkSheet.Cell(1, Columns_list.SelectedItems.Count)).Style.Fill.BackgroundColor = XLColor.Black
                    shWorkSheet.Range(shWorkSheet.Cell(1, 1), shWorkSheet.Cell(1, Columns_list.SelectedItems.Count)).Style.Font.FontColor = XLColor.White
                    shWorkSheet.Range(shWorkSheet.Cell(1, 2), shWorkSheet.Cell(1, Columns_list.SelectedItems.Count)).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right

                    For i = 0 To OvertimeForm.select_employees.CheckedItems.Count - 1
                        For j = 0 To Columns_list.SelectedIndices.Count - 1
                            Dim clr = OvertimeForm.select_employees.CheckedItems(i).SubItems(Columns_list.SelectedIndices(j)).ForeColor
                            If j = 0 Then
                                shWorkSheet.Cell(i + 2, j + 1).Value = OvertimeForm.select_employees.CheckedItems(i).SubItems(Columns_list.SelectedIndices(j)).Text
                            Else
                                If OvertimeForm.select_employees.CheckedItems(i).SubItems(Columns_list.SelectedIndices(j)).Text = "" Then
                                    shWorkSheet.Cell(i + 2, j + 1).Value = ""
                                Else
                                    shWorkSheet.Cell(i + 2, j + 1).Value = CDbl(OvertimeForm.select_employees.CheckedItems(i).SubItems(Columns_list.SelectedIndices(j)).Text)
                                End If
                            End If
                            shWorkSheet.Cell(i + 2, j + 1).Style.Font.FontColor = XLColor.FromColor(clr)
                        Next
                    Next

                    For i = 1 To Columns_list.SelectedItems.Count
                        Dim max_length As Integer = 0
                        For j = 1 To OvertimeForm.select_employees.CheckedItems.Count + 1
                            If Len(shWorkSheet.Cell(j, i).GetFormattedString) > max_length Then max_length = Len(shWorkSheet.Cell(j, i).GetFormattedString)
                        Next
                        shWorkSheet.Columns(i).Width = OvertimeForm.CalculateColumnWidth(max_length)
                    Next

                    Dim range As IXLRange = shWorkSheet.Range(shWorkSheet.Cell(1, 1), shWorkSheet.Cell(OvertimeForm.select_employees.CheckedItems.Count + 1, Columns_list.SelectedIndices.Count))
                    range.Style.Border.LeftBorder = XLBorderStyleValues.Hair
                    range.Style.Border.RightBorder = XLBorderStyleValues.Hair
                    range.Style.Border.TopBorder = XLBorderStyleValues.Hair
                    range.Style.Border.BottomBorder = XLBorderStyleValues.Hair
                    range.Style.Border.InsideBorder = XLBorderStyleValues.Hair

                    For i = 0 To OvertimeForm.select_employees.CheckedItems.Count - 1
                        Dim sheetName As String
                        If Len(OvertimeForm.select_employees.CheckedItems(i).Text) > 31 Then
                            sheetName = Strings.Left(OvertimeForm.select_employees.CheckedItems(i).Text, 31)
                        Else
                            sheetName = OvertimeForm.select_employees.CheckedItems(i).Text
                        End If

                        shWorkSheet = bkWorkBook.Worksheets.Add(sheetName)
                        shWorkSheet.Range("A1").Value = "Datum"
                        shWorkSheet.Range("B1").Value = "Schichtart"
                        shWorkSheet.Range("C1").Value = "Start"
                        shWorkSheet.Range("D1").Value = "Ende"
                        shWorkSheet.Range("E1").Value = "Pause von"
                        shWorkSheet.Range("F1").Value = "Pause bis"
                        shWorkSheet.Range("G1").Value = "Pause von"
                        shWorkSheet.Range("H1").Value = "Pause bis"
                        shWorkSheet.Range("I1").Value = "IST"
                        shWorkSheet.Range("J1").Value = "Urlaubstage"

                        Dim row_ As Integer = 2

                        If OvertimeForm.employee_kw_hours.ContainsKey(sheetName) Then

                            For j = 0 To OvertimeForm.employee_kw_hours(sheetName).Count - 1
                                If j <> OvertimeForm.employee_kw_hours(sheetName).Count - 1 Then
                                    shWorkSheet.Range("A" & row_).Value = OvertimeForm.employee_kw_hours(sheetName)(j).Shift_date
                                    shWorkSheet.Range("B" & row_).Value = OvertimeForm.employee_kw_hours(sheetName)(j).Shift_type
                                    shWorkSheet.Range("C" & row_).Value = OvertimeForm.employee_kw_hours(sheetName)(j).Shift_from
                                    shWorkSheet.Range("C" & row_).Style.NumberFormat.Format = "hh:mm"
                                    shWorkSheet.Range("D" & row_).Value = OvertimeForm.employee_kw_hours(sheetName)(j).Shift_to
                                    shWorkSheet.Range("D" & row_).Style.NumberFormat.Format = "hh:mm"
                                    shWorkSheet.Range("E" & row_).Value = OvertimeForm.employee_kw_hours(sheetName)(j).Shift_break_1_from
                                    shWorkSheet.Range("E" & row_).Style.NumberFormat.Format = "hh:mm"
                                    shWorkSheet.Range("F" & row_).Value = OvertimeForm.employee_kw_hours(sheetName)(j).Shift_break_1_to
                                    shWorkSheet.Range("F" & row_).Style.NumberFormat.Format = "hh:mm"
                                    shWorkSheet.Range("G" & row_).Value = OvertimeForm.employee_kw_hours(sheetName)(j).Shift_break_2_from
                                    shWorkSheet.Range("G" & row_).Style.NumberFormat.Format = "hh:mm"
                                    shWorkSheet.Range("H" & row_).Value = OvertimeForm.employee_kw_hours(sheetName)(j).Shift_break_2_to
                                    shWorkSheet.Range("H" & row_).Style.NumberFormat.Format = "hh:mm"
                                    shWorkSheet.Range("I" & row_).Value = OvertimeForm.employee_kw_hours(sheetName)(j).Shift_length
                                    shWorkSheet.Range("I" & row_).Style.Font.FontColor = XLColor.Blue
                                Else
                                    shWorkSheet.Range("H" & row_).Value = "Urlaub:"
                                    shWorkSheet.Range("I" & row_).Value = OvertimeForm.employee_kw_hours(sheetName)(j).Shift_length
                                    shWorkSheet.Range("I" & row_).Style.Font.FontColor = XLColor.Blue
                                    shWorkSheet.Range("J" & row_).Value = OvertimeForm.employee_kw_hours(sheetName)(j).Holiday_days
                                End If
                                row_ = row_ + 1
                            Next

                            shWorkSheet.Range("A" & row_).Value = OvertimeForm.select_employees.CheckedItems(i).Text
                            shWorkSheet.Range("A" & row_).Style.Font.Bold = True
                            shWorkSheet.Range("H" & row_).Value = "TOTAL:"
                            shWorkSheet.Range("H" & row_).Style.Font.Bold = True
                            shWorkSheet.Range("I" & row_).FormulaA1 = "=SUM(I2:I" & row_ - 1 & ")"
                            shWorkSheet.Range("I" & row_).Style.Font.Bold = True
                            shWorkSheet.Range("A1:J1").Style.Fill.BackgroundColor = XLColor.Black
                            shWorkSheet.Range("A1:J1").Style.Font.FontColor = XLColor.White
                            shWorkSheet.Range("A1:J1").Style.Font.Bold = True
                            shWorkSheet.Columns("C:J").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right

                            shWorkSheet.Range("A1:J" & row_ - 1).Style.Border.LeftBorder = XLBorderStyleValues.Hair
                            shWorkSheet.Range("A1:J" & row_ - 1).Style.Border.RightBorder = XLBorderStyleValues.Hair
                            shWorkSheet.Range("A1:J" & row_ - 1).Style.Border.TopBorder = XLBorderStyleValues.Hair
                            shWorkSheet.Range("A1:J" & row_ - 1).Style.Border.BottomBorder = XLBorderStyleValues.Hair
                            shWorkSheet.Range("A1:J" & row_ - 1).Style.Border.InsideBorder = XLBorderStyleValues.Hair
                            For column = 1 To 10
                                Dim max_length As Integer = 0
                                For row = 1 To row_ - 1
                                    If Len(shWorkSheet.Cell(row, column).GetFormattedString) > max_length Then max_length = Len(shWorkSheet.Cell(row, column).GetFormattedString)
                                Next
                                shWorkSheet.Columns(column).Width = OvertimeForm.CalculateColumnWidth(max_length)
                            Next
                        End If
                    Next
                    bkWorkBook.SaveAs(export_folder & "\" & file_name)
                End If

                myWorkingForm.BeginInvoke(Sub() myWorkingForm.Close())
                Columns_list.DrawMode = DrawMode.Normal
            End If
            Columns_list.SelectedIndex = -1
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub ExportPDF_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim P As String = System.IO.Path.GetDirectoryName(Reflection.Assembly.GetEntryAssembly().Location)
        export_folder = P
        Label_folder.Text = P
        ToolTip1.SetToolTip(Label_folder, P)

        Columns_list.Items.Clear()

        With Columns_list
            For i = 0 To OvertimeForm.select_employees.Columns.Count - 1
                .Items.Add(OvertimeForm.select_employees.Columns(i).Text)
            Next
        End With
    End Sub

    Public Sub Label_folder_Click(sender As Object, e As EventArgs) Handles Label_folder.Click
        SelectFolder(Label_folder, ToolTip1)
    End Sub

    Public Sub SelectFolder(label As Label, tooltip As ToolTip)
        Dim FolderBrowser As FolderBrowserDialog = New FolderBrowserDialog
        FolderBrowser.RootFolder = Environment.SpecialFolder.MyComputer
        FolderBrowser.SelectedPath = export_folder
        FolderBrowser.ShowDialog()
        export_folder = FolderBrowser.SelectedPath
        label.Text = export_folder
        tooltip.SetToolTip(label, export_folder)
    End Sub


    Function GetDateTimeStamp()
        Dim strNow
        strNow = Now()
        GetDateTimeStamp = Pad2(DateAndTime.Day(strNow)) & "." & Pad2(DateAndTime.Month(strNow)) _
              & "." & DateAndTime.Year(strNow) & "_um_" & Pad2(DateAndTime.Hour(strNow)) _
              & "." & Pad2(DateAndTime.Minute(strNow))
    End Function

    Function Pad2(strIn As Integer) As String
        Dim Result As String = strIn
        If Len(Result) < 2 Then
            Result = "0" & Result
        End If
        Return Result
    End Function
    '~~> Release the objects
    Public Sub releaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
            GC.WaitForPendingFinalizers()
            GC.Collect()
            GC.WaitForPendingFinalizers()
        End Try
    End Sub

    Private Sub ExportPDF_LocationChanged(sender As Object, e As EventArgs) Handles Me.LocationChanged
        WorkingForm.Location = New Point(Location.X + Width / 2 - WorkingForm.Width / 2, Location.Y + Height / 2 - WorkingForm.Height / 2 + 52)
    End Sub
End Class