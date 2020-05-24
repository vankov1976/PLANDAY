Imports System.ComponentModel
Imports System.Threading
Imports ClosedXML.Excel
Imports PLANDAY.OvertimeForm

Public Class ExportForm

    Private Sub ExportForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ImChangingStuff = True

        If CreateGraphics().DpiX <= 96 Then
            normalDPI = True
        Else
            normalDPI = False
        End If

        Dim P As String = System.IO.Path.GetDirectoryName(Reflection.Assembly.GetEntryAssembly().Location)
        export_folder = P
        Label_folder.Text = P
        ToolTip1.SetToolTip(Label_folder, P)

        With select_month
            For Each monat In System.Globalization.CultureInfo.CurrentCulture.DateTimeFormat.MonthNames()
                If monat <> "" Then .Items.Add(monat)
            Next
        End With
        select_month.Text = Now.ToString("MMMM", System.Globalization.CultureInfo.CurrentCulture)

        OvertimeForm.init_payroll_periods()
        YEAR_button.Text = payroll_periods.Last.Key
        OvertimeForm.set_params()
        OvertimeForm.initialize()

        With select_status
            .Items.Add("ALL")
            .Items.Add("Active")
            .Items.Add("Deactivated")
        End With
        select_status.Text = "ALL"

        If ownDepartment = "" Then ownDepartment = OvertimeForm.get_ownDepartment()

        With select_hotel
            .Columns.Add("Hotel", width:=276)
            .Columns.Add("Nr", width:=74)
            For Each department_id In departments.Keys
                If ownDepartment = "NONE" Then Exit For
                If ownDepartment <> "ALL" Then
                    .Items.Add(ownDepartment).SubItems.Add(department_numbers(departments.FirstOrDefault(Function(x) x.Value = ownDepartment).Key))
                    Exit For
                End If
                Dim item As ListViewItem = .Items.Add(departments(department_id))
                item.SubItems.Add(department_numbers(department_id))
            Next
        End With

        If normalDPI Then
            select_hotel.BeginUpdate()
            For i = 0 To select_hotel.Columns.Count - 1
                Dim colPercentage As Double = Convert.ToInt32(select_hotel.Columns(i).Width) / 390
                select_hotel.Columns(i).Width = CInt(colPercentage * select_hotel.Width)
            Next
            select_hotel.EndUpdate()
        End If

        ''finish loader''
        myWorkingForm.BeginInvoke(Sub() myWorkingForm.Close())

        ImChangingStuff = False
    End Sub

    Public Sub Loader(x As Integer, y As Integer, width As Integer, height As Integer, normalDPI As Boolean)
        If normalDPI Then
            myWorkingForm = New WorkingForm()
        Else
            myWorkingForm = New LoadingForm()
        End If
        myWorkingForm.StartPosition = FormStartPosition.Manual
        myWorkingForm.Location = New Point(x + width / 2 - myWorkingForm.Width / 2, y + height / 2 - myWorkingForm.Height / 2 + If(normalDPI, 92, 120))
        System.Windows.Forms.Application.Run(myWorkingForm)
    End Sub


    Private Sub START_button_Click(sender As Object, e As EventArgs) Handles START_button.Click
        ImChangingStuff = True
        start()
        ImChangingStuff = False
    End Sub

    Private Sub start()

        For i = 0 To select_hotel.SelectedItems.Count - 1
            If selected_departments Is Nothing Then
                ReDim Preserve selected_departments(0)
                selected_departments(0) = OvertimeForm.getDepartmentID(select_hotel.SelectedItems(i).Text)
            Else
                ReDim Preserve selected_departments(UBound(selected_departments) + 1)
                selected_departments(UBound(selected_departments)) = OvertimeForm.getDepartmentID(select_hotel.SelectedItems(i).Text)
            End If
        Next i

        If selected_departments Is Nothing Then
            MsgBox("Please select department")
            Exit Sub
        End If

        OvertimeForm.init_FT(YEAR_button.Text)
        OvertimeForm.current_year = YEAR_button.Text

        select_hotel.OwnerDraw = True
        select_hotel.Refresh()

        Select Case select_status.Text
            Case "ALL"
                export_ALL()
            Case "Deactivated"
                export_deactivated()
            Case "Active"
                export_active()
        End Select

        select_hotel.SelectedIndices.Clear()
        select_hotel.OwnerDraw = False

    End Sub

    Sub export_ALL()

        ''''loader'''
        Dim th As System.Threading.Thread = New Threading.Thread(Sub() Loader(Me.Location.X, Me.Location.Y, Me.Width, Me.Height, normalDPI))
        th.SetApartmentState(ApartmentState.STA)
        th.Start()
        '''''''''''''
        Dim Parsed As Dictionary(Of String, Object)
        Dim value As Dictionary(Of String, Object)
        Dim objhttp As New WinHttp.WinHttpRequest

        '''''''employee_records'''''''''

        '''''''''''''''get employees_count''''''''''''''
        objhttp.Open("GET", "https://openapi.planday.com/hr/v1.0/employees?limit=0", False)
        objhttp.SetRequestHeader("Authorization", "Bearer " & api_token)
        objhttp.SetRequestHeader("X-ClientId", "ddca428b-8530-405d-9960-047132c49531")
        objhttp.Send()
        Parsed = New Web.Script.Serialization.JavaScriptSerializer().Deserialize(Of Object)(objhttp.ResponseText)

        Dim employees_count As Integer
        If Parsed("paging").ContainsKey("total") Then employees_count = Parsed("paging")("total")

        For y = 0 To CInt(employees_count / 50)

            objhttp.Open("GET", "https://openapi.planday.com/hr/v1.0/employees?limit=0&offset=" & y * 50 & "&special=BirthDate", False)
            objhttp.SetRequestHeader("Authorization", "Bearer " & api_token)
            objhttp.SetRequestHeader("X-ClientId", "ddca428b-8530-405d-9960-047132c49531")
            objhttp.Send()
            Parsed = New Web.Script.Serialization.JavaScriptSerializer().Deserialize(Of Object)(objhttp.ResponseText)


            For Each value In Parsed("data")
                For Each Values In value("departments")
                    If selected_departments.Contains(Values) And value("userName") <> "markus.pettinger@meininger-hotels.com" And value("userName") <> "daniela.simicguel@meininger-hotels.com" Then
                        'find total count
                        BarTotalCount = BarTotalCount + 1
                    End If
                Next
            Next
        Next


        Dim found()
        ''''find count'''''
        For Each Item In selected_departments
            objhttp.Open("GET", "https://openapi.planday.com/payroll/v1.0/payroll?departmentIds=" & Item & "&from=" & start_ & "&to=" & end_ & "&shiftStatus=Approved", False)
            objhttp.SetRequestHeader("Authorization", "Bearer " & api_token)
            objhttp.SetRequestHeader("X-ClientId", "ddca428b-8530-405d-9960-047132c49531")
            objhttp.Send()
            Parsed = New Web.Script.Serialization.JavaScriptSerializer().Deserialize(Of Object)(objhttp.ResponseText)

            For Each value In Parsed("shiftsPayroll")

                If deactivated.Contains(value("employeeId")) Then
                    If found Is Nothing Then
                        ReDim Preserve found(0)
                        found(0) = value("employeeId")
                    Else
                        If Not found.Contains(value("employeeId")) Then
                            ReDim Preserve found(UBound(found) + 1)
                            found(UBound(found)) = value("employeeId")
                        End If
                    End If

                End If
            Next

        Next


        BarTotalCount = BarTotalCount + If(found Is Nothing, 0, found.Length)

        ''finish loader''
        myWorkingForm.BeginInvoke(Sub() myWorkingForm.Close())

        If BarTotalCount > 0 Then
            ''''start backgroundworker
            ''''Initialize ProgressBar
            ProgressBarForm.TotalCount = BarTotalCount
            bgw_export_ALL.RunWorkerAsync()
            ProgressBarForm.ShowDialog(Me)
        Else
            MsgBox("No records found!")
            Exit Sub
        End If

    End Sub

    Private Sub bgw_export_ALL_DoWork(sender As Object, e As DoWorkEventArgs) Handles bgw_export_ALL.DoWork

        Dim bkWorkBook As XLWorkbook
        Dim shWorkSheet As IXLWorksheet
        Dim workbook_names()
        Dim found()
        Dim quartal As String
        Dim row As Integer
        Dim Parsed As Dictionary(Of String, Object)
        Dim Parsed_temp As Dictionary(Of String, Object)
        Dim value As Dictionary(Of String, Object)
        Dim objhttp As New WinHttp.WinHttpRequest
        Dim current_row As New Dictionary(Of String, Integer)

        report_workbooks.Clear()
        '''''''employee_records'''''''''

        '''''''''''''''get employees_count''''''''''''''
        objhttp.Open("GET", "https://openapi.planday.com/hr/v1.0/employees?limit=0", False)
        objhttp.SetRequestHeader("Authorization", "Bearer " & api_token)
        objhttp.SetRequestHeader("X-ClientId", "ddca428b-8530-405d-9960-047132c49531")
        objhttp.Send()
        Parsed = New Web.Script.Serialization.JavaScriptSerializer().Deserialize(Of Object)(objhttp.ResponseText)

        Dim employees_count As Integer
        If Parsed("paging").ContainsKey("total") Then employees_count = Parsed("paging")("total")

        For y = 0 To CInt(employees_count / 50)

            objhttp.Open("GET", "https://openapi.planday.com/hr/v1.0/employees?limit=0&offset=" & y * 50 & "&special=BirthDate", False)
            objhttp.SetRequestHeader("Authorization", "Bearer " & api_token)
            objhttp.SetRequestHeader("X-ClientId", "ddca428b-8530-405d-9960-047132c49531")
            objhttp.Send()
            Parsed = New Web.Script.Serialization.JavaScriptSerializer().Deserialize(Of Object)(objhttp.ResponseText)


            For Each value In Parsed("data")
                For Each Values In value("departments")
                    If selected_departments.Contains(Values) And value("userName") <> "markus.pettinger@meininger-hotels.com" And value("userName") <> "daniela.simicguel@meininger-hotels.com" Then


                        Dim department_name As String
                        department_name = Replace(departments(Values), " ", "_")
                        department_name = Replace(department_name, "/", "_")
                        department_name = Replace(department_name, """", "")
                        Dim hotel As String
                        Dim department_number As String = department_numbers(Values)
                        hotel = department_number & "_" & department_name & name_extension & ".xlsx"

                        '''''''
                        If Not current_row.ContainsKey(hotel) Then
                            If workbook_names Is Nothing Then
                                ReDim Preserve workbook_names(0)
                                workbook_names(0) = hotel
                            Else
                                ReDim Preserve workbook_names(UBound(workbook_names) + 1)
                                workbook_names(UBound(workbook_names)) = hotel
                            End If

                            row = 2

                            bkWorkBook = New XLWorkbook()
                            shWorkSheet = bkWorkBook.Worksheets.Add("Exported data")

                            Dim table As IXLTable

                            If selected_month = 3 Or
                               selected_month = 6 Or
                               selected_month = 9 Or
                               selected_month = 12 Then
                                table = shWorkSheet.Range("$A$1:$AE$51").CreateTable("export")
                            Else
                                table = shWorkSheet.Range("$A$1:$AB$51").CreateTable("export")
                            End If
                            table.ShowTotalsRow = True
                            table.Field(0).TotalsRowLabel = "Total"

                            shWorkSheet.Range("A1").Value = "Hotel"
                            shWorkSheet.Range("B1").Value = "ADP Nr."
                            shWorkSheet.Range("C1").Value = "First name"
                            shWorkSheet.Range("D1").Value = "Last name"
                            shWorkSheet.Range("E1").Value = "Employee type"
                            shWorkSheet.Range("F1").Value = "Contract rule"
                            shWorkSheet.Range("G1").Value = "Start night"
                            shWorkSheet.Range("H1").Value = "End night"
                            shWorkSheet.Range("I1").Value = "Sunday paycode"
                            shWorkSheet.Range("J1").Value = "Sunday surcharge"
                            shWorkSheet.Range("K1").Value = "Sunday hours"
                            shWorkSheet.Range("L1").Value = "Sunday payment"
                            shWorkSheet.Range("M1").Value = "Night paycode"
                            shWorkSheet.Range("N1").Value = "Night surcharge"
                            shWorkSheet.Range("O1").Value = "Night hours"
                            shWorkSheet.Range("P1").Value = "Night payment"
                            shWorkSheet.Range("Q1").Value = "105 hours"
                            shWorkSheet.Range("R1").Value = "All shifts approved"
                            shWorkSheet.Range("S1").Value = "420 payment"
                            shWorkSheet.Range("T1").Value = "420 Sunday hours"
                            shWorkSheet.Range("U1").Value = "420 Night hours"
                            shWorkSheet.Range("V1").Value = "Paid leave sunday hours"
                            shWorkSheet.Range("W1").Value = "Sickness sunday hours"
                            shWorkSheet.Range("X1").Value = "Paid leave night hours"
                            shWorkSheet.Range("Y1").Value = "Sickness night hours"
                            shWorkSheet.Range("Z1").Value = "Last 3 months Sunday hours"
                            shWorkSheet.Range("AA1").Value = "Last 3 months Night hours"
                            shWorkSheet.Range("AB1").Value = "Paid leave days"

                            If selected_month = 3 Or
                               selected_month = 6 Or
                               selected_month = 9 Or
                               selected_month = 12 Then

                                Select Case selected_month
                                    Case 3
                                        quartal = "I"
                                    Case 6
                                        quartal = "II"
                                    Case 9
                                        quartal = "III"
                                    Case 12
                                        quartal = "IV"
                                End Select

                                shWorkSheet.Range("AC1").Value = quartal & " quartal working days"
                                shWorkSheet.Range("AD1").Value = quartal & " quartal working hours"
                                shWorkSheet.Range("AE1").Value = "paid leave in hours"
                                table.DataRange.Column(29).Style.NumberFormat.Format = "General;;"
                                table.DataRange.Column(30).Style.NumberFormat.Format = "General;;"
                                table.DataRange.Column(31).Style.NumberFormat.Format = "General;;"
                                table.DataRange.Column(29).Style.Border.LeftBorder = XLBorderStyleValues.Thin
                                table.DataRange.Column(29).Style.Border.RightBorder = XLBorderStyleValues.Thin
                                table.DataRange.Column(30).Style.Border.LeftBorder = XLBorderStyleValues.Thin
                                table.DataRange.Column(30).Style.Border.RightBorder = XLBorderStyleValues.Thin
                                table.DataRange.Column(31).Style.Border.LeftBorder = XLBorderStyleValues.Thin
                                table.DataRange.Column(31).Style.Border.RightBorder = XLBorderStyleValues.Thin
                            End If

                            table.DataRange.Column(7).Style.NumberFormat.Format = "hh:mm"
                            table.DataRange.Column(8).Style.NumberFormat.Format = "hh:mm"
                            table.DataRange.Column(10).Style.NumberFormat.Format = "#,##0.00 [$€-de-DE]"
                            table.DataRange.Column(11).Style.NumberFormat.Format = "General;;"
                            table.DataRange.Column(12).Style.NumberFormat.Format = "#,##0.00 [$€-de-DE]"
                            table.DataRange.Column(14).Style.NumberFormat.Format = "#,##0.00 [$€-de-DE]"
                            table.DataRange.Column(15).Style.NumberFormat.Format = "General;;"
                            table.DataRange.Column(16).Style.NumberFormat.Format = "#,##0.00 [$€-de-DE]"
                            table.DataRange.Column(17).Style.NumberFormat.Format = "General;;"
                            table.DataRange.Column(19).Style.NumberFormat.Format = "#,##0.00 [$€-de-DE]"
                            table.DataRange.Column(20).Style.NumberFormat.Format = "0.00;;"
                            table.DataRange.Column(21).Style.NumberFormat.Format = "0.00;;"
                            table.DataRange.Column(22).Style.NumberFormat.Format = "0.00;;"
                            table.DataRange.Column(23).Style.NumberFormat.Format = "General;;"
                            table.DataRange.Column(24).Style.NumberFormat.Format = "0.00;;"
                            table.DataRange.Column(25).Style.NumberFormat.Format = "General;;"
                            table.DataRange.Column(26).Style.NumberFormat.Format = "General;;"
                            table.DataRange.Column(27).Style.NumberFormat.Format = "General;;"
                            table.DataRange.Column(28).Style.NumberFormat.Format = "General;;"
                            OvertimeForm.format_yellow(table.DataRange.Column(12).AsRange)
                            OvertimeForm.format_yellow(table.DataRange.Column(16).AsRange)
                            OvertimeForm.format_yellow(table.DataRange.Column(19).AsRange)
                            table.DataRange.Column(12).Style.Font.Bold = True
                            table.DataRange.Column(16).Style.Font.Bold = True
                            table.DataRange.Column(19).Style.Font.Bold = True
                            OvertimeForm.format_blue(table.DataRange.Column(11).AsRange)
                            OvertimeForm.format_blue(table.DataRange.Column(15).AsRange)
                            OvertimeForm.format_blue(table.DataRange.Column(20).AsRange)
                            OvertimeForm.format_blue(table.DataRange.Column(21).AsRange)
                            table.DataRange.Column(18).Style.Border.LeftBorder = XLBorderStyleValues.Thin
                            table.DataRange.Column(18).Style.Border.RightBorder = XLBorderStyleValues.Thin
                            table.Theme = XLTableTheme.TableStyleLight8

                            OvertimeForm.format_white_borders(table.HeadersRow().Cells())

                            shWorkSheet.ShowGridLines = False
                            shWorkSheet.SheetView.Freeze(1, 4)

                            report_workbooks(hotel) = bkWorkBook
                        End If
                        '''''''

                        sunday_FT_hours = 0
                        night_hours = 0
                        working_hours = 0
                        working_hours_krank = 0
                        sunday_FT_hours_krank = 0
                        night_hours_krank = 0
                        last_3_months_sunday_FT_hours = 0
                        last_3_months_night_hours = 0
                        quarter_working_days = 0
                        quarter_working_hours = 0
                        Dim azubi As Boolean = False

                        'Do Action
                        CurrentAction = CurrentAction + 1
                        ActionName = "Processing " & Trim(value("firstName")) & " " & Trim(value("lastName")) & "... "
                        'Report current status in %
                        bgw_export_ALL.ReportProgress(100 * CurrentAction / BarTotalCount)

                        Dim employee_type As String = employee_types(value("employeeTypeId"))

                        objhttp.Open("GET", "https://openapi.planday.com/contractrules/v1.0/employees/" & value("id"), False)
                        objhttp.SetRequestHeader("Authorization", "Bearer " & api_token)
                        objhttp.SetRequestHeader("X-ClientId", "ddca428b-8530-405d-9960-047132c49531")
                        objhttp.Send()
                        Parsed_temp = New Web.Script.Serialization.JavaScriptSerializer().Deserialize(Of Object)(objhttp.ResponseText)

                        Dim id_contract_rule As Integer
                        If objhttp.ResponseText = "{}" Then
                            id_contract_rule = 3527
                        Else
                            id_contract_rule = Parsed_temp("data")("id")
                        End If


                        objhttp.Open("GET", "https://openapi.planday.com/hr/v1.0/employees/" & value("id"), False)
                        objhttp.SetRequestHeader("Authorization", "Bearer " & api_token)
                        objhttp.SetRequestHeader("X-ClientId", "ddca428b-8530-405d-9960-047132c49531")
                        objhttp.Send()
                        Parsed_temp = New Web.Script.Serialization.JavaScriptSerializer().Deserialize(Of Object)(objhttp.ResponseText)

                        Dim senior As Boolean
                        If Parsed_temp("data").ContainsKey("custom_221267") Then
                            If Parsed_temp("data")("custom_221267")("value") = "20:00" Then
                                senior = True
                            Else
                                senior = False
                            End If
                        Else
                            senior = False
                        End If

                        For Each temp_values In value("employeeGroups")
                            If employee_groups.ContainsKey(temp_values) Then
                                If employee_groups(temp_values) = "Azubis" Then
                                    azubi = True
                                    Exit For
                                Else
                                    azubi = False
                                End If
                            End If
                        Next

                        current_employee = Trim(value("firstName")) & " " & Trim(value("lastName"))
                        If Len(current_employee) > 31 Then
                            current_employee = Strings.Left(current_employee, 31)
                        End If
                        current_workbook = hotel
                        OvertimeForm.calculate_last_3_months(value("id"), Values, senior, If(value.ContainsKey("birthDate"), value("birthDate"), ""))
                        OvertimeForm.calculate_All(value("id"), Values, senior, If(value.ContainsKey("birthDate"), value("birthDate"), ""))

                        sunday_FT_hours = Math.Round(sunday_FT_hours, 2)
                        night_hours = Math.Round(night_hours, 2)
                        working_hours = Math.Round(working_hours, 2)
                        working_hours_krank = Math.Round(working_hours_krank, 2)
                        sunday_FT_hours_krank = Math.Round(sunday_FT_hours_krank, 2)
                        night_hours_krank = Math.Round(night_hours_krank, 2)
                        last_3_months_sunday_FT_hours = Math.Round(last_3_months_sunday_FT_hours, 2)
                        last_3_months_night_hours = Math.Round(last_3_months_night_hours, 2)

                        If current_row.ContainsKey(hotel) Then row = current_row.Item(hotel)

                        If azubi Then
                            OvertimeForm.format_grey(report_workbooks(hotel).Worksheet("Exported data").Range("A" & row & ":J" & row))
                            OvertimeForm.format_grey(report_workbooks(hotel).Worksheet("Exported data").Range("M" & row & ":N" & row))
                        End If
                        report_workbooks(hotel).Worksheet("Exported data").Cell(row, 1).Value = department_number
                        report_workbooks(hotel).Worksheet("Exported data").Cell(row, 2).Value = value("salaryIdentifier")
                        report_workbooks(hotel).Worksheet("Exported data").Cell(row, 3).Value = Trim(value("firstName"))
                        report_workbooks(hotel).Worksheet("Exported data").Cell(row, 4).Value = Trim(value("lastName"))
                        report_workbooks(hotel).Worksheet("Exported data").Cell(row, 5).Value = employee_type
                        report_workbooks(hotel).Worksheet("Exported data").Cell(row, 6).Value = contract_rules(id_contract_rule)
                        If senior Then
                            report_workbooks(hotel).Worksheet("Exported data").Cell(row, 7).Value = 20 / 24
                        Else
                            report_workbooks(hotel).Worksheet("Exported data").Cell(row, 7).Value = 22 / 24
                        End If
                        report_workbooks(hotel).Worksheet("Exported data").Cell(row, 8).Value = 6 / 24
                        If azubi Then
                            report_workbooks(hotel).Worksheet("Exported data").Cell(row, 9).Value = 413
                        Else
                            report_workbooks(hotel).Worksheet("Exported data").Cell(row, 9).Value = 411
                        End If
                        If azubi Then
                            report_workbooks(hotel).Worksheet("Exported data").Cell(row, 10).Value = 2.5
                        Else
                            report_workbooks(hotel).Worksheet("Exported data").Cell(row, 10).Value = 4
                        End If
                        report_workbooks(hotel).Worksheet("Exported data").Cell(row, 11).Value = sunday_FT_hours
                        If azubi Then
                            report_workbooks(hotel).Worksheet("Exported data").Cell(row, 13).Value = 412
                        Else
                            report_workbooks(hotel).Worksheet("Exported data").Cell(row, 13).Value = 410
                        End If
                        If azubi Then
                            report_workbooks(hotel).Worksheet("Exported data").Cell(row, 14).Value = 1.25
                        Else
                            report_workbooks(hotel).Worksheet("Exported data").Cell(row, 14).Value = 2
                        End If
                        report_workbooks(hotel).Worksheet("Exported data").Cell(row, 15).Value = night_hours
                        If id_contract_rule = 3527 Or id_contract_rule = 4856 Then
                            report_workbooks(hotel).Worksheet("Exported data").Cell(row, 17).Value = working_hours + working_hours_krank
                            OvertimeForm.format_blue(report_workbooks(hotel).Worksheet("Exported data").Cell(row, 17).AsRange)
                        End If ''''
                        If OvertimeForm.all_shifts_approved(value("id"), Values) Then
                            report_workbooks(hotel).Worksheet("Exported data").Cell(row, 18).Value = "YES"
                            report_workbooks(hotel).Worksheet("Exported data").Cell(row, 18).Style.Fill.BackgroundColor = XLColor.Green
                        Else
                            report_workbooks(hotel).Worksheet("Exported data").Cell(row, 18).Value = "NO"
                            report_workbooks(hotel).Worksheet("Exported data").Cell(row, 18).Style.Fill.BackgroundColor = XLColor.Red
                        End If
                        report_workbooks(hotel).Worksheet("Exported data").Cell(row, 23).Value = sunday_FT_hours_krank
                        report_workbooks(hotel).Worksheet("Exported data").Cell(row, 25).Value = night_hours_krank
                        report_workbooks(hotel).Worksheet("Exported data").Cell(row, 26).Value = last_3_months_sunday_FT_hours
                        report_workbooks(hotel).Worksheet("Exported data").Cell(row, 27).Value = last_3_months_night_hours
                        report_workbooks(hotel).Worksheet("Exported data").Cell(row, 28).Value = OvertimeForm.paid_leave_days(value("id"))
                        If selected_month = 3 Or
                           selected_month = 6 Or
                           selected_month = 9 Or
                           selected_month = 12 Then

                            OvertimeForm.calculate_quarter(value("id"), Values, senior, If(value.ContainsKey("birthDate"), value("birthDate"), ""))
                            quarter_working_hours = Math.Round(quarter_working_hours, 2)

                            If id_contract_rule = 3527 Or id_contract_rule = 4856 Then
                                report_workbooks(hotel).Worksheet("Exported data").Cell(row, 29).Value = quarter_working_days
                                report_workbooks(hotel).Worksheet("Exported data").Cell(row, 30).Value = quarter_working_hours
                                OvertimeForm.format_blue_noBorders(report_workbooks(hotel).Worksheet("Exported data").Cell(row, 29).AsRange)
                                OvertimeForm.format_blue_noBorders(report_workbooks(hotel).Worksheet("Exported data").Cell(row, 30).AsRange)
                                OvertimeForm.format_blue_noBorders(report_workbooks(hotel).Worksheet("Exported data").Cell(row, 31).AsRange)
                            End If
                        End If

                        row = row + 1
                        current_row(hotel) = row
                    End If

                Next
            Next

        Next

        '''''''''process all deactivated'''''''''''''
        For Each Item In selected_departments

            objhttp.Open("GET", "https://openapi.planday.com/payroll/v1.0/payroll?departmentIds=" & Item & "&from=" & start_ & "&to=" & end_ & "&shiftStatus=Approved", False)
            objhttp.SetRequestHeader("Authorization", "Bearer " & api_token)
            objhttp.SetRequestHeader("X-ClientId", "ddca428b-8530-405d-9960-047132c49531")
            objhttp.Send()
            Parsed = New Web.Script.Serialization.JavaScriptSerializer().Deserialize(Of Object)(objhttp.ResponseText)

            Erase found
            For Each value In Parsed("shiftsPayroll")

                If deactivated.Contains(value("employeeId")) Then
                    If found Is Nothing Then
                        ReDim Preserve found(0)
                        found(0) = value("employeeId")
                    Else
                        If Not found.Contains(value("employeeId")) Then
                            ReDim Preserve found(UBound(found) + 1)
                            found(UBound(found)) = value("employeeId")
                        End If
                    End If

                End If
            Next

            Dim department_name As String
            department_name = Replace(departments(Item), " ", "_")
            department_name = Replace(department_name, "/", "_")
            department_name = Replace(department_name, """", "")

            Dim hotel As String
            Dim department_number As String = department_numbers(Item)
            hotel = department_number & "_" & department_name & name_extension & ".xlsx"

            If Not found Is Nothing Then
                If current_row.ContainsKey(hotel) Then row = current_row(hotel)

                For Each item_found In found

                    objhttp.Open("GET", "https://openapi.planday.com/hr/v1.0/employees/" & item_found & "?special=BirthDate", False)
                    objhttp.SetRequestHeader("Authorization", "Bearer " & api_token)
                    objhttp.SetRequestHeader("X-ClientId", "ddca428b-8530-405d-9960-047132c49531")
                    objhttp.Send()
                    Parsed = New Web.Script.Serialization.JavaScriptSerializer().Deserialize(Of Object)(objhttp.ResponseText)

                    sunday_FT_hours = 0
                    night_hours = 0
                    working_hours = 0
                    working_hours_krank = 0
                    sunday_FT_hours_krank = 0
                    night_hours_krank = 0
                    last_3_months_sunday_FT_hours = 0
                    last_3_months_night_hours = 0
                    quarter_working_days = 0
                    quarter_working_hours = 0
                    Dim azubi As Boolean = False

                    'Do Action
                    CurrentAction = CurrentAction + 1
                    ActionName = "Processing " & Trim(Parsed("data")("firstName")) & " " & Trim(Parsed("data")("lastName")) & "... "
                    'Report current status in %
                    bgw_export_ALL.ReportProgress(100 * CurrentAction / BarTotalCount)

                    Dim employee_type As String = employee_types(Parsed("data")("employeeTypeId"))

                    objhttp.Open("GET", "https://openapi.planday.com/contractrules/v1.0/employees/" & item_found, False)
                    objhttp.SetRequestHeader("Authorization", "Bearer " & api_token)
                    objhttp.SetRequestHeader("X-ClientId", "ddca428b-8530-405d-9960-047132c49531")
                    objhttp.Send()
                    Parsed_temp = New Web.Script.Serialization.JavaScriptSerializer().Deserialize(Of Object)(objhttp.ResponseText)

                    Dim id_contract_rule As Integer
                    If objhttp.ResponseText = "{}" Then
                        id_contract_rule = 3527
                    Else
                        id_contract_rule = Parsed_temp("data")("id")
                    End If

                    objhttp.Open("GET", "https://openapi.planday.com/hr/v1.0/employees/" & item_found, False)
                    objhttp.SetRequestHeader("Authorization", "Bearer " & api_token)
                    objhttp.SetRequestHeader("X-ClientId", "ddca428b-8530-405d-9960-047132c49531")
                    objhttp.Send()
                    Parsed_temp = New Web.Script.Serialization.JavaScriptSerializer().Deserialize(Of Object)(objhttp.ResponseText)

                    Dim senior As Boolean
                    If Parsed_temp("data").ContainsKey("custom_221267") Then
                        If Parsed_temp("data")("custom_221267")("value") = "20:00" Then
                            senior = True
                        Else
                            senior = False
                        End If
                    Else
                        senior = False
                    End If

                    For Each temp_values In Parsed("data")("employeeGroups")
                        If employee_groups.ContainsKey(temp_values) Then
                            If employee_groups(temp_values) = "Azubis" Then
                                azubi = True
                                Exit For
                            Else
                                azubi = False
                            End If
                        End If
                    Next

                    current_employee = Trim(Parsed("data")("firstName")) & " " & Trim(Parsed("data")("lastName"))
                    If Len(current_employee) > 31 Then
                        current_employee = Strings.Left(current_employee, 31)
                    End If
                    current_workbook = hotel
                    OvertimeForm.calculate_last_3_months(item_found, Item, senior, If(Parsed("data").ContainsKey("birthDate"), Parsed("data")("birthDate"), ""))
                    OvertimeForm.calculate_All(item_found, Item, senior, If(Parsed("data").ContainsKey("birthDate"), Parsed("data")("birthDate"), ""))


                    sunday_FT_hours = Math.Round(sunday_FT_hours, 2)
                    night_hours = Math.Round(night_hours, 2)
                    working_hours = Math.Round(working_hours, 2)
                    working_hours_krank = Math.Round(working_hours_krank, 2)
                    sunday_FT_hours_krank = Math.Round(sunday_FT_hours_krank, 2)
                    night_hours_krank = Math.Round(night_hours_krank, 2)
                    last_3_months_sunday_FT_hours = Math.Round(last_3_months_sunday_FT_hours, 2)
                    last_3_months_night_hours = Math.Round(last_3_months_night_hours, 2)

                    If azubi Then
                        OvertimeForm.format_grey(report_workbooks(hotel).Worksheet("Exported data").Range("A" & row & ":J" & row))
                        OvertimeForm.format_grey(report_workbooks(hotel).Worksheet("Exported data").Range("M" & row & ":N" & row))
                    End If
                    report_workbooks(hotel).Worksheet("Exported data").Cell(row, 1).Value = department_number
                    report_workbooks(hotel).Worksheet("Exported data").Cell(row, 2).Value = Parsed("data")("salaryIdentifier")
                    report_workbooks(hotel).Worksheet("Exported data").Cell(row, 3).Value = Trim(Parsed("data")("firstName"))
                    report_workbooks(hotel).Worksheet("Exported data").Cell(row, 4).Value = Trim(Parsed("data")("lastName"))
                    report_workbooks(hotel).Worksheet("Exported data").Cell(row, 5).Value = employee_type
                    report_workbooks(hotel).Worksheet("Exported data").Cell(row, 6).Value = contract_rules(id_contract_rule)
                    If senior Then
                        report_workbooks(hotel).Worksheet("Exported data").Cell(row, 7).Value = 20 / 24
                    Else
                        report_workbooks(hotel).Worksheet("Exported data").Cell(row, 7).Value = 22 / 24
                    End If
                    report_workbooks(hotel).Worksheet("Exported data").Cell(row, 8).Value = 6 / 24
                    If azubi Then
                        report_workbooks(hotel).Worksheet("Exported data").Cell(row, 9).Value = 413
                    Else
                        report_workbooks(hotel).Worksheet("Exported data").Cell(row, 9).Value = 411
                    End If
                    If azubi Then
                        report_workbooks(hotel).Worksheet("Exported data").Cell(row, 10).Value = 2.5
                    Else
                        report_workbooks(hotel).Worksheet("Exported data").Cell(row, 10).Value = 4
                    End If
                    report_workbooks(hotel).Worksheet("Exported data").Cell(row, 11).Value = sunday_FT_hours
                    If azubi Then
                        report_workbooks(hotel).Worksheet("Exported data").Cell(row, 13).Value = 412
                    Else
                        report_workbooks(hotel).Worksheet("Exported data").Cell(row, 13).Value = 410
                    End If
                    If azubi Then
                        report_workbooks(hotel).Worksheet("Exported data").Cell(row, 14).Value = 1.25
                    Else
                        report_workbooks(hotel).Worksheet("Exported data").Cell(row, 14).Value = 2
                    End If
                    report_workbooks(hotel).Worksheet("Exported data").Cell(row, 15).Value = night_hours
                    If id_contract_rule = 3527 Or id_contract_rule = 4856 Then
                        report_workbooks(hotel).Worksheet("Exported data").Cell(row, 17).Value = working_hours + working_hours_krank
                        OvertimeForm.format_blue(report_workbooks(hotel).Worksheet("Exported data").Cell(row, 17).AsRange)
                    End If
                    If OvertimeForm.all_shifts_approved(item_found, Item) Then
                        report_workbooks(hotel).Worksheet("Exported data").Cell(row, 18).Value = "YES"
                        report_workbooks(hotel).Worksheet("Exported data").Cell(row, 18).Style.Fill.BackgroundColor = XLColor.Green
                    Else
                        report_workbooks(hotel).Worksheet("Exported data").Cell(row, 18).Value = "NO"
                        report_workbooks(hotel).Worksheet("Exported data").Cell(row, 18).Style.Fill.BackgroundColor = XLColor.Red
                    End If
                    report_workbooks(hotel).Worksheet("Exported data").Cell(row, 23).Value = sunday_FT_hours_krank
                    report_workbooks(hotel).Worksheet("Exported data").Cell(row, 25).Value = night_hours_krank
                    report_workbooks(hotel).Worksheet("Exported data").Cell(row, 26).Value = last_3_months_sunday_FT_hours
                    report_workbooks(hotel).Worksheet("Exported data").Cell(row, 27).Value = last_3_months_night_hours
                    report_workbooks(hotel).Worksheet("Exported data").Cell(row, 28).Value = OvertimeForm.paid_leave_days(item_found)
                    If selected_month = 3 Or
                       selected_month = 6 Or
                       selected_month = 9 Or
                       selected_month = 12 Then

                        OvertimeForm.calculate_quarter(item_found, Item, senior, If(Parsed("data").ContainsKey("birthDate"), Parsed("data")("birthDate"), ""))
                        quarter_working_hours = Math.Round(quarter_working_hours, 2)

                        If id_contract_rule = 3527 Or id_contract_rule = 4856 Then
                            report_workbooks(hotel).Worksheet("Exported data").Cell(row, 29).Value = quarter_working_days
                            report_workbooks(hotel).Worksheet("Exported data").Cell(row, 30).Value = quarter_working_hours
                            OvertimeForm.format_blue_noBorders(report_workbooks(hotel).Worksheet("Exported data").Cell(row, 29).AsRange)
                            OvertimeForm.format_blue_noBorders(report_workbooks(hotel).Worksheet("Exported data").Cell(row, 30).AsRange)
                            OvertimeForm.format_blue_noBorders(report_workbooks(hotel).Worksheet("Exported data").Cell(row, 31).AsRange)
                        End If
                    End If

                    row = row + 1
                    current_row(hotel) = row
                Next
            End If

        Next

        For Each workbook_name In workbook_names

            report_workbooks(workbook_name).Worksheet("Exported data").Table("export").HeadersRow.Style.Alignment.WrapText = True
            For i = 3 To 5
                Dim max_length As Integer = 0
                For Each cell In report_workbooks(workbook_name).Worksheet("Exported data").Table("export").DataRange.Column(i).Cells
                    If Len(cell.GetFormattedString) > max_length Then max_length = Len(cell.GetFormattedString)
                Next
                report_workbooks(workbook_name).Worksheet("Exported data").Columns(i).Width = OvertimeForm.CalculateColumnWidth(max_length)
            Next
            report_workbooks(workbook_name).Worksheet("Exported data").Columns("A:A").Width = 7
            report_workbooks(workbook_name).Worksheet("Exported data").Columns("B:B").Width = 9
            report_workbooks(workbook_name).Worksheet("Exported data").Columns("F:H").Width = 7.7
            report_workbooks(workbook_name).Worksheet("Exported data").Columns("I:J").Width = 11
            report_workbooks(workbook_name).Worksheet("Exported data").Columns("K:L").Width = 10
            report_workbooks(workbook_name).Worksheet("Exported data").Columns("M:N").Width = 11
            report_workbooks(workbook_name).Worksheet("Exported data").Columns("O:P").Width = 10
            report_workbooks(workbook_name).Worksheet("Exported data").Columns("Q:Q").Width = 7.7
            report_workbooks(workbook_name).Worksheet("Exported data").Columns("R:U").Width = 11
            report_workbooks(workbook_name).Worksheet("Exported data").Columns("V:AA").Width = 16
            report_workbooks(workbook_name).Worksheet("Exported data").Columns("AB:AB").Width = 11
            report_workbooks(workbook_name).Worksheet("Exported data").Columns("AC:AE").Width = 16.5

            report_workbooks(workbook_name).Worksheet("Exported data").Table("export").SetAutoFilter.Column(1).NotEqualTo(Of String)("")
            report_workbooks(workbook_name).Worksheet("Exported data").Table("export").AutoFilter.Sort(3, XLSortOrder.Ascending)
            Dim formula_row As Integer

            formula_row = 2
            For Each cell In report_workbooks(workbook_name).Worksheet("Exported data").Table("export").DataRange.Column(12).Cells
                cell.FormulaA1 = "=IF(OR(C" & formula_row & "="""",K" & formula_row & "=0),"""",J" & formula_row & "*K" & formula_row & ")"
                formula_row = formula_row + 1
            Next

            formula_row = 2
            For Each cell In report_workbooks(workbook_name).Worksheet("Exported data").Table("export").DataRange.Column(16).Cells
                cell.FormulaA1 = "=IF(OR(C" & formula_row & "="""",O" & formula_row & "=0),"""",N" & formula_row & "*O" & formula_row & ")"
                formula_row = formula_row + 1
            Next

            formula_row = 2
            For Each cell In report_workbooks(workbook_name).Worksheet("Exported data").Table("export").DataRange.Column(19).Cells
                cell.FormulaA1 = "=IF(OR(C" & formula_row & "="""",AND(T" & formula_row & "=0,U" & formula_row & "=0)),"""",T" & formula_row & "*J" & formula_row & "+U" & formula_row & "*N" & formula_row & ")"
                formula_row = formula_row + 1
            Next

            formula_row = 2
            For Each cell In report_workbooks(workbook_name).Worksheet("Exported data").Table("export").DataRange.Column(22).Cells
                cell.FormulaA1 = "=IF(C" & formula_row & "="""","""",Z" & formula_row & "/65*(AB" & formula_row & "))"
                formula_row = formula_row + 1
            Next

            formula_row = 2
            For Each cell In report_workbooks(workbook_name).Worksheet("Exported data").Table("export").DataRange.Column(24).Cells
                cell.FormulaA1 = "=IF(C" & formula_row & "="""","""",AA" & formula_row & "/65*(AB" & formula_row & "))"
                formula_row = formula_row + 1
            Next

            formula_row = 2
            For Each cell In report_workbooks(workbook_name).Worksheet("Exported data").Table("export").DataRange.Column(20).Cells
                cell.FormulaA1 = "=IF(C" & formula_row & "="""","""",V" & formula_row & "+W" & formula_row & ")"
                formula_row = formula_row + 1
            Next

            formula_row = 2
            For Each cell In report_workbooks(workbook_name).Worksheet("Exported data").Table("export").DataRange.Column(21).Cells
                cell.FormulaA1 = "=IF(C" & formula_row & "="""","""",X" & formula_row & "+Y" & formula_row & ")"
                formula_row = formula_row + 1
            Next

            If report_workbooks(workbook_name).Worksheet("Exported data").Table("export").Columns.Count = 31 Then
                formula_row = 2
                For Each cell In report_workbooks(workbook_name).Worksheet("Exported data").Table("export").DataRange.Column(31).Cells
                    cell.FormulaA1 = "=IFERROR(ROUND((AD" & formula_row & "/AC" & formula_row & ")*(6.25*((AC" & formula_row & "/13)/5)),2),0)"
                    formula_row = formula_row + 1
                Next
            End If

            report_workbooks(workbook_name).SaveAs(export_folder & "\" & workbook_name)
        Next
    End Sub

    Private Sub bgw_export_ALL_ProgressChanged(sender As Object, e As ProgressChangedEventArgs) Handles bgw_export_ALL.ProgressChanged
        ProgressBarForm.NextAction(CurrentAction, ActionName)
    End Sub

    Private Sub bgw_export_ALL_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles bgw_export_ALL.RunWorkerCompleted
        CurrentAction = 0
        BarTotalCount = 0
        ProgressBarForm.Complete()
        Erase selected_departments
    End Sub

    Sub export_active()

        ''''loader'''
        Dim th As System.Threading.Thread = New Threading.Thread(Sub() Loader(Me.Location.X, Me.Location.Y, Me.Width, Me.Height, normalDPI))
        th.SetApartmentState(ApartmentState.STA)
        th.Start()
        '''''''''''''
        Dim Parsed As Dictionary(Of String, Object)
        Dim value As Dictionary(Of String, Object)
        Dim objhttp As New WinHttp.WinHttpRequest

        '''''''employee_records'''''''''

        '''''''''''''''get employees_count''''''''''''''
        objhttp.Open("GET", "https://openapi.planday.com/hr/v1.0/employees?limit=0", False)
        objhttp.SetRequestHeader("Authorization", "Bearer " & api_token)
        objhttp.SetRequestHeader("X-ClientId", "ddca428b-8530-405d-9960-047132c49531")
        objhttp.Send()
        Parsed = New Web.Script.Serialization.JavaScriptSerializer().Deserialize(Of Object)(objhttp.ResponseText)

        Dim employees_count As Integer
        If Parsed("paging").ContainsKey("total") Then employees_count = Parsed("paging")("total")

        For y = 0 To CInt(employees_count / 50)

            objhttp.Open("GET", "https://openapi.planday.com/hr/v1.0/employees?limit=0&offset=" & y * 50 & "&special=BirthDate", False)
            objhttp.SetRequestHeader("Authorization", "Bearer " & api_token)
            objhttp.SetRequestHeader("X-ClientId", "ddca428b-8530-405d-9960-047132c49531")
            objhttp.Send()
            Parsed = New Web.Script.Serialization.JavaScriptSerializer().Deserialize(Of Object)(objhttp.ResponseText)


            For Each value In Parsed("data")
                For Each Values In value("departments")
                    If selected_departments.Contains(Values) And value("userName") <> "markus.pettinger@meininger-hotels.com" And value("userName") <> "daniela.simicguel@meininger-hotels.com" Then
                        'find total count
                        BarTotalCount = BarTotalCount + 1
                    End If
                Next
            Next
        Next

        ''finish loader''
        myWorkingForm.BeginInvoke(Sub() myWorkingForm.Close())

        If BarTotalCount > 0 Then
            ''''start backgroundworker
            ''''Initialize ProgressBar
            ProgressBarForm.TotalCount = BarTotalCount
            bgw_export_active.RunWorkerAsync()
            ProgressBarForm.ShowDialog(Me)
        Else
            MsgBox("No records found!")
            Exit Sub
        End If

    End Sub

    Private Sub bgw_export_active_DoWork(sender As Object, e As DoWorkEventArgs) Handles bgw_export_active.DoWork

        Dim bkWorkBook As XLWorkbook
        Dim shWorkSheet As IXLWorksheet
        Dim workbook_names()
        Dim quartal As String
        Dim row As Integer
        Dim Parsed As Dictionary(Of String, Object)
        Dim Parsed_temp As Dictionary(Of String, Object)
        Dim value As Dictionary(Of String, Object)
        Dim objhttp As New WinHttp.WinHttpRequest
        Dim current_row As New Dictionary(Of String, Integer)

        report_workbooks.Clear()
        '''''''employee_records'''''''''

        '''''''''''''''get employees_count''''''''''''''
        objhttp.Open("GET", "https://openapi.planday.com/hr/v1.0/employees?limit=0", False)
        objhttp.SetRequestHeader("Authorization", "Bearer " & api_token)
        objhttp.SetRequestHeader("X-ClientId", "ddca428b-8530-405d-9960-047132c49531")
        objhttp.Send()
        Parsed = New Web.Script.Serialization.JavaScriptSerializer().Deserialize(Of Object)(objhttp.ResponseText)

        Dim employees_count As Integer
        If Parsed("paging").ContainsKey("total") Then employees_count = Parsed("paging")("total")

        For y = 0 To CInt(employees_count / 50)

            objhttp.Open("GET", "https://openapi.planday.com/hr/v1.0/employees?limit=0&offset=" & y * 50 & "&special=BirthDate", False)
            objhttp.SetRequestHeader("Authorization", "Bearer " & api_token)
            objhttp.SetRequestHeader("X-ClientId", "ddca428b-8530-405d-9960-047132c49531")
            objhttp.Send()
            Parsed = New Web.Script.Serialization.JavaScriptSerializer().Deserialize(Of Object)(objhttp.ResponseText)


            For Each value In Parsed("data")
                For Each Values In value("departments")
                    If selected_departments.Contains(Values) And value("userName") <> "markus.pettinger@meininger-hotels.com" And value("userName") <> "daniela.simicguel@meininger-hotels.com" Then


                        Dim department_name As String
                        department_name = Replace(departments(Values), " ", "_")
                        department_name = Replace(department_name, "/", "_")
                        department_name = Replace(department_name, """", "")
                        Dim hotel As String
                        Dim department_number As String = department_numbers(Values)
                        hotel = department_number & "_" & department_name & name_extension & "_ACTIVE.xlsx"

                        '''''''
                        If Not current_row.ContainsKey(hotel) Then
                            If workbook_names Is Nothing Then
                                ReDim Preserve workbook_names(0)
                                workbook_names(0) = hotel
                            Else
                                ReDim Preserve workbook_names(UBound(workbook_names) + 1)
                                workbook_names(UBound(workbook_names)) = hotel
                            End If

                            row = 2

                            bkWorkBook = New XLWorkbook()
                            shWorkSheet = bkWorkBook.Worksheets.Add("Exported data")

                            Dim table As IXLTable

                            If selected_month = 3 Or
                               selected_month = 6 Or
                               selected_month = 9 Or
                               selected_month = 12 Then
                                table = shWorkSheet.Range("$A$1:$AE$51").CreateTable("export")
                            Else
                                table = shWorkSheet.Range("$A$1:$AB$51").CreateTable("export")
                            End If
                            table.ShowTotalsRow = True
                            table.Field(0).TotalsRowLabel = "Total"

                            shWorkSheet.Range("A1").Value = "Hotel"
                            shWorkSheet.Range("B1").Value = "ADP Nr."
                            shWorkSheet.Range("C1").Value = "First name"
                            shWorkSheet.Range("D1").Value = "Last name"
                            shWorkSheet.Range("E1").Value = "Employee type"
                            shWorkSheet.Range("F1").Value = "Contract rule"
                            shWorkSheet.Range("G1").Value = "Start night"
                            shWorkSheet.Range("H1").Value = "End night"
                            shWorkSheet.Range("I1").Value = "Sunday paycode"
                            shWorkSheet.Range("J1").Value = "Sunday surcharge"
                            shWorkSheet.Range("K1").Value = "Sunday hours"
                            shWorkSheet.Range("L1").Value = "Sunday payment"
                            shWorkSheet.Range("M1").Value = "Night paycode"
                            shWorkSheet.Range("N1").Value = "Night surcharge"
                            shWorkSheet.Range("O1").Value = "Night hours"
                            shWorkSheet.Range("P1").Value = "Night payment"
                            shWorkSheet.Range("Q1").Value = "105 hours"
                            shWorkSheet.Range("R1").Value = "All shifts approved"
                            shWorkSheet.Range("S1").Value = "420 payment"
                            shWorkSheet.Range("T1").Value = "420 Sunday hours"
                            shWorkSheet.Range("U1").Value = "420 Night hours"
                            shWorkSheet.Range("V1").Value = "Paid leave sunday hours"
                            shWorkSheet.Range("W1").Value = "Sickness sunday hours"
                            shWorkSheet.Range("X1").Value = "Paid leave night hours"
                            shWorkSheet.Range("Y1").Value = "Sickness night hours"
                            shWorkSheet.Range("Z1").Value = "Last 3 months Sunday hours"
                            shWorkSheet.Range("AA1").Value = "Last 3 months Night hours"
                            shWorkSheet.Range("AB1").Value = "Paid leave days"

                            If selected_month = 3 Or
                               selected_month = 6 Or
                               selected_month = 9 Or
                               selected_month = 12 Then

                                Select Case selected_month
                                    Case 3
                                        quartal = "I"
                                    Case 6
                                        quartal = "II"
                                    Case 9
                                        quartal = "III"
                                    Case 12
                                        quartal = "IV"
                                End Select

                                shWorkSheet.Range("AC1").Value = quartal & " quartal working days"
                                shWorkSheet.Range("AD1").Value = quartal & " quartal working hours"
                                shWorkSheet.Range("AE1").Value = "paid leave in hours"
                                table.DataRange.Column(29).Style.NumberFormat.Format = "General;;"
                                table.DataRange.Column(30).Style.NumberFormat.Format = "General;;"
                                table.DataRange.Column(31).Style.NumberFormat.Format = "General;;"
                                table.DataRange.Column(29).Style.Border.LeftBorder = XLBorderStyleValues.Thin
                                table.DataRange.Column(29).Style.Border.RightBorder = XLBorderStyleValues.Thin
                                table.DataRange.Column(30).Style.Border.LeftBorder = XLBorderStyleValues.Thin
                                table.DataRange.Column(30).Style.Border.RightBorder = XLBorderStyleValues.Thin
                                table.DataRange.Column(31).Style.Border.LeftBorder = XLBorderStyleValues.Thin
                                table.DataRange.Column(31).Style.Border.RightBorder = XLBorderStyleValues.Thin
                            End If

                            table.DataRange.Column(7).Style.NumberFormat.Format = "hh:mm"
                            table.DataRange.Column(8).Style.NumberFormat.Format = "hh:mm"
                            table.DataRange.Column(10).Style.NumberFormat.Format = "#,##0.00 [$€-de-DE]"
                            table.DataRange.Column(11).Style.NumberFormat.Format = "General;;"
                            table.DataRange.Column(12).Style.NumberFormat.Format = "#,##0.00 [$€-de-DE]"
                            table.DataRange.Column(14).Style.NumberFormat.Format = "#,##0.00 [$€-de-DE]"
                            table.DataRange.Column(15).Style.NumberFormat.Format = "General;;"
                            table.DataRange.Column(16).Style.NumberFormat.Format = "#,##0.00 [$€-de-DE]"
                            table.DataRange.Column(17).Style.NumberFormat.Format = "General;;"
                            table.DataRange.Column(19).Style.NumberFormat.Format = "#,##0.00 [$€-de-DE]"
                            table.DataRange.Column(20).Style.NumberFormat.Format = "0.00;;"
                            table.DataRange.Column(21).Style.NumberFormat.Format = "0.00;;"
                            table.DataRange.Column(22).Style.NumberFormat.Format = "0.00;;"
                            table.DataRange.Column(23).Style.NumberFormat.Format = "General;;"
                            table.DataRange.Column(24).Style.NumberFormat.Format = "0.00;;"
                            table.DataRange.Column(25).Style.NumberFormat.Format = "General;;"
                            table.DataRange.Column(26).Style.NumberFormat.Format = "General;;"
                            table.DataRange.Column(27).Style.NumberFormat.Format = "General;;"
                            table.DataRange.Column(28).Style.NumberFormat.Format = "General;;"
                            OvertimeForm.format_yellow(table.DataRange.Column(12).AsRange)
                            OvertimeForm.format_yellow(table.DataRange.Column(16).AsRange)
                            OvertimeForm.format_yellow(table.DataRange.Column(19).AsRange)
                            table.DataRange.Column(12).Style.Font.Bold = True
                            table.DataRange.Column(16).Style.Font.Bold = True
                            table.DataRange.Column(19).Style.Font.Bold = True
                            OvertimeForm.format_blue(table.DataRange.Column(11).AsRange)
                            OvertimeForm.format_blue(table.DataRange.Column(15).AsRange)
                            OvertimeForm.format_blue(table.DataRange.Column(20).AsRange)
                            OvertimeForm.format_blue(table.DataRange.Column(21).AsRange)
                            table.DataRange.Column(18).Style.Border.LeftBorder = XLBorderStyleValues.Thin
                            table.DataRange.Column(18).Style.Border.RightBorder = XLBorderStyleValues.Thin
                            table.Theme = XLTableTheme.TableStyleLight8

                            OvertimeForm.format_white_borders(table.HeadersRow().Cells())

                            shWorkSheet.ShowGridLines = False
                            shWorkSheet.SheetView.Freeze(1, 4)

                            report_workbooks(hotel) = bkWorkBook
                        End If
                        '''''''

                        sunday_FT_hours = 0
                        night_hours = 0
                        working_hours = 0
                        working_hours_krank = 0
                        sunday_FT_hours_krank = 0
                        night_hours_krank = 0
                        last_3_months_sunday_FT_hours = 0
                        last_3_months_night_hours = 0
                        quarter_working_days = 0
                        quarter_working_hours = 0
                        Dim azubi As Boolean = False

                        'Do Action
                        CurrentAction = CurrentAction + 1
                        ActionName = "Processing " & Trim(value("firstName")) & " " & Trim(value("lastName")) & "... "
                        'Report current status in %
                        bgw_export_active.ReportProgress(100 * CurrentAction / BarTotalCount)

                        Dim employee_type As String = employee_types(value("employeeTypeId"))

                        objhttp.Open("GET", "https://openapi.planday.com/contractrules/v1.0/employees/" & value("id"), False)
                        objhttp.SetRequestHeader("Authorization", "Bearer " & api_token)
                        objhttp.SetRequestHeader("X-ClientId", "ddca428b-8530-405d-9960-047132c49531")
                        objhttp.Send()
                        Parsed_temp = New Web.Script.Serialization.JavaScriptSerializer().Deserialize(Of Object)(objhttp.ResponseText)

                        Dim id_contract_rule As Integer
                        If objhttp.ResponseText = "{}" Then
                            id_contract_rule = 3527
                        Else
                            id_contract_rule = Parsed_temp("data")("id")
                        End If


                        objhttp.Open("GET", "https://openapi.planday.com/hr/v1.0/employees/" & value("id"), False)
                        objhttp.SetRequestHeader("Authorization", "Bearer " & api_token)
                        objhttp.SetRequestHeader("X-ClientId", "ddca428b-8530-405d-9960-047132c49531")
                        objhttp.Send()
                        Parsed_temp = New Web.Script.Serialization.JavaScriptSerializer().Deserialize(Of Object)(objhttp.ResponseText)

                        Dim senior As Boolean
                        If Parsed_temp("data").ContainsKey("custom_221267") Then
                            If Parsed_temp("data")("custom_221267")("value") = "20:00" Then
                                senior = True
                            Else
                                senior = False
                            End If
                        Else
                            senior = False
                        End If

                        For Each temp_values In value("employeeGroups")
                            If employee_groups.ContainsKey(temp_values) Then
                                If employee_groups(temp_values) = "Azubis" Then
                                    azubi = True
                                    Exit For
                                Else
                                    azubi = False
                                End If
                            End If
                        Next

                        current_employee = Trim(value("firstName")) & " " & Trim(value("lastName"))
                        If Len(current_employee) > 31 Then
                            current_employee = Strings.Left(current_employee, 31)
                        End If
                        current_workbook = hotel
                        OvertimeForm.calculate_last_3_months(value("id"), Values, senior, If(value.ContainsKey("birthDate"), value("birthDate"), ""))
                        OvertimeForm.calculate_All(value("id"), Values, senior, If(value.ContainsKey("birthDate"), value("birthDate"), ""))

                        sunday_FT_hours = Math.Round(sunday_FT_hours, 2)
                        night_hours = Math.Round(night_hours, 2)
                        working_hours = Math.Round(working_hours, 2)
                        working_hours_krank = Math.Round(working_hours_krank, 2)
                        sunday_FT_hours_krank = Math.Round(sunday_FT_hours_krank, 2)
                        night_hours_krank = Math.Round(night_hours_krank, 2)
                        last_3_months_sunday_FT_hours = Math.Round(last_3_months_sunday_FT_hours, 2)
                        last_3_months_night_hours = Math.Round(last_3_months_night_hours, 2)

                        If current_row.ContainsKey(hotel) Then row = current_row.Item(hotel)

                        If azubi Then
                            OvertimeForm.format_grey(report_workbooks(hotel).Worksheet("Exported data").Range("A" & row & ":J" & row))
                            OvertimeForm.format_grey(report_workbooks(hotel).Worksheet("Exported data").Range("M" & row & ":N" & row))
                        End If
                        report_workbooks(hotel).Worksheet("Exported data").Cell(row, 1).Value = department_number
                        report_workbooks(hotel).Worksheet("Exported data").Cell(row, 2).Value = value("salaryIdentifier")
                        report_workbooks(hotel).Worksheet("Exported data").Cell(row, 3).Value = Trim(value("firstName"))
                        report_workbooks(hotel).Worksheet("Exported data").Cell(row, 4).Value = Trim(value("lastName"))
                        report_workbooks(hotel).Worksheet("Exported data").Cell(row, 5).Value = employee_type
                        report_workbooks(hotel).Worksheet("Exported data").Cell(row, 6).Value = contract_rules(id_contract_rule)
                        If senior Then
                            report_workbooks(hotel).Worksheet("Exported data").Cell(row, 7).Value = 20 / 24
                        Else
                            report_workbooks(hotel).Worksheet("Exported data").Cell(row, 7).Value = 22 / 24
                        End If
                        report_workbooks(hotel).Worksheet("Exported data").Cell(row, 8).Value = 6 / 24
                        If azubi Then
                            report_workbooks(hotel).Worksheet("Exported data").Cell(row, 9).Value = 413
                        Else
                            report_workbooks(hotel).Worksheet("Exported data").Cell(row, 9).Value = 411
                        End If
                        If azubi Then
                            report_workbooks(hotel).Worksheet("Exported data").Cell(row, 10).Value = 2.5
                        Else
                            report_workbooks(hotel).Worksheet("Exported data").Cell(row, 10).Value = 4
                        End If
                        report_workbooks(hotel).Worksheet("Exported data").Cell(row, 11).Value = sunday_FT_hours
                        If azubi Then
                            report_workbooks(hotel).Worksheet("Exported data").Cell(row, 13).Value = 412
                        Else
                            report_workbooks(hotel).Worksheet("Exported data").Cell(row, 13).Value = 410
                        End If
                        If azubi Then
                            report_workbooks(hotel).Worksheet("Exported data").Cell(row, 14).Value = 1.25
                        Else
                            report_workbooks(hotel).Worksheet("Exported data").Cell(row, 14).Value = 2
                        End If
                        report_workbooks(hotel).Worksheet("Exported data").Cell(row, 15).Value = night_hours
                        If id_contract_rule = 3527 Or id_contract_rule = 4856 Then
                            report_workbooks(hotel).Worksheet("Exported data").Cell(row, 17).Value = working_hours + working_hours_krank
                            OvertimeForm.format_blue(report_workbooks(hotel).Worksheet("Exported data").Cell(row, 17).AsRange)
                        End If ''''
                        If OvertimeForm.all_shifts_approved(value("id"), Values) Then
                            report_workbooks(hotel).Worksheet("Exported data").Cell(row, 18).Value = "YES"
                            report_workbooks(hotel).Worksheet("Exported data").Cell(row, 18).Style.Fill.BackgroundColor = XLColor.Green
                        Else
                            report_workbooks(hotel).Worksheet("Exported data").Cell(row, 18).Value = "NO"
                            report_workbooks(hotel).Worksheet("Exported data").Cell(row, 18).Style.Fill.BackgroundColor = XLColor.Red
                        End If
                        report_workbooks(hotel).Worksheet("Exported data").Cell(row, 23).Value = sunday_FT_hours_krank
                        report_workbooks(hotel).Worksheet("Exported data").Cell(row, 25).Value = night_hours_krank
                        report_workbooks(hotel).Worksheet("Exported data").Cell(row, 26).Value = last_3_months_sunday_FT_hours
                        report_workbooks(hotel).Worksheet("Exported data").Cell(row, 27).Value = last_3_months_night_hours
                        report_workbooks(hotel).Worksheet("Exported data").Cell(row, 28).Value = OvertimeForm.paid_leave_days(value("id"))
                        If selected_month = 3 Or
                           selected_month = 6 Or
                           selected_month = 9 Or
                           selected_month = 12 Then

                            OvertimeForm.calculate_quarter(value("id"), Values, senior, If(value.ContainsKey("birthDate"), value("birthDate"), ""))
                            quarter_working_hours = Math.Round(quarter_working_hours, 2)

                            If id_contract_rule = 3527 Or id_contract_rule = 4856 Then
                                report_workbooks(hotel).Worksheet("Exported data").Cell(row, 29).Value = quarter_working_days
                                report_workbooks(hotel).Worksheet("Exported data").Cell(row, 30).Value = quarter_working_hours
                                OvertimeForm.format_blue_noBorders(report_workbooks(hotel).Worksheet("Exported data").Cell(row, 29).AsRange)
                                OvertimeForm.format_blue_noBorders(report_workbooks(hotel).Worksheet("Exported data").Cell(row, 30).AsRange)
                                OvertimeForm.format_blue_noBorders(report_workbooks(hotel).Worksheet("Exported data").Cell(row, 31).AsRange)
                            End If
                        End If

                        row = row + 1
                        current_row(hotel) = row
                    End If

                Next
            Next

        Next

        For Each workbook_name In workbook_names

            report_workbooks(workbook_name).Worksheet("Exported data").Table("export").HeadersRow.Style.Alignment.WrapText = True
            For i = 3 To 5
                Dim max_length As Integer = 0
                For Each cell In report_workbooks(workbook_name).Worksheet("Exported data").Table("export").DataRange.Column(i).Cells
                    If Len(cell.GetFormattedString) > max_length Then max_length = Len(cell.GetFormattedString)
                Next
                report_workbooks(workbook_name).Worksheet("Exported data").Columns(i).Width = OvertimeForm.CalculateColumnWidth(max_length)
            Next
            report_workbooks(workbook_name).Worksheet("Exported data").Columns("A:A").Width = 7
            report_workbooks(workbook_name).Worksheet("Exported data").Columns("B:B").Width = 9
            report_workbooks(workbook_name).Worksheet("Exported data").Columns("F:H").Width = 7.7
            report_workbooks(workbook_name).Worksheet("Exported data").Columns("I:J").Width = 11
            report_workbooks(workbook_name).Worksheet("Exported data").Columns("K:L").Width = 10
            report_workbooks(workbook_name).Worksheet("Exported data").Columns("M:N").Width = 11
            report_workbooks(workbook_name).Worksheet("Exported data").Columns("O:P").Width = 10
            report_workbooks(workbook_name).Worksheet("Exported data").Columns("Q:Q").Width = 7.7
            report_workbooks(workbook_name).Worksheet("Exported data").Columns("R:U").Width = 11
            report_workbooks(workbook_name).Worksheet("Exported data").Columns("V:AA").Width = 16
            report_workbooks(workbook_name).Worksheet("Exported data").Columns("AB:AB").Width = 11
            report_workbooks(workbook_name).Worksheet("Exported data").Columns("AC:AE").Width = 16.5

            report_workbooks(workbook_name).Worksheet("Exported data").Table("export").SetAutoFilter.Column(1).NotEqualTo(Of String)("")
            report_workbooks(workbook_name).Worksheet("Exported data").Table("export").AutoFilter.Sort(3, XLSortOrder.Ascending)
            Dim formula_row As Integer

            formula_row = 2
            For Each cell In report_workbooks(workbook_name).Worksheet("Exported data").Table("export").DataRange.Column(12).Cells
                cell.FormulaA1 = "=IF(OR(C" & formula_row & "="""",K" & formula_row & "=0),"""",J" & formula_row & "*K" & formula_row & ")"
                formula_row = formula_row + 1
            Next

            formula_row = 2
            For Each cell In report_workbooks(workbook_name).Worksheet("Exported data").Table("export").DataRange.Column(16).Cells
                cell.FormulaA1 = "=IF(OR(C" & formula_row & "="""",O" & formula_row & "=0),"""",N" & formula_row & "*O" & formula_row & ")"
                formula_row = formula_row + 1
            Next

            formula_row = 2
            For Each cell In report_workbooks(workbook_name).Worksheet("Exported data").Table("export").DataRange.Column(19).Cells
                cell.FormulaA1 = "=IF(OR(C" & formula_row & "="""",AND(T" & formula_row & "=0,U" & formula_row & "=0)),"""",T" & formula_row & "*J" & formula_row & "+U" & formula_row & "*N" & formula_row & ")"
                formula_row = formula_row + 1
            Next

            formula_row = 2
            For Each cell In report_workbooks(workbook_name).Worksheet("Exported data").Table("export").DataRange.Column(22).Cells
                cell.FormulaA1 = "=IF(C" & formula_row & "="""","""",Z" & formula_row & "/65*(AB" & formula_row & "))"
                formula_row = formula_row + 1
            Next

            formula_row = 2
            For Each cell In report_workbooks(workbook_name).Worksheet("Exported data").Table("export").DataRange.Column(24).Cells
                cell.FormulaA1 = "=IF(C" & formula_row & "="""","""",AA" & formula_row & "/65*(AB" & formula_row & "))"
                formula_row = formula_row + 1
            Next

            formula_row = 2
            For Each cell In report_workbooks(workbook_name).Worksheet("Exported data").Table("export").DataRange.Column(20).Cells
                cell.FormulaA1 = "=IF(C" & formula_row & "="""","""",V" & formula_row & "+W" & formula_row & ")"
                formula_row = formula_row + 1
            Next

            formula_row = 2
            For Each cell In report_workbooks(workbook_name).Worksheet("Exported data").Table("export").DataRange.Column(21).Cells
                cell.FormulaA1 = "=IF(C" & formula_row & "="""","""",X" & formula_row & "+Y" & formula_row & ")"
                formula_row = formula_row + 1
            Next

            If report_workbooks(workbook_name).Worksheet("Exported data").Table("export").Columns.Count = 31 Then
                formula_row = 2
                For Each cell In report_workbooks(workbook_name).Worksheet("Exported data").Table("export").DataRange.Column(31).Cells
                    cell.FormulaA1 = "=IFERROR(ROUND((AD" & formula_row & "/AC" & formula_row & ")*(6.25*((AC" & formula_row & "/13)/5)),2),0)"
                    formula_row = formula_row + 1
                Next
            End If

            report_workbooks(workbook_name).SaveAs(export_folder & "\" & workbook_name)
        Next

    End Sub

    Private Sub bgw_export_active_ProgressChanged(sender As Object, e As ProgressChangedEventArgs) Handles bgw_export_active.ProgressChanged
        ProgressBarForm.NextAction(CurrentAction, ActionName)
    End Sub

    Private Sub bgw_export_active_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles bgw_export_active.RunWorkerCompleted
        CurrentAction = 0
        BarTotalCount = 0
        ProgressBarForm.Complete()
        Erase selected_departments
    End Sub

    Sub export_deactivated()

        ''''loader'''
        Dim th As System.Threading.Thread = New Threading.Thread(Sub() Loader(Me.Location.X, Me.Location.Y, Me.Width, Me.Height, normalDPI))
        th.SetApartmentState(ApartmentState.STA)
        th.Start()
        '''''''''''''
        Dim Parsed As Dictionary(Of String, Object)
        Dim value As Dictionary(Of String, Object)
        Dim objhttp As New WinHttp.WinHttpRequest

        Dim found()
        ''''find count'''''
        For Each Item In selected_departments

            objhttp.Open("GET", "https://openapi.planday.com/payroll/v1.0/payroll?departmentIds=" & Item & "&from=" & start_ & "&to=" & end_ & "&shiftStatus=Approved", False)
            objhttp.SetRequestHeader("Authorization", "Bearer " & api_token)
            objhttp.SetRequestHeader("X-ClientId", "ddca428b-8530-405d-9960-047132c49531")
            objhttp.Send()
            Parsed = New Web.Script.Serialization.JavaScriptSerializer().Deserialize(Of Object)(objhttp.ResponseText)

            For Each value In Parsed("shiftsPayroll")

                If deactivated.Contains(value("employeeId")) Then
                    If found Is Nothing Then
                        ReDim Preserve found(0)
                        found(0) = value("employeeId")
                    Else
                        If Not found.Contains(value("employeeId")) Then
                            ReDim Preserve found(UBound(found) + 1)
                            found(UBound(found)) = value("employeeId")
                        End If
                    End If

                End If
            Next

        Next


        BarTotalCount = If(found Is Nothing, 0, found.Length)

        ''finish loader''
        myWorkingForm.BeginInvoke(Sub() myWorkingForm.Close())

        If BarTotalCount > 0 Then
            ''''start backgroundworker
            ''''Initialize ProgressBar
            ProgressBarForm.TotalCount = BarTotalCount
            bgw_export_deactivated.RunWorkerAsync()
            ProgressBarForm.ShowDialog(Me)
        Else
            MsgBox("No records found!")
            Exit Sub
        End If

    End Sub

    Private Sub bgw_export_deactivated_DoWork(sender As Object, e As DoWorkEventArgs) Handles bgw_export_deactivated.DoWork

        Dim bkWorkBook As XLWorkbook
        Dim shWorkSheet As IXLWorksheet
        Dim workbook_names()
        Dim found()
        Dim quartal As String
        Dim row As Integer
        Dim Parsed As Dictionary(Of String, Object)
        Dim Parsed_temp As Dictionary(Of String, Object)
        Dim value As Dictionary(Of String, Object)
        Dim objhttp As New WinHttp.WinHttpRequest 'CreateObject("MSXML2.ServerXMLHTTP")
        Dim current_row As New Dictionary(Of String, Integer)

        report_workbooks.Clear()
        '''''''''process all deactivated'''''''''''''
        For Each Item In selected_departments

            objhttp.open("GET", "https://openapi.planday.com/payroll/v1.0/payroll?departmentIds=" & Item & "&from=" & start_ & "&to=" & end_ & "&shiftStatus=Approved", False)
            objhttp.setRequestHeader("Authorization", "Bearer " & api_token)
            objhttp.setRequestHeader("X-ClientId", "ddca428b-8530-405d-9960-047132c49531")
            objhttp.send()
            Parsed = New Web.Script.Serialization.JavaScriptSerializer().Deserialize(Of Object)(objhttp.responseText)

            Erase found
            For Each value In Parsed("shiftsPayroll")

                If deactivated.Contains(value("employeeId")) Then
                    If found Is Nothing Then
                        ReDim Preserve found(0)
                        found(0) = value("employeeId")
                    Else
                        If Not found.Contains(value("employeeId")) Then
                            ReDim Preserve found(UBound(found) + 1)
                            found(UBound(found)) = value("employeeId")
                        End If
                    End If

                End If
            Next

            Dim department_name As String
            department_name = Replace(departments(Item), " ", "_")
            department_name = Replace(department_name, "/", "_")
            department_name = Replace(department_name, """", "")

            Dim hotel As String
            Dim department_number As String = department_numbers(Item)
            hotel = department_number & "_" & department_name & name_extension & "_DEACTIVATED.xlsx"

            If Not found Is Nothing Then
                '''''''
                If Not current_row.ContainsKey(hotel) Then
                    If workbook_names Is Nothing Then
                        ReDim Preserve workbook_names(0)
                        workbook_names(0) = hotel
                    Else
                        ReDim Preserve workbook_names(UBound(workbook_names) + 1)
                        workbook_names(UBound(workbook_names)) = hotel
                    End If

                    row = 2

                    bkWorkBook = New XLWorkbook()
                    shWorkSheet = bkWorkBook.Worksheets.Add("Exported data")

                    Dim table As IXLTable

                    If selected_month = 3 Or
                       selected_month = 6 Or
                       selected_month = 9 Or
                       selected_month = 12 Then
                        table = shWorkSheet.Range("$A$1:$AE$51").CreateTable("export")
                    Else
                        table = shWorkSheet.Range("$A$1:$AB$51").CreateTable("export")
                    End If
                    table.ShowTotalsRow = True
                    table.Field(0).TotalsRowLabel = "Total"

                    shWorkSheet.Range("A1").Value = "Hotel"
                    shWorkSheet.Range("B1").Value = "ADP Nr."
                    shWorkSheet.Range("C1").Value = "First name"
                    shWorkSheet.Range("D1").Value = "Last name"
                    shWorkSheet.Range("E1").Value = "Employee type"
                    shWorkSheet.Range("F1").Value = "Contract rule"
                    shWorkSheet.Range("G1").Value = "Start night"
                    shWorkSheet.Range("H1").Value = "End night"
                    shWorkSheet.Range("I1").Value = "Sunday paycode"
                    shWorkSheet.Range("J1").Value = "Sunday surcharge"
                    shWorkSheet.Range("K1").Value = "Sunday hours"
                    shWorkSheet.Range("L1").Value = "Sunday payment"
                    shWorkSheet.Range("M1").Value = "Night paycode"
                    shWorkSheet.Range("N1").Value = "Night surcharge"
                    shWorkSheet.Range("O1").Value = "Night hours"
                    shWorkSheet.Range("P1").Value = "Night payment"
                    shWorkSheet.Range("Q1").Value = "105 hours"
                    shWorkSheet.Range("R1").Value = "All shifts approved"
                    shWorkSheet.Range("S1").Value = "420 payment"
                    shWorkSheet.Range("T1").Value = "420 Sunday hours"
                    shWorkSheet.Range("U1").Value = "420 Night hours"
                    shWorkSheet.Range("V1").Value = "Paid leave sunday hours"
                    shWorkSheet.Range("W1").Value = "Sickness sunday hours"
                    shWorkSheet.Range("X1").Value = "Paid leave night hours"
                    shWorkSheet.Range("Y1").Value = "Sickness night hours"
                    shWorkSheet.Range("Z1").Value = "Last 3 months Sunday hours"
                    shWorkSheet.Range("AA1").Value = "Last 3 months Night hours"
                    shWorkSheet.Range("AB1").Value = "Paid leave days"

                    If selected_month = 3 Or
                       selected_month = 6 Or
                       selected_month = 9 Or
                       selected_month = 12 Then

                        Select Case selected_month
                            Case 3
                                quartal = "I"
                            Case 6
                                quartal = "II"
                            Case 9
                                quartal = "III"
                            Case 12
                                quartal = "IV"
                        End Select

                        shWorkSheet.Range("AC1").Value = quartal & " quartal working days"
                        shWorkSheet.Range("AD1").Value = quartal & " quartal working hours"
                        shWorkSheet.Range("AE1").Value = "paid leave in hours"
                        table.DataRange.Column(29).Style.NumberFormat.Format = "General;;"
                        table.DataRange.Column(30).Style.NumberFormat.Format = "General;;"
                        table.DataRange.Column(31).Style.NumberFormat.Format = "General;;"
                        table.DataRange.Column(29).Style.Border.LeftBorder = XLBorderStyleValues.Thin
                        table.DataRange.Column(29).Style.Border.RightBorder = XLBorderStyleValues.Thin
                        table.DataRange.Column(30).Style.Border.LeftBorder = XLBorderStyleValues.Thin
                        table.DataRange.Column(30).Style.Border.RightBorder = XLBorderStyleValues.Thin
                        table.DataRange.Column(31).Style.Border.LeftBorder = XLBorderStyleValues.Thin
                        table.DataRange.Column(31).Style.Border.RightBorder = XLBorderStyleValues.Thin
                    End If

                    table.DataRange.Column(7).Style.NumberFormat.Format = "hh:mm"
                    table.DataRange.Column(8).Style.NumberFormat.Format = "hh:mm"
                    table.DataRange.Column(10).Style.NumberFormat.Format = "#,##0.00 [$€-de-DE]"
                    table.DataRange.Column(11).Style.NumberFormat.Format = "General;;"
                    table.DataRange.Column(12).Style.NumberFormat.Format = "#,##0.00 [$€-de-DE]"
                    table.DataRange.Column(14).Style.NumberFormat.Format = "#,##0.00 [$€-de-DE]"
                    table.DataRange.Column(15).Style.NumberFormat.Format = "General;;"
                    table.DataRange.Column(16).Style.NumberFormat.Format = "#,##0.00 [$€-de-DE]"
                    table.DataRange.Column(17).Style.NumberFormat.Format = "General;;"
                    table.DataRange.Column(19).Style.NumberFormat.Format = "#,##0.00 [$€-de-DE]"
                    table.DataRange.Column(20).Style.NumberFormat.Format = "0.00;;"
                    table.DataRange.Column(21).Style.NumberFormat.Format = "0.00;;"
                    table.DataRange.Column(22).Style.NumberFormat.Format = "0.00;;"
                    table.DataRange.Column(23).Style.NumberFormat.Format = "General;;"
                    table.DataRange.Column(24).Style.NumberFormat.Format = "0.00;;"
                    table.DataRange.Column(25).Style.NumberFormat.Format = "General;;"
                    table.DataRange.Column(26).Style.NumberFormat.Format = "General;;"
                    table.DataRange.Column(27).Style.NumberFormat.Format = "General;;"
                    table.DataRange.Column(28).Style.NumberFormat.Format = "General;;"
                    OvertimeForm.format_yellow(table.DataRange.Column(12).AsRange)
                    OvertimeForm.format_yellow(table.DataRange.Column(16).AsRange)
                    OvertimeForm.format_yellow(table.DataRange.Column(19).AsRange)
                    table.DataRange.Column(12).Style.Font.Bold = True
                    table.DataRange.Column(16).Style.Font.Bold = True
                    table.DataRange.Column(19).Style.Font.Bold = True
                    OvertimeForm.format_blue(table.DataRange.Column(11).AsRange)
                    OvertimeForm.format_blue(table.DataRange.Column(15).AsRange)
                    OvertimeForm.format_blue(table.DataRange.Column(20).AsRange)
                    OvertimeForm.format_blue(table.DataRange.Column(21).AsRange)
                    table.DataRange.Column(18).Style.Border.LeftBorder = XLBorderStyleValues.Thin
                    table.DataRange.Column(18).Style.Border.RightBorder = XLBorderStyleValues.Thin
                    table.Theme = XLTableTheme.TableStyleLight8

                    OvertimeForm.format_white_borders(table.HeadersRow().Cells())

                    shWorkSheet.ShowGridLines = False
                    shWorkSheet.SheetView.Freeze(1, 4)

                    report_workbooks(hotel) = bkWorkBook
                End If
                '''''''

                If current_row.ContainsKey(hotel) Then row = current_row(hotel)

                For Each item_found In found

                    objhttp.open("GET", "https://openapi.planday.com/hr/v1.0/employees/" & item_found & "?special=BirthDate", False)
                    objhttp.setRequestHeader("Authorization", "Bearer " & api_token)
                    objhttp.setRequestHeader("X-ClientId", "ddca428b-8530-405d-9960-047132c49531")
                    objhttp.send()
                    Parsed = New Web.Script.Serialization.JavaScriptSerializer().Deserialize(Of Object)(objhttp.responseText)

                    sunday_FT_hours = 0
                    night_hours = 0
                    working_hours = 0
                    working_hours_krank = 0
                    sunday_FT_hours_krank = 0
                    night_hours_krank = 0
                    last_3_months_sunday_FT_hours = 0
                    last_3_months_night_hours = 0
                    quarter_working_days = 0
                    quarter_working_hours = 0
                    Dim azubi As Boolean = False

                    'Do Action
                    CurrentAction = CurrentAction + 1
                    ActionName = "Processing " & Trim(Parsed("data")("firstName")) & " " & Trim(Parsed("data")("lastName")) & "... "
                    'Report current status in %
                    bgw_export_deactivated.ReportProgress(100 * CurrentAction / BarTotalCount)

                    Dim employee_type As String = employee_types(Parsed("data")("employeeTypeId"))

                    objhttp.open("GET", "https://openapi.planday.com/contractrules/v1.0/employees/" & item_found, False)
                    objhttp.setRequestHeader("Authorization", "Bearer " & api_token)
                    objhttp.setRequestHeader("X-ClientId", "ddca428b-8530-405d-9960-047132c49531")
                    objhttp.send()
                    Parsed_temp = New Web.Script.Serialization.JavaScriptSerializer().Deserialize(Of Object)(objhttp.responseText)

                    Dim id_contract_rule As Integer
                    If objhttp.responseText = "{}" Then
                        id_contract_rule = 3527
                    Else
                        id_contract_rule = Parsed_temp("data")("id")
                    End If

                    objhttp.open("GET", "https://openapi.planday.com/hr/v1.0/employees/" & item_found, False)
                    objhttp.setRequestHeader("Authorization", "Bearer " & api_token)
                    objhttp.setRequestHeader("X-ClientId", "ddca428b-8530-405d-9960-047132c49531")
                    objhttp.send()
                    Parsed_temp = New Web.Script.Serialization.JavaScriptSerializer().Deserialize(Of Object)(objhttp.responseText)

                    Dim senior As Boolean
                    If Parsed_temp("data").ContainsKey("custom_221267") Then
                        If Parsed_temp("data")("custom_221267")("value") = "20:00" Then
                            senior = True
                        Else
                            senior = False
                        End If
                    Else
                        senior = False
                    End If

                    For Each temp_values In Parsed("data")("employeeGroups")
                        If employee_groups.ContainsKey(temp_values) Then
                            If employee_groups(temp_values) = "Azubis" Then
                                azubi = True
                                Exit For
                            Else
                                azubi = False
                            End If
                        End If
                    Next

                    current_employee = Trim(Parsed("data")("firstName")) & " " & Trim(Parsed("data")("lastName"))
                    If Len(current_employee) > 31 Then
                        current_employee = Strings.Left(current_employee, 31)
                    End If
                    current_workbook = hotel
                    OvertimeForm.calculate_last_3_months(item_found, Item, senior, If(Parsed("data").ContainsKey("birthDate"), Parsed("data")("birthDate"), ""))
                    OvertimeForm.calculate_All(item_found, Item, senior, If(Parsed("data").ContainsKey("birthDate"), Parsed("data")("birthDate"), ""))


                    sunday_FT_hours = Math.Round(sunday_FT_hours, 2)
                    night_hours = Math.Round(night_hours, 2)
                    working_hours = Math.Round(working_hours, 2)
                    working_hours_krank = Math.Round(working_hours_krank, 2)
                    sunday_FT_hours_krank = Math.Round(sunday_FT_hours_krank, 2)
                    night_hours_krank = Math.Round(night_hours_krank, 2)
                    last_3_months_sunday_FT_hours = Math.Round(last_3_months_sunday_FT_hours, 2)
                    last_3_months_night_hours = Math.Round(last_3_months_night_hours, 2)

                    If azubi Then
                        OvertimeForm.format_grey(report_workbooks(hotel).Worksheet("Exported data").Range("A" & row & ":J" & row))
                        OvertimeForm.format_grey(report_workbooks(hotel).Worksheet("Exported data").Range("M" & row & ":N" & row))
                    End If
                    report_workbooks(hotel).Worksheet("Exported data").Cell(row, 1).Value = department_number
                    report_workbooks(hotel).Worksheet("Exported data").Cell(row, 2).Value = Parsed("data")("salaryIdentifier")
                    report_workbooks(hotel).Worksheet("Exported data").Cell(row, 3).Value = Trim(Parsed("data")("firstName"))
                    report_workbooks(hotel).Worksheet("Exported data").Cell(row, 4).Value = Trim(Parsed("data")("lastName"))
                    report_workbooks(hotel).Worksheet("Exported data").Cell(row, 5).Value = employee_type
                    report_workbooks(hotel).Worksheet("Exported data").Cell(row, 6).Value = contract_rules(id_contract_rule)
                    If senior Then
                        report_workbooks(hotel).Worksheet("Exported data").Cell(row, 7).Value = 20 / 24
                    Else
                        report_workbooks(hotel).Worksheet("Exported data").Cell(row, 7).Value = 22 / 24
                    End If
                    report_workbooks(hotel).Worksheet("Exported data").Cell(row, 8).Value = 6 / 24
                    If azubi Then
                        report_workbooks(hotel).Worksheet("Exported data").Cell(row, 9).Value = 413
                    Else
                        report_workbooks(hotel).Worksheet("Exported data").Cell(row, 9).Value = 411
                    End If
                    If azubi Then
                        report_workbooks(hotel).Worksheet("Exported data").Cell(row, 10).Value = 2.5
                    Else
                        report_workbooks(hotel).Worksheet("Exported data").Cell(row, 10).Value = 4
                    End If
                    report_workbooks(hotel).Worksheet("Exported data").Cell(row, 11).Value = sunday_FT_hours
                    If azubi Then
                        report_workbooks(hotel).Worksheet("Exported data").Cell(row, 13).Value = 412
                    Else
                        report_workbooks(hotel).Worksheet("Exported data").Cell(row, 13).Value = 410
                    End If
                    If azubi Then
                        report_workbooks(hotel).Worksheet("Exported data").Cell(row, 14).Value = 1.25
                    Else
                        report_workbooks(hotel).Worksheet("Exported data").Cell(row, 14).Value = 2
                    End If
                    report_workbooks(hotel).Worksheet("Exported data").Cell(row, 15).Value = night_hours
                    If id_contract_rule = 3527 Or id_contract_rule = 4856 Then
                        report_workbooks(hotel).Worksheet("Exported data").Cell(row, 17).Value = working_hours + working_hours_krank
                        OvertimeForm.format_blue(report_workbooks(hotel).Worksheet("Exported data").Cell(row, 17).AsRange)
                    End If
                    If OvertimeForm.all_shifts_approved(item_found, Item) Then
                        report_workbooks(hotel).Worksheet("Exported data").Cell(row, 18).Value = "YES"
                        report_workbooks(hotel).Worksheet("Exported data").Cell(row, 18).Style.Fill.BackgroundColor = XLColor.Green
                    Else
                        report_workbooks(hotel).Worksheet("Exported data").Cell(row, 18).Value = "NO"
                        report_workbooks(hotel).Worksheet("Exported data").Cell(row, 18).Style.Fill.BackgroundColor = XLColor.Red
                    End If
                    report_workbooks(hotel).Worksheet("Exported data").Cell(row, 23).Value = sunday_FT_hours_krank
                    report_workbooks(hotel).Worksheet("Exported data").Cell(row, 25).Value = night_hours_krank
                    report_workbooks(hotel).Worksheet("Exported data").Cell(row, 26).Value = last_3_months_sunday_FT_hours
                    report_workbooks(hotel).Worksheet("Exported data").Cell(row, 27).Value = last_3_months_night_hours
                    report_workbooks(hotel).Worksheet("Exported data").Cell(row, 28).Value = OvertimeForm.paid_leave_days(item_found)
                    If selected_month = 3 Or
                       selected_month = 6 Or
                       selected_month = 9 Or
                       selected_month = 12 Then

                        OvertimeForm.calculate_quarter(item_found, Item, senior, If(Parsed("data").ContainsKey("birthDate"), Parsed("data")("birthDate"), ""))
                        quarter_working_hours = Math.Round(quarter_working_hours, 2)

                        If id_contract_rule = 3527 Or id_contract_rule = 4856 Then
                            report_workbooks(hotel).Worksheet("Exported data").Cell(row, 29).Value = quarter_working_days
                            report_workbooks(hotel).Worksheet("Exported data").Cell(row, 30).Value = quarter_working_hours
                            OvertimeForm.format_blue_noBorders(report_workbooks(hotel).Worksheet("Exported data").Cell(row, 29).AsRange)
                            OvertimeForm.format_blue_noBorders(report_workbooks(hotel).Worksheet("Exported data").Cell(row, 30).AsRange)
                            OvertimeForm.format_blue_noBorders(report_workbooks(hotel).Worksheet("Exported data").Cell(row, 31).AsRange)
                        End If
                    End If

                    row = row + 1
                    current_row(hotel) = row
                Next
            End If

        Next

        For Each workbook_name In workbook_names

            report_workbooks(workbook_name).Worksheet("Exported data").Table("export").HeadersRow.Style.Alignment.WrapText = True
            For i = 3 To 5
                Dim max_length As Integer = 0
                For Each cell In report_workbooks(workbook_name).Worksheet("Exported data").Table("export").DataRange.Column(i).Cells
                    If Len(cell.GetFormattedString) > max_length Then max_length = Len(cell.GetFormattedString)
                Next
                report_workbooks(workbook_name).Worksheet("Exported data").Columns(i).Width = OvertimeForm.CalculateColumnWidth(max_length)
            Next
            report_workbooks(workbook_name).Worksheet("Exported data").Columns("A:A").Width = 7
            report_workbooks(workbook_name).Worksheet("Exported data").Columns("B:B").Width = 9
            report_workbooks(workbook_name).Worksheet("Exported data").Columns("F:H").Width = 7.7
            report_workbooks(workbook_name).Worksheet("Exported data").Columns("I:J").Width = 11
            report_workbooks(workbook_name).Worksheet("Exported data").Columns("K:L").Width = 10
            report_workbooks(workbook_name).Worksheet("Exported data").Columns("M:N").Width = 11
            report_workbooks(workbook_name).Worksheet("Exported data").Columns("O:P").Width = 10
            report_workbooks(workbook_name).Worksheet("Exported data").Columns("Q:Q").Width = 7.7
            report_workbooks(workbook_name).Worksheet("Exported data").Columns("R:U").Width = 11
            report_workbooks(workbook_name).Worksheet("Exported data").Columns("V:AA").Width = 16
            report_workbooks(workbook_name).Worksheet("Exported data").Columns("AB:AB").Width = 11
            report_workbooks(workbook_name).Worksheet("Exported data").Columns("AC:AE").Width = 16.5

            report_workbooks(workbook_name).Worksheet("Exported data").Table("export").SetAutoFilter.Column(1).NotEqualTo(Of String)("")
            report_workbooks(workbook_name).Worksheet("Exported data").Table("export").AutoFilter.Sort(3, XLSortOrder.Ascending)
            Dim formula_row As Integer

            formula_row = 2
            For Each cell In report_workbooks(workbook_name).Worksheet("Exported data").Table("export").DataRange.Column(12).Cells
                cell.FormulaA1 = "=IF(OR(C" & formula_row & "="""",K" & formula_row & "=0),"""",J" & formula_row & "*K" & formula_row & ")"
                formula_row = formula_row + 1
            Next

            formula_row = 2
            For Each cell In report_workbooks(workbook_name).Worksheet("Exported data").Table("export").DataRange.Column(16).Cells
                cell.FormulaA1 = "=IF(OR(C" & formula_row & "="""",O" & formula_row & "=0),"""",N" & formula_row & "*O" & formula_row & ")"
                formula_row = formula_row + 1
            Next

            formula_row = 2
            For Each cell In report_workbooks(workbook_name).Worksheet("Exported data").Table("export").DataRange.Column(19).Cells
                cell.FormulaA1 = "=IF(OR(C" & formula_row & "="""",AND(T" & formula_row & "=0,U" & formula_row & "=0)),"""",T" & formula_row & "*J" & formula_row & "+U" & formula_row & "*N" & formula_row & ")"
                formula_row = formula_row + 1
            Next

            formula_row = 2
            For Each cell In report_workbooks(workbook_name).Worksheet("Exported data").Table("export").DataRange.Column(22).Cells
                cell.FormulaA1 = "=IF(C" & formula_row & "="""","""",Z" & formula_row & "/65*(AB" & formula_row & "))"
                formula_row = formula_row + 1
            Next

            formula_row = 2
            For Each cell In report_workbooks(workbook_name).Worksheet("Exported data").Table("export").DataRange.Column(24).Cells
                cell.FormulaA1 = "=IF(C" & formula_row & "="""","""",AA" & formula_row & "/65*(AB" & formula_row & "))"
                formula_row = formula_row + 1
            Next

            formula_row = 2
            For Each cell In report_workbooks(workbook_name).Worksheet("Exported data").Table("export").DataRange.Column(20).Cells
                cell.FormulaA1 = "=IF(C" & formula_row & "="""","""",V" & formula_row & "+W" & formula_row & ")"
                formula_row = formula_row + 1
            Next

            formula_row = 2
            For Each cell In report_workbooks(workbook_name).Worksheet("Exported data").Table("export").DataRange.Column(21).Cells
                cell.FormulaA1 = "=IF(C" & formula_row & "="""","""",X" & formula_row & "+Y" & formula_row & ")"
                formula_row = formula_row + 1
            Next

            If report_workbooks(workbook_name).Worksheet("Exported data").Table("export").Columns.Count = 31 Then
                formula_row = 2
                For Each cell In report_workbooks(workbook_name).Worksheet("Exported data").Table("export").DataRange.Column(31).Cells
                    cell.FormulaA1 = "=IFERROR(ROUND((AD" & formula_row & "/AC" & formula_row & ")*(6.25*((AC" & formula_row & "/13)/5)),2),0)"
                    formula_row = formula_row + 1
                Next
            End If

            report_workbooks(workbook_name).SaveAs(export_folder & "\" & workbook_name)
        Next
    End Sub

    Private Sub bgw_export_deactivated_ProgressChanged(sender As Object, e As ProgressChangedEventArgs) Handles bgw_export_deactivated.ProgressChanged
        ProgressBarForm.NextAction(CurrentAction, ActionName)
    End Sub

    Private Sub bgw_export_deactivated_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles bgw_export_deactivated.RunWorkerCompleted
        CurrentAction = 0
        BarTotalCount = 0
        ProgressBarForm.Complete()
        Erase selected_departments
    End Sub

    Private Sub CLOSE_button_Click(sender As Object, e As EventArgs) Handles CLOSE_button.Click
        Me.Close()
    End Sub

    Private Sub select_hotel_ColumnWidthChanging(sender As Object, e As ColumnWidthChangingEventArgs) Handles select_hotel.ColumnWidthChanging
        e.Cancel = True
        e.NewWidth = select_hotel.Columns(e.ColumnIndex).Width
    End Sub

    Private Sub select_hotel_KeyDown(sender As Object, e As KeyEventArgs) Handles select_hotel.KeyDown
        If e.KeyCode = Keys.Enter Then
            START_button_Click(sender, e)
        End If
        If e.KeyData = (Keys.A Or Keys.Control) Then
            NativeMethods.SelectAllItems(select_hotel)
        End If
    End Sub

    Private Sub Label_folder_Click(sender As Object, e As EventArgs) Handles Label_folder.Click
        ExportPDF.SelectFolder(Label_folder, ToolTip1)
    End Sub

    Private Sub YEAR_button_Click(sender As Object, e As EventArgs) Handles YEAR_button.Click
        PayrollPeriodForm.select_year.Items.Clear()
        With PayrollPeriodForm.select_year
            For Each key In payroll_periods.Keys
                .Items.Add(key)
            Next
        End With
        PayrollPeriodForm.select_year.Text = YEAR_button.Text
        PayrollPeriodForm.ShowDialog(Me)
    End Sub

    Private Sub ExportForm_Activated(sender As Object, e As EventArgs) Handles Me.Activated
        If ImChangingStuff Then Exit Sub
        If YEAR_button.Text <> 2019 Then
            select_month.Items.Clear()
            With select_month
                For Each monat In System.Globalization.CultureInfo.CurrentCulture.DateTimeFormat.MonthNames()
                    If monat <> "" Then .Items.Add(monat)
                Next
            End With
            select_month.Text = Now.ToString("MMMM", System.Globalization.CultureInfo.CurrentCulture)
        End If
    End Sub

    Private Sub select_month_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles select_month.SelectionChangeCommitted
        OvertimeForm.set_params()
    End Sub

End Class