Imports System.ComponentModel
Imports System.Net
Imports ClosedXML.Excel
Imports System.Configuration

Public Class OvertimeForm

    Public Shared current_width As Integer
    Public Shared ImChangingStuff As Boolean
    Public Shared width_changed As Boolean

    Public Shared ActionName As String
    Public Shared CurrentAction As Long
    Public Shared BarTotalCount As Long
    Public Shared export_folder As String
    Public Shared selected_month As Integer
    Public Shared myWorkingForm As Form
    Public Shared normalDPI As Boolean
    Public Shared ownDepartment As String

    Public Shared api_token As String
    Public Shared api_client_id As String
    Public Shared api_refresh_token As String
    Public Shared PLANDAY_portal As String
    Public Shared name_extension As String 'MM_yyyy for the payroll period
    Public Shared vacation_types() ' array with vacationTypeIDs
    Public Shared overtime_types() ' array with overtimeTypeIDs
    Public Shared deactivated()    'array with employeeIDs from all deactivated employees
    Public Shared selected_departments() As Integer
    Public Shared employees_records() As EmployeeRecord
    Public Shared payroll_periods As New Dictionary(Of Integer, PayrollPeriods)
    Public Shared departments_FT As New Dictionary(Of String, Date()) 'dictionary with keys: department_id ,values: array of FT
    Public Shared shift_types As New Dictionary(Of String, String) 'dictionary with keys: shiftTypeIds ,values: names
    Public Shared employee_groups As New Dictionary(Of String, String) 'dictionary with keys: employeeGroupIDs ,values: names
    Public Shared employee_types As New Dictionary(Of String, String) 'dictionary with keys: shiftTypeIds ,values: names
    Public Shared departments As New Dictionary(Of String, String) 'dictionary with keys: department_id ,values: names
    Public Shared department_numbers As New Dictionary(Of String, String) 'dictionary with keys: department_id ,values: hotel number
    Public Shared contract_rules As New Dictionary(Of String, String) 'dictionary with keys: contract_rule_id ,values: hours
    Public Shared employees As New Dictionary(Of String, String) 'dictionary with keys: names , values: employee_id
    Public Shared column_widths As New Dictionary(Of Integer, Integer) 'dictionary with keys: column_index , values: column_width
    Public Shared FT_Name As New Dictionary(Of Date, String) 'dictionary with keys: public holiday , values: name
    Public Shared employee_kw_hours As New Dictionary(Of String, ShiftRecord()) 'dictionary with keys: employee_name ,values: array of shift records for the week
    Public Shared report_workbooks As New Dictionary(Of String, XLWorkbook)

    Public Shared current_employee As String
    Public Shared current_workbook As String
    Public Shared start_monday As Date
    Public Shared hotel_id As String
    Public Shared Boolean_write As Boolean
    Public Shared sunday_FT_hours_krank As Double
    Public Shared night_hours_krank As Double
    Public Shared working_hours_krank As Double
    Public Shared sunday_FT_hours As Double
    Public Shared XMAS_150 As Double
    Public Shared XMAS_125 As Double
    Public Shared working_hours As Double
    Public Shared night_hours As Double
    Public Shared start_ As String
    Public Shared end_ As String
    Public Shared last_3_months_start_ As String
    Public Shared last_3_months_end_ As String
    Public Shared last_3_months_sunday_FT_hours As Double
    Public Shared last_3_months_night_hours As Double
    Public Shared quarter_start_ As String
    Public Shared quarter_end_ As String
    Public Shared quarter_working_hours As Double
    Public Shared quarter_working_days As Double
    Public Shared current_year As Integer

    Sub test_1()
        Dim index As Integer
        Setup.ShowDialog()

        If PLANDAY_portal = "DE" Then
            For Each item In My.Settings.DE_shifts
                index = My.Settings.DE_shifts.IndexOf(item)
                Debug.Print(index & " " & item & " " & My.Settings.DE_shifts_names.Item(index) & " " & My.Settings.DE_shifts_surcharges.Item(index) & " " & My.Settings.DE_shifts_overtime.Item(index) & " " & My.Settings.DE_shifts_sick_paid.Item(index) & " " & My.Settings.DE_shifts_sick_surcharges.Item(index))
            Next
            For Each item In My.Settings.DE_departments
                index = My.Settings.DE_departments.IndexOf(item)
                Debug.Print(index & " " & item & " " & My.Settings.DE_departments_names.Item(index) & " " & My.Settings.DE_departments_bundesland.Item(index))
            Next
        End If
        If PLANDAY_portal = "HQ" Then
            For Each item In My.Settings.HQ_shifts
                index = My.Settings.HQ_shifts.IndexOf(item)
                Debug.Print(index & " " & item & " " & My.Settings.HQ_shifts_names.Item(index) & " " & My.Settings.HQ_shifts_surcharges.Item(index) & " " & My.Settings.HQ_shifts_overtime.Item(index) & " " & My.Settings.HQ_shifts_sick_paid.Item(index) & " " & My.Settings.HQ_shifts_sick_surcharges.Item(index))
            Next
            For Each item In My.Settings.HQ_departments
                index = My.Settings.HQ_departments.IndexOf(item)
                Debug.Print(index & " " & item & " " & My.Settings.HQ_departments_names.Item(index) & " " & My.Settings.HQ_departments_bundesland.Item(index))
            Next
        End If

    End Sub

    Public Function GetItem(ByVal i As Integer) As ListViewItem
        If select_employees.InvokeRequired Then
            Return CType(select_employees.Invoke(New Func(Of ListViewItem)(Function() GetItem(i))), ListViewItem)
        Else
            Return select_employees.Items.Item(i)
        End If
    End Function

    Public Function GetItemChecked(ByVal i As Integer) As Boolean
        If select_employees.InvokeRequired Then
            Return CType(select_employees.Invoke(New Func(Of Boolean)(Function() GetItemChecked(i))), Boolean)
        Else
            Return select_employees.Items.Item(i).Checked
        End If
    End Function

    Public Sub SetItemChecked(ByVal i As Integer, ByVal checked As Boolean)
        If select_employees.InvokeRequired Then
            select_employees.Invoke(Sub() select_employees.Items.Item(i).Checked = checked)
        Else
            select_employees.Items.Item(i).Checked = checked
        End If
    End Sub

    Public Sub SetSubitemText(ByVal item_index As Integer, ByVal subitem_index As Integer, ByVal text As String)
        If select_employees.InvokeRequired Then
            select_employees.Invoke(Sub() select_employees.Items.Item(item_index).SubItems(subitem_index).Text = text)
        Else
            select_employees.Items.Item(item_index).SubItems(subitem_index).Text = text
        End If
    End Sub

    Private Sub CLOSE_button_Click(sender As Object, e As EventArgs) Handles CLOSE_button.Click
        Me.Close()
    End Sub

    Private Sub MODUS_button_Click(sender As Object, e As EventArgs) Handles MODUS_button.Click
        If Me.MODUS_button.Checked Then
            Me.MODUS_button.Text = "Senden AN"
            Me.WRITE_button.Visible = True
            Me.DELETE_button.Visible = True
            Me.PDF_button.Left = Me.DELETE_button.Left - Me.DELETE_button.Width
            Me.EXCEL_button.Left = Me.WRITE_button.Left - Me.WRITE_button.Width
        Else
            Me.MODUS_button.Text = "Senden AUS"
            Me.WRITE_button.Visible = False
            Me.DELETE_button.Visible = False
            Me.PDF_button.Left = Me.DELETE_button.Left
            Me.EXCEL_button.Left = Me.WRITE_button.Left
        End If
    End Sub

    Private Sub select_hotel_KeyDown(sender As Object, e As KeyEventArgs) Handles select_hotel.KeyDown
        If e.KeyCode = Keys.Enter Then
            LOAD_button_Click(sender, e)
        End If
    End Sub

    Private Sub select_kw_KeyDown(sender As Object, e As KeyEventArgs) Handles select_kw.KeyDown
        If e.KeyCode = Keys.Enter Then
            LOAD_button_Click(sender, e)
        End If
    End Sub

    Private Sub shiftsApproved_Click(sender As Object, e As EventArgs) Handles shiftsApproved.Click
        If Me.select_employees.Items.Count > 0 Then MsgBox("Ein Neuladen der Daten ist erforderlich!", , "Hinweis")
    End Sub

    Private Sub active_employees_Click(sender As Object, e As EventArgs) Handles active_employees.Click
        If Me.select_employees.Items.Count > 0 Then MsgBox("Ein Neuladen der Daten ist erforderlich!", , "Hinweis")
        If Me.active_employees.Checked Then
            Me.active_employees.Text = "Nur in PLANDAY aktive Mitarbeiter"
            Me.active_employees.ForeColor = Color.Black
        Else
            Me.active_employees.Text = "ALLE Mitarbieter => langsam"
            Me.active_employees.ForeColor = Color.Red
        End If
    End Sub

    Private Sub WRITE_button_Click(sender As Object, e As EventArgs) Handles WRITE_button.Click
        If Me.select_employees.Items.Count = 0 Then Exit Sub
        If Me.select_employees.CheckedItems.Count = 0 Then
            MsgBox("Kein Mitarbeiter ausgewählt!")
            Exit Sub
        End If
        Write_overtime()
    End Sub

    Private Sub Write_overtime()
        Boolean_write = True
        'Do work in
        If normalDPI Then
            WorkingForm.Show(Me)
        Else
            LoadingForm.Show(Me)
        End If
        bgw_write_overtime.RunWorkerAsync()
    End Sub

    Private Sub DELETE_button_Click(sender As Object, e As EventArgs) Handles DELETE_button.Click
        If Me.select_employees.Items.Count = 0 Then Exit Sub
        If Me.select_employees.CheckedItems.Count = 0 Then
            MsgBox("Kein Mitarbeiter ausgewählt!")
            Exit Sub
        End If
        Delete_overtime()
    End Sub

    Private Sub Delete_overtime()
        Boolean_write = False
        If normalDPI Then
            WorkingForm.Show(Me)
        Else
            LoadingForm.Show(Me)
        End If
        bgw_write_overtime.RunWorkerAsync()
    End Sub

    Private Sub select_employees_ItemSelectionChanged(sender As Object, e As ListViewItemSelectionChangedEventArgs) Handles select_employees.ItemSelectionChanged
        e.Item.Selected = Nothing
    End Sub

    Private Sub select_employees_ColumnWidthChanged(sender As Object, e As ColumnWidthChangedEventArgs) Handles select_employees.ColumnWidthChanged
        WidthChanging = False
        If ImChangingStuff Then Exit Sub

        ImChangingStuff = True

        If select_employees.Columns(e.ColumnIndex).Width < column_widths(e.ColumnIndex) Then
            select_employees.Columns(e.ColumnIndex).Width = -2
        End If

        column_widths(e.ColumnIndex) = select_employees.Columns(e.ColumnIndex).Width
        ImChangingStuff = False

    End Sub

    Private Sub select_employees_MouseMove(sender As Object, e As MouseEventArgs) Handles select_employees.MouseMove
        Static prevMousePos As Point = New Point(-1, -1)

        Dim lv As ListView = TryCast(sender, ListView)
        If lv Is Nothing Then _
            Exit Sub
        If prevMousePos = MousePosition Then _
            Exit Sub  ' to avoid annoying flickering

        With lv.HitTest(lv.PointToClient(MousePosition))
            If .SubItem IsNot Nothing AndAlso Not String.IsNullOrEmpty(.SubItem.Text) AndAlso .Item.SubItems.IndexOf(.SubItem) = 2 Then
                ToolTip1.Active = True
                ToolTip1.UseFading = True
                Cursor = Cursors.Hand
                ToolTip1.Show("Klick hier für detailierte KW-Stunden Ansicht für " & .Item.Text, .Item.ListView,
                .Item.ListView.PointToClient(Cursor.Position + New Size(1, 20)), 2000)
            Else
                Cursor = Cursors.Default
                ToolTip1.Active = False
            End If
            prevMousePos = MousePosition
        End With
    End Sub

    Private Sub select_employees_MouseDown(sender As Object, e As MouseEventArgs) Handles select_employees.MouseDown
        Dim info As ListViewHitTestInfo = select_employees.HitTest(e.X, e.Y)
        If info.SubItem IsNot Nothing AndAlso Not String.IsNullOrEmpty(info.SubItem.Text) AndAlso info.Item.SubItems.IndexOf(info.SubItem) = 2 Then
            If normalDPI And Not width_changed Then
                OvertimeRecords.Width = OvertimeRecords.Width + 80
                width_changed = True
            End If
            load_employee_records(info.Item.Text)
            OvertimeRecords.Text = loaded.Text & " Stunden"
            OvertimeRecords.ShowDialog(Me)
        End If
    End Sub

    Public Sub init_payroll_periods()
        Dim payroll_2019 As PayrollPeriods = New PayrollPeriods(2019,
                                                                New MonthPeriod(CDate("05.12.2018"), CDate("04.01.2019")),
                                                                New MonthPeriod(CDate("05.01.2019"), CDate("04.02.2019")),
                                                                New MonthPeriod(CDate("05.02.2019"), CDate("04.03.2019")),
                                                                New MonthPeriod(CDate("05.03.2019"), CDate("05.04.2019")),
                                                                New MonthPeriod(CDate("06.04.2019"), CDate("05.05.2019")),
                                                                New MonthPeriod(CDate("06.05.2019"), CDate("04.06.2019")),
                                                                New MonthPeriod(CDate("05.06.2019"), CDate("05.07.2019")),
                                                                New MonthPeriod(CDate("06.07.2019"), CDate("04.08.2019")),
                                                                New MonthPeriod(CDate("05.08.2019"), CDate("04.09.2019")),
                                                                New MonthPeriod(CDate("05.09.2019"), CDate("04.10.2019")),
                                                                New MonthPeriod(CDate("05.10.2019"), CDate("04.11.2019")),
                                                                New MonthPeriod(CDate("05.11.2019"), CDate("03.12.2019")))

        Dim payroll_2020 As PayrollPeriods = New PayrollPeriods(2020,
                                                                New MonthPeriod(CDate("04.12.2019"), CDate("03.01.2020")),
                                                                New MonthPeriod(CDate("04.01.2020"), CDate("03.02.2020")),
                                                                New MonthPeriod(CDate("04.02.2020"), CDate("03.03.2020")),
                                                                New MonthPeriod(CDate("04.03.2020"), CDate("03.04.2020")),
                                                                New MonthPeriod(CDate("04.04.2020"), CDate("03.05.2020")),
                                                                New MonthPeriod(CDate("04.05.2020"), CDate("03.06.2020")),
                                                                New MonthPeriod(CDate("04.06.2020"), CDate("03.07.2020")),
                                                                New MonthPeriod(CDate("04.07.2020"), CDate("03.08.2020")),
                                                                New MonthPeriod(CDate("04.08.2020"), CDate("03.09.2020")),
                                                                New MonthPeriod(CDate("04.09.2020"), CDate("03.10.2020")),
                                                                New MonthPeriod(CDate("04.10.2020"), CDate("03.11.2020")),
                                                                New MonthPeriod(CDate("04.11.2020"), CDate("03.12.2020")))

        Dim payroll_2021 As PayrollPeriods = New PayrollPeriods(2021,
                                                                New MonthPeriod(CDate("04.12.2020"), CDate("03.01.2021")),
                                                                New MonthPeriod(CDate("04.01.2021"), CDate("03.02.2021")),
                                                                New MonthPeriod(CDate("04.02.2021"), CDate("03.03.2021")),
                                                                New MonthPeriod(CDate("04.03.2021"), CDate("03.04.2021")),
                                                                New MonthPeriod(CDate("04.04.2021"), CDate("03.05.2021")),
                                                                New MonthPeriod(CDate("04.05.2021"), CDate("03.06.2021")),
                                                                New MonthPeriod(CDate("04.06.2021"), CDate("03.07.2021")),
                                                                New MonthPeriod(CDate("04.07.2021"), CDate("03.08.2021")),
                                                                New MonthPeriod(CDate("04.08.2021"), CDate("03.09.2021")),
                                                                New MonthPeriod(CDate("04.09.2021"), CDate("03.10.2021")),
                                                                New MonthPeriod(CDate("04.10.2021"), CDate("03.11.2021")),
                                                                New MonthPeriod(CDate("04.11.2021"), CDate("03.12.2021")))

        Dim payroll_2022 As PayrollPeriods = New PayrollPeriods(2022,
                                                                New MonthPeriod(CDate("04.12.2021"), CDate("03.01.2022")),
                                                                New MonthPeriod(CDate("04.01.2022"), CDate("03.02.2022")),
                                                                New MonthPeriod(CDate("04.02.2022"), CDate("03.03.2022")),
                                                                New MonthPeriod(CDate("04.03.2022"), CDate("03.04.2022")),
                                                                New MonthPeriod(CDate("04.04.2022"), CDate("03.05.2022")),
                                                                New MonthPeriod(CDate("04.05.2022"), CDate("03.06.2022")),
                                                                New MonthPeriod(CDate("04.06.2022"), CDate("03.07.2022")),
                                                                New MonthPeriod(CDate("04.07.2022"), CDate("03.08.2022")),
                                                                New MonthPeriod(CDate("04.08.2022"), CDate("03.09.2022")),
                                                                New MonthPeriod(CDate("04.09.2022"), CDate("03.10.2022")),
                                                                New MonthPeriod(CDate("04.10.2022"), CDate("03.11.2022")),
                                                                New MonthPeriod(CDate("04.11.2022"), CDate("03.12.2022")))

        payroll_periods(2019) = payroll_2019
        payroll_periods(2020) = payroll_2020
        payroll_periods(2021) = payroll_2021
        payroll_periods(2022) = payroll_2022
    End Sub

    Public Sub set_params()

        selected_month = DateTime.ParseExact(ExportForm.select_month.Text, "MMMM", System.Globalization.CultureInfo.CurrentCulture).Month

        start_ = Format(payroll_periods(ExportForm.YEAR_button.Text).Periods(selected_month - 1).StartDate, "yyyy-MM-dd")
        end_ = Format(payroll_periods(ExportForm.YEAR_button.Text).Periods(selected_month - 1).EndDate, "yyyy-MM-dd")
        name_extension = Format(DateSerial(CInt(ExportForm.YEAR_button.Text), selected_month, 1), "_MM_yyyy")

        Select Case selected_month

            Case 1
                last_3_months_start_ = Format(payroll_periods(CInt(ExportForm.YEAR_button.Text) - 1).Periods(9).StartDate, "yyyy-MM-dd")
                last_3_months_end_ = Format(payroll_periods(CInt(ExportForm.YEAR_button.Text) - 1).Periods(11).EndDate, "yyyy-MM-dd")
            Case 2
                last_3_months_start_ = Format(payroll_periods(CInt(ExportForm.YEAR_button.Text) - 1).Periods(10).StartDate, "yyyy-MM-dd")
                last_3_months_end_ = Format(payroll_periods(ExportForm.YEAR_button.Text).Periods(selected_month - 2).EndDate, "yyyy-MM-dd")
            Case 3
                last_3_months_start_ = Format(payroll_periods(CInt(ExportForm.YEAR_button.Text) - 1).Periods(11).StartDate, "yyyy-MM-dd")
                last_3_months_end_ = Format(payroll_periods(ExportForm.YEAR_button.Text).Periods(selected_month - 2).EndDate, "yyyy-MM-dd")
                quarter_start_ = Format(payroll_periods(ExportForm.YEAR_button.Text).Periods(selected_month - 3).StartDate, "yyyy-MM-dd")
                quarter_end_ = Format(payroll_periods(ExportForm.YEAR_button.Text).Periods(selected_month - 1).EndDate, "yyyy-MM-dd")
            Case Else
                If selected_month Mod 3 = 0 Then
                    quarter_start_ = Format(payroll_periods(ExportForm.YEAR_button.Text).Periods(selected_month - 3).StartDate, "yyyy-MM-dd")
                    quarter_end_ = Format(payroll_periods(ExportForm.YEAR_button.Text).Periods(selected_month - 1).EndDate, "yyyy-MM-dd")
                End If
                last_3_months_start_ = Format(payroll_periods(ExportForm.YEAR_button.Text).Periods(selected_month - 4).StartDate, "yyyy-MM-dd")
                last_3_months_end_ = Format(payroll_periods(ExportForm.YEAR_button.Text).Periods(selected_month - 2).EndDate, "yyyy-MM-dd")

        End Select
    End Sub

    Public Sub load_employee_records(employee_name As String)
        ImChangingStuff = True

        Dim col_count, i, y As Integer
        Dim summe As Double

        OvertimeRecords.select_employee.Items.Clear()
        OvertimeRecords.KW_Records.Clear()

        With OvertimeRecords.select_employee
            For Each Key In employee_kw_hours.Keys
                .Items.Add(Key)
            Next
        End With

        'Add the column headers
        OvertimeRecords.KW_Records.Columns.Add(text:="Datum", width:=90)
        OvertimeRecords.KW_Records.Columns.Add(text:="Schichtart", width:=190)
        OvertimeRecords.KW_Records.Columns.Add(text:="Start", width:=60, textAlign:=HorizontalAlignment.Right)
        OvertimeRecords.KW_Records.Columns.Add(text:="Ende", width:=60, textAlign:=HorizontalAlignment.Right)
        OvertimeRecords.KW_Records.Columns.Add(text:="Pause von", width:=90, textAlign:=HorizontalAlignment.Right)
        OvertimeRecords.KW_Records.Columns.Add(text:="Pause bis", width:=90, textAlign:=HorizontalAlignment.Right)
        OvertimeRecords.KW_Records.Columns.Add(text:="Pause von", width:=90, textAlign:=HorizontalAlignment.Right)
        OvertimeRecords.KW_Records.Columns.Add(text:="Pause bis", width:=90, textAlign:=HorizontalAlignment.Right)
        OvertimeRecords.KW_Records.Columns.Add(text:="IST", width:=60, textAlign:=HorizontalAlignment.Right)
        OvertimeRecords.KW_Records.Columns.Add(text:="Urlaubstage", width:=100, textAlign:=HorizontalAlignment.Right)

        col_count = OvertimeRecords.KW_Records.Columns.Count


        If employee_kw_hours.ContainsKey(employee_name) Then
            Dim ListItem As ListViewItem
            For i = 0 To employee_kw_hours(employee_name).Count - 1
                If i <> employee_kw_hours(employee_name).Count - 1 Then

                    ListItem = OvertimeRecords.KW_Records.Items.Add(text:=employee_kw_hours(employee_name)(i).Shift_date)
                    ListItem.UseItemStyleForSubItems = False
                    ListItem.SubItems.Add(text:=employee_kw_hours(employee_name)(i).Shift_type)
                    ListItem.SubItems.Add(text:=employee_kw_hours(employee_name)(i).Shift_from)
                    ListItem.SubItems.Add(text:=employee_kw_hours(employee_name)(i).Shift_to)
                    ListItem.SubItems.Add(text:=employee_kw_hours(employee_name)(i).Shift_break_1_from)
                    ListItem.SubItems.Add(text:=employee_kw_hours(employee_name)(i).Shift_break_1_to)
                    ListItem.SubItems.Add(text:=employee_kw_hours(employee_name)(i).Shift_break_2_from)
                    ListItem.SubItems.Add(text:=employee_kw_hours(employee_name)(i).Shift_break_2_to)
                    ListItem.SubItems.Add(text:=employee_kw_hours(employee_name)(i).Shift_length)
                    ListItem.SubItems(col_count - 2).ForeColor = Color.Blue
                Else
                    ''''Urlaubs Reihe'''''
                    ListItem = OvertimeRecords.KW_Records.Items.Add(text:="")
                    ListItem.UseItemStyleForSubItems = False

                    For y = 1 To col_count - 1
                        ListItem.SubItems.Add(text:="")
                    Next

                    ListItem.SubItems(col_count - 3).Text = "Urlaub:"
                    ListItem.SubItems(col_count - 2).Text = employee_kw_hours(employee_name)(i).Shift_length
                    ListItem.SubItems(col_count - 2).ForeColor = Color.Blue
                    ListItem.SubItems(col_count - 1).Text = employee_kw_hours(employee_name)(i).Holiday_days
                End If
            Next


            'IST Summe berechnen
            For Each Item In OvertimeRecords.KW_Records.Items
                summe = summe + If(Item.SubItems(col_count - 2).Text = "", 0, CDbl(Item.SubItems(col_count - 2).Text))
            Next

            'TOTAL Reihe'
            ListItem = OvertimeRecords.KW_Records.Items.Add(text:="")
            ListItem.UseItemStyleForSubItems = False

            For i = 1 To col_count - 1
                ListItem.SubItems.Add(text:="")
            Next
            ListItem.SubItems(col_count - 3).Text = "TOTAL:"
            ListItem.SubItems(col_count - 3).Font = New Font(ListItem.SubItems(col_count - 3).Font, FontStyle.Bold)
            ListItem.SubItems(col_count - 2).Text = summe
            ListItem.SubItems(col_count - 2).Font = New Font(ListItem.SubItems(col_count - 3).Font, FontStyle.Bold)
            ListItem.SubItems(col_count - 2).ForeColor = Color.Blue
        End If

        If normalDPI Then
            OvertimeRecords.KW_Records.BeginUpdate()
            For i = 0 To OvertimeRecords.KW_Records.Columns.Count - 1
                Dim colPercentage As Double = Convert.ToInt32(OvertimeRecords.KW_Records.Columns(i).Width) / 925
                OvertimeRecords.KW_Records.Columns(i).Width = CInt(colPercentage * OvertimeRecords.KW_Records.Width)
            Next
            OvertimeRecords.KW_Records.EndUpdate()
        End If

        OvertimeRecords.select_employee.Text = employee_name

        ImChangingStuff = False
    End Sub

    Private Sub OvertimeForm_LocationChanged(sender As Object, e As EventArgs) Handles Me.LocationChanged
        If normalDPI Then
            WorkingForm.Location = New Point(Location.X + Width / 2 - WorkingForm.Width / 2, Location.Y + Height / 2 - WorkingForm.Height / 2 + 60)
        Else
            LoadingForm.Location = New Point(Location.X + Width / 2 - LoadingForm.Width / 2, Location.Y + Height / 2 - LoadingForm.Height / 2 + 100)
        End If
    End Sub

    ' Berechnet das Datum des Ostersonntags
    ' des angegeben Jahres
    Public Function Ostern(Optional ByVal nYear As Long = 0) As Date
        Dim d As Long
        Dim Delta As Long

        ' Wenn kein Jahr angegeben, aktuelles Jahr heranziehen
        If nYear = 0 Then nYear = Year(Now)

        ' Delta
        d = (((255 - 11 * (nYear Mod 19)) - 21) Mod 30) + 21
        Delta = d + IIf(d > 48, 1, 0) + 6 -
    ((nYear + Int(nYear / 4) + d + IIf(d > 48, 1, 0) + 1) Mod 7)

        ' Osterdatum zurückgeben
        Ostern = DateAdd("d", Delta, DateSerial(nYear, 3, 1))
    End Function

    ''''''''''Initialisieren'''''''''''''''''''''''''
    Sub initialize()

        Dim objhttp As Object = CreateObject("MSXML2.XMLHTTP")
        Dim Parsed As Dictionary(Of String, Object)

        start_monday = Date.Today.AddDays(-Weekday(Date.Today, 2) - 6)

        'api_client_id = "ddca428b-8530-405d-9960-047132c49531" 'Ventzi
        'api_refresh_token = "r3xT0r-WAUiIlgQWDYWVsw" 'Ventzi

        If PLANDAY_portal = "DE" Then
            api_client_id = "3b0f80f6-6049-4396-9c1c-06bd2492abe1" 'Donovan DE
            api_refresh_token = "cup3ruhzYkaDRnpXkcwdpw" 'Donovan DE
        End If

        If PLANDAY_portal = "HQ" Then
            api_client_id = "5edbf7ee-5bc4-41f0-9e35-7fcd85d89b7f" 'Donovan HQ
            api_refresh_token = "I8IHzHsG20COusX60fXAYg" 'Donovan HQ
        End If



        If api_token = "" Then
            ''''''''''''''''''''get api_token''''''''''''''''''''''''
            objhttp.Open("POST", "https://openapi-login.planday.com/connect/token", False)
            objhttp.SetRequestHeader("Content-Type", "application/x-www-form-urlencoded")
            objhttp.Send("client_id=" & api_client_id & "&grant_type=refresh_token&refresh_token=" & api_refresh_token)
            Parsed = New Web.Script.Serialization.JavaScriptSerializer().Deserialize(Of Object)(objhttp.ResponseText)
            api_token = Parsed("access_token")

        End If

        '''''''''''EMPLYOEE TYPES''''''''''''''''
        objhttp.Open("GET", "https://openapi.planday.com/hr/v1.0/employeetypes", False)
        objhttp.SetRequestHeader("Authorization", "Bearer " & api_token)
        objhttp.SetRequestHeader("X-ClientId", api_client_id)
        objhttp.Send()
        Parsed = New Web.Script.Serialization.JavaScriptSerializer().Deserialize(Of Object)(objhttp.ResponseText)

        For Each key In Parsed("data")
            employee_types(key("id")) = key("name")
        Next

        ''''''''''''''''''EMPLOYEE GROUPS'''''''''''
        objhttp.Open("GET", "https://openapi.planday.com/hr/v1.0/employeegroups", False)
        objhttp.SetRequestHeader("Authorization", "Bearer " & api_token)
        objhttp.SetRequestHeader("X-ClientId", api_client_id)
        objhttp.Send()
        Parsed = New Web.Script.Serialization.JavaScriptSerializer().Deserialize(Of Object)(objhttp.ResponseText)

        For Each key In Parsed("data")
            If key("name").contains(" days a week") Or key("name") = "Azubis" Then employee_groups(key("id")) = key("name")
        Next

        '''''''''''''''''''VACATION ACCOUNT TYPES'''''''''''
        objhttp.Open("GET", "https://openapi.planday.com/absence/v1.0/accounttypes?absenceType=Vacation", False)
        objhttp.SetRequestHeader("Authorization", "Bearer " & api_token)
        objhttp.SetRequestHeader("X-ClientId", api_client_id)
        objhttp.Send()
        Parsed = New Web.Script.Serialization.JavaScriptSerializer().Deserialize(Of Object)(objhttp.ResponseText)


        For Each key In Parsed("data")
            If vacation_types Is Nothing Then
                ReDim Preserve vacation_types(0)
                vacation_types(0) = key("id")
            Else
                If Not vacation_types.Contains(key("id")) Then
                    ReDim Preserve vacation_types(UBound(vacation_types) + 1)
                    vacation_types(UBound(vacation_types)) = key("id")
                End If
            End If
        Next

        '''''''''''''''''''OVERTIME ACCOUNT TYPES'''''''''''
        objhttp.Open("GET", "https://openapi.planday.com/absence/v1.0/accounttypes?absenceType=Flextime", False)
        objhttp.SetRequestHeader("Authorization", "Bearer " & api_token)
        objhttp.SetRequestHeader("X-ClientId", api_client_id)
        objhttp.Send()
        Parsed = New Web.Script.Serialization.JavaScriptSerializer().Deserialize(Of Object)(objhttp.ResponseText)

        For Each key In Parsed("data")
            If overtime_types Is Nothing Then
                ReDim Preserve overtime_types(0)
                overtime_types(0) = key("id")
            Else
                If Not overtime_types.Contains(key("id")) Then
                    ReDim Preserve overtime_types(UBound(overtime_types) + 1)
                    overtime_types(UBound(overtime_types)) = key("id")
                End If
            End If
        Next


        '''''''''''''''get shift_types''''''''''''''
        objhttp.Open("GET", "https://openapi.planday.com/scheduling/v1.0/shifttypes", False)
        objhttp.SetRequestHeader("Authorization", "Bearer " & api_token)
        objhttp.SetRequestHeader("X-ClientId", api_client_id)
        objhttp.Send()
        Parsed = New Web.Script.Serialization.JavaScriptSerializer().Deserialize(Of Object)(objhttp.ResponseText)

        Dim shifts_count As Integer
        If Parsed("paging").ContainsKey("total") Then shifts_count = Parsed("paging")("total")

        For y = 0 To CInt(shifts_count / 50)
            objhttp.Open("GET", "https://openapi.planday.com/scheduling/v1.0/shifttypes?limit=0&offset=" & y * 50, False)
            objhttp.SetRequestHeader("Authorization", "Bearer " & api_token)
            objhttp.SetRequestHeader("X-ClientId", api_client_id)
            objhttp.Send()
            Parsed = New Web.Script.Serialization.JavaScriptSerializer().Deserialize(Of Object)(objhttp.ResponseText)
            Debug.Print(objhttp.ResponseText)
            For Each key In Parsed("data")
                shift_types(key("id")) = key("name")
                ''''''''''
                If PLANDAY_portal = "DE" Then
                    If My.Settings.DE_shifts Is Nothing Then
                        My.Settings.DE_shifts = New Specialized.StringCollection
                        My.Settings.DE_shifts_names = New Specialized.StringCollection
                        My.Settings.DE_shifts_surcharges = New Specialized.StringCollection
                        My.Settings.DE_shifts_overtime = New Specialized.StringCollection
                        My.Settings.DE_shifts_sick_paid = New Specialized.StringCollection
                        My.Settings.DE_shifts_sick_surcharges = New Specialized.StringCollection
                    End If
                    If Not My.Settings.DE_shifts.Contains(key("id")) Then
                        If key("isActive") = True Then
                            My.Settings.DE_shifts.Add(key("id"))
                            My.Settings.DE_shifts_names.Add(key("name"))
                            My.Settings.DE_shifts_surcharges.Add("false")
                            My.Settings.DE_shifts_overtime.Add("false")
                            My.Settings.DE_shifts_sick_paid.Add("false")
                            My.Settings.DE_shifts_sick_surcharges.Add("false")
                        End If
                    Else
                        My.Settings.DE_shifts_names.Item(My.Settings.DE_shifts.IndexOf(key("id"))) = key("name")
                    End If
                End If

                If PLANDAY_portal = "HQ" Then
                    If My.Settings.HQ_shifts Is Nothing Then
                        My.Settings.HQ_shifts = New Specialized.StringCollection
                        My.Settings.HQ_shifts_names = New Specialized.StringCollection
                        My.Settings.HQ_shifts_surcharges = New Specialized.StringCollection
                        My.Settings.HQ_shifts_overtime = New Specialized.StringCollection
                        My.Settings.HQ_shifts_sick_paid = New Specialized.StringCollection
                        My.Settings.HQ_shifts_sick_surcharges = New Specialized.StringCollection
                    End If
                    If Not My.Settings.HQ_shifts.Contains(key("id")) Then
                        If key("isActive") = True Then
                            My.Settings.HQ_shifts.Add(key("id"))
                            My.Settings.HQ_shifts_names.Add(key("name"))
                            My.Settings.HQ_shifts_surcharges.Add("false")
                            My.Settings.HQ_shifts_overtime.Add("false")
                            My.Settings.HQ_shifts_sick_paid.Add("false")
                            My.Settings.HQ_shifts_sick_surcharges.Add("false")
                        End If
                    Else
                        My.Settings.HQ_shifts_names.Item(My.Settings.HQ_shifts.IndexOf(key("id"))) = key("name")
                    End If
                End If
                '''''''''
            Next

        Next
        '''''''
        Dim index As Integer
        If PLANDAY_portal = "DE" Then
            For Each shift In My.Settings.DE_shifts
                If Not shift_types.ContainsKey(shift) Then
                    index = My.Settings.DE_shifts.IndexOf(shift)
                    My.Settings.DE_shifts.RemoveAt(index)
                    My.Settings.DE_shifts_names.RemoveAt(index)
                    My.Settings.DE_shifts_surcharges.RemoveAt(index)
                    My.Settings.DE_shifts_overtime.RemoveAt(index)
                    My.Settings.DE_shifts_sick_paid.RemoveAt(index)
                    My.Settings.DE_shifts_sick_surcharges.RemoveAt(index)
                End If
            Next
        End If

        If PLANDAY_portal = "HQ" Then
            For Each shift In My.Settings.HQ_shifts
                If Not shift_types.ContainsKey(shift) Then
                    index = My.Settings.HQ_shifts.IndexOf(shift)
                    My.Settings.HQ_shifts.RemoveAt(index)
                    My.Settings.HQ_shifts_names.RemoveAt(index)
                    My.Settings.HQ_shifts_surcharges.RemoveAt(index)
                    My.Settings.HQ_shifts_overtime.RemoveAt(index)
                    My.Settings.HQ_shifts_sick_paid.RemoveAt(index)
                    My.Settings.HQ_shifts_sick_surcharges.RemoveAt(index)
                End If
            Next
        End If
        '''''''


        ''''''''''''CONTRACT RULES''''''''''''''''''
        objhttp.Open("GET", "https://openapi.planday.com/contractrules/v1.0/contractrules", False)
        objhttp.SetRequestHeader("Authorization", "Bearer " & api_token)
        objhttp.SetRequestHeader("X-ClientId", api_client_id)
        objhttp.Send()
        Parsed = New Web.Script.Serialization.JavaScriptSerializer().Deserialize(Of Object)(objhttp.ResponseText)
        Debug.Print(objhttp.ResponseText)
        For Each key In Parsed("data")
            contract_rules(key("id")) = Val(key("name"))
        Next

        ''''''''''''GET List of departments''''''''
        objhttp.Open("GET", "https://openapi.planday.com/hr/v1.0/departments", False)
        objhttp.SetRequestHeader("Authorization", "Bearer " & api_token)
        objhttp.SetRequestHeader("X-ClientId", api_client_id)
        objhttp.Send()
        Parsed = New Web.Script.Serialization.JavaScriptSerializer().Deserialize(Of Object)(objhttp.ResponseText)
        Debug.Print(objhttp.ResponseText)
        For Each key In Parsed("data")
            departments(key("id")) = key("name")
            department_numbers(key("id")) = key("number")
            '''''''''
            If PLANDAY_portal = "DE" Then
                If My.Settings.DE_departments Is Nothing Then
                    My.Settings.DE_departments = New Specialized.StringCollection
                    My.Settings.DE_departments_names = New Specialized.StringCollection
                    My.Settings.DE_departments_bundesland = New Specialized.StringCollection
                End If
                If Not My.Settings.DE_departments.Contains(key("id")) Then
                    My.Settings.DE_departments.Add(key("id"))
                    My.Settings.DE_departments_names.Add(key("name"))
                    My.Settings.DE_departments_bundesland.Add("not_set")
                Else
                    My.Settings.DE_departments_names.Item(My.Settings.DE_departments.IndexOf(key("id"))) = key("name")
                End If
            End If

            If PLANDAY_portal = "HQ" Then
                If My.Settings.HQ_departments Is Nothing Then
                    My.Settings.HQ_departments = New Specialized.StringCollection
                    My.Settings.HQ_departments_names = New Specialized.StringCollection
                    My.Settings.HQ_departments_bundesland = New Specialized.StringCollection
                End If
                If Not My.Settings.HQ_departments.Contains(key("id")) Then
                    My.Settings.HQ_departments.Add(key("id"))
                    My.Settings.HQ_departments_names.Add(key("name"))
                    My.Settings.HQ_departments_bundesland.Add("not_set")
                Else
                    My.Settings.HQ_departments_names.Item(My.Settings.HQ_departments.IndexOf(key("id"))) = key("name")
                End If
            End If
            '''''''''
        Next
        '''''''

        If PLANDAY_portal = "DE" Then
            For Each department In My.Settings.DE_departments
                If Not departments.ContainsKey(department) Then
                    index = My.Settings.DE_departments.IndexOf(department)
                    My.Settings.DE_departments.RemoveAt(index)
                    My.Settings.DE_departments_names.RemoveAt(index)
                    My.Settings.DE_departments_bundesland.RemoveAt(index)
                End If
            Next
        End If

        If PLANDAY_portal = "HQ" Then
            For Each department In My.Settings.HQ_departments
                If Not departments.ContainsKey(department) Then
                    index = My.Settings.HQ_departments.IndexOf(department)
                    My.Settings.HQ_departments.RemoveAt(index)
                    My.Settings.HQ_departments_names.RemoveAt(index)
                    My.Settings.HQ_departments_bundesland.RemoveAt(index)
                End If
            Next
        End If
        ''''''

        '''''''employee_records'''''''''
        '''''''''''deactivated'''''''''

        objhttp.Open("GET", "https://openapi.planday.com/hr/v1.0/employees/deactivated?limit=0", False)
        objhttp.SetRequestHeader("Authorization", "Bearer " & api_token)
        objhttp.SetRequestHeader("X-ClientId", api_client_id)
        objhttp.Send()
        Parsed = New Web.Script.Serialization.JavaScriptSerializer().Deserialize(Of Object)(objhttp.ResponseText)

        Dim employees_count_deactivated As Integer
        If Parsed("paging").ContainsKey("total") Then employees_count_deactivated = Parsed("paging")("total")

        For y = 0 To CInt(employees_count_deactivated / 50)

            objhttp.Open("GET", "https://openapi.planday.com/hr/v1.0/employees/deactivated?limit=0&offset=" & y * 50 & "&special=BirthDate", False)
            objhttp.SetRequestHeader("Authorization", "Bearer " & api_token)
            objhttp.SetRequestHeader("X-ClientId", api_client_id)
            objhttp.Send()
            Parsed = New Web.Script.Serialization.JavaScriptSerializer().Deserialize(Of Object)(objhttp.ResponseText)


            For Each value In Parsed("data")
                If deactivated Is Nothing Then
                    ReDim Preserve deactivated(0)
                    deactivated(0) = value("id")
                Else
                    ReDim Preserve deactivated(UBound(deactivated) + 1)
                    deactivated(UBound(deactivated)) = value("id")
                End If
            Next
        Next

        ''''Bundesländer
        'BW = Baden - Württemberg;
        'BY = Bayern;
        'BE = Berlin;
        'BB = Brandenburg;
        'HB = Bremen;
        'HH = Hamburg;
        'HE = Hessen;
        'MV = Mecklenburg - Vorpommern;
        'NI = Niedersachsen;
        'NW = Nordrhein - Westfalen;
        'RP = Rheinland - Pfalz;
        'SL = Saarland;
        'SN = Sachsen;
        'ST = Sachsen - Anhalt;
        'SH = Schleswig - Holstein;
        'TH = Thüringen.
        '''''''''''''''FT für departments'''''''''''''''''''''''''''''''''
        Dim J
        Dim O
        J = Year(start_monday.AddDays(6))
        'Osterberechnung
        O = Ostern(J)
        'Feiertage berechnen
        Dim EWeihnachtstag_Vorjahr As Date = DateSerial(J - 1, 12, 25)
        Dim ZWeihnachtstag_Vorjahr = DateSerial(J - 1, 12, 26)
        Dim Neujahr As Date = DateSerial(J, 1, 1)
        Dim Heilige_3_Koenige As Date = DateSerial(J, 1, 6)
        Dim Frauentag As Date = DateSerial(J, 3, 8)
        Dim Karfreitag As Date = DateAdd("D", -2, O)
        Dim Ostermontag As Date = DateAdd("D", 1, O)
        Dim EMai As Date = DateSerial(J, 5, 1)
        Dim Tag_der_Befreiung As Date = DateSerial(J, 5, 8)
        Dim Christi_Himmelfahrt As Date = DateAdd("D", 39, O)
        Dim Pfingstmontag As Date = DateAdd("D", 50, O)
        Dim Fronleichnam As Date = DateAdd("D", 60, O)
        Dim Maria_Himmelfahrt As Date = DateSerial(J, 8, 15)
        Dim Tag_der_deutschen_Einheit As Date = DateSerial(J, 10, 3)
        Dim Reformationstag As Date = DateSerial(J, 10, 31)
        Dim Allerheiligen As Date = DateSerial(J, 11, 1)
        Dim Buss_und_Bettag As Date = DateSerial(J, 11, 22).AddDays(-(DateSerial(J, 11, 18)).ToOADate Mod 7)
        Dim EWeihnachtstag As Date = DateSerial(J, 12, 25)
        Dim ZWeihnachtstag As Date = DateSerial(J, 12, 26)

        FT_Name(EWeihnachtstag_Vorjahr) = "1. Weihnachtsfeiertag"
        FT_Name(ZWeihnachtstag_Vorjahr) = "2. Weihnachtsfeiertag"
        FT_Name(Neujahr) = "Neujahr"
        FT_Name(Heilige_3_Koenige) = "Heilige Drei Könige"
        FT_Name(Frauentag) = "Frauentag"
        FT_Name(Karfreitag) = "Karfreitag"
        FT_Name(Ostermontag) = "Ostermontag"
        FT_Name(EMai) = "Tag der Arbeit"
        If J = 2020 Then FT_Name(Tag_der_Befreiung) = "Tag der Befreiung"
        FT_Name(Christi_Himmelfahrt) = "Christi Himmelfahrt"
        FT_Name(Pfingstmontag) = "Pfingstmontag"
        FT_Name(Fronleichnam) = "Fronleichnam"
        FT_Name(Maria_Himmelfahrt) = "Mariä Himmelfahrt"
        FT_Name(Tag_der_deutschen_Einheit) = "Tag der Deutschen Einheit"
        FT_Name(Reformationstag) = "Reformationstag"
        FT_Name(Allerheiligen) = "Allerheiligen"
        FT_Name(Buss_und_Bettag) = "Buß - und Bettag"
        FT_Name(EWeihnachtstag) = "1. Weihnachtsfeiertag"
        FT_Name(ZWeihnachtstag) = "2. Weihnachtsfeiertag"

        For Each department_id In departments.Keys


            If departments(department_id).Contains("BER_") Or departments(department_id).Contains("Administration") Then
                If J = 2020 Then
                    departments_FT(department_id) = New Date() {EWeihnachtstag_Vorjahr, ZWeihnachtstag_Vorjahr, Neujahr, Frauentag,
                            Karfreitag, Ostermontag, EMai, Tag_der_Befreiung, Christi_Himmelfahrt, Pfingstmontag,
                            Tag_der_deutschen_Einheit, EWeihnachtstag, ZWeihnachtstag}
                Else
                    departments_FT(department_id) = New Date() {EWeihnachtstag_Vorjahr, ZWeihnachtstag_Vorjahr, Neujahr, Frauentag,
                            Karfreitag, Ostermontag, EMai, Christi_Himmelfahrt, Pfingstmontag,
                            Tag_der_deutschen_Einheit, EWeihnachtstag, ZWeihnachtstag}
                End If
            End If

            If departments(department_id).Contains("FRA_") Then

                departments_FT(department_id) = New Date() {EWeihnachtstag_Vorjahr, ZWeihnachtstag_Vorjahr, Neujahr,
                            Karfreitag, Ostermontag, EMai, Fronleichnam, Christi_Himmelfahrt, Pfingstmontag,
                            Tag_der_deutschen_Einheit, EWeihnachtstag, ZWeihnachtstag}
            End If

            If departments(department_id).Contains("HAM_") Then

                departments_FT(department_id) = New Date() {EWeihnachtstag_Vorjahr, ZWeihnachtstag_Vorjahr, Neujahr,
                            Karfreitag, Ostermontag, EMai, Christi_Himmelfahrt, Pfingstmontag,
                            Tag_der_deutschen_Einheit, Reformationstag, EWeihnachtstag, ZWeihnachtstag}
            End If

            If departments(department_id).Contains("HEI_") Then

                departments_FT(department_id) = New Date() {EWeihnachtstag_Vorjahr, ZWeihnachtstag_Vorjahr, Neujahr, Heilige_3_Koenige,
                            Karfreitag, Ostermontag, EMai, Fronleichnam, Christi_Himmelfahrt, Pfingstmontag,
                            Tag_der_deutschen_Einheit, Allerheiligen, EWeihnachtstag, ZWeihnachtstag}
            End If

            If departments(department_id).Contains("LEI_") Then

                departments_FT(department_id) = New Date() {EWeihnachtstag_Vorjahr, ZWeihnachtstag_Vorjahr, Neujahr,
                            Karfreitag, Ostermontag, EMai, Christi_Himmelfahrt, Pfingstmontag,
                            Tag_der_deutschen_Einheit, Reformationstag, Buss_und_Bettag, EWeihnachtstag, ZWeihnachtstag}
            End If

            If departments(department_id).Contains("MUC_") Then

                departments_FT(department_id) = New Date() {EWeihnachtstag_Vorjahr, ZWeihnachtstag_Vorjahr, Neujahr, Heilige_3_Koenige,
                            Karfreitag, Ostermontag, EMai, Fronleichnam, Christi_Himmelfahrt, Pfingstmontag, Maria_Himmelfahrt,
                            Tag_der_deutschen_Einheit, Allerheiligen, EWeihnachtstag, ZWeihnachtstag}
            End If

            If departments(department_id).Contains("BRE_") Then

                departments_FT(department_id) = New Date() {EWeihnachtstag_Vorjahr, ZWeihnachtstag_Vorjahr, Neujahr,
                            Karfreitag, Ostermontag, EMai, Christi_Himmelfahrt, Pfingstmontag,
                            Tag_der_deutschen_Einheit, Reformationstag, EWeihnachtstag, ZWeihnachtstag}
            End If

        Next


    End Sub

    Sub init_FT(year As Integer)
        '''''''''''''''FT für departments'''''''''''''''''''''''''''''''''
        Dim J
        Dim O
        J = year
        'Osterberechnung
        O = Ostern(J)
        'Feiertage berechnen
        Dim EWeihnachtstag_Vorjahr As Date = DateSerial(J - 1, 12, 25)
        Dim ZWeihnachtstag_Vorjahr = DateSerial(J - 1, 12, 26)
        Dim Neujahr As Date = DateSerial(J, 1, 1)
        Dim Heilige_3_Koenige As Date = DateSerial(J, 1, 6)
        Dim Frauentag As Date = DateSerial(J, 3, 8)
        Dim Karfreitag As Date = DateAdd("D", -2, O)
        Dim Ostermontag As Date = DateAdd("D", 1, O)
        Dim EMai As Date = DateSerial(J, 5, 1)
        Dim Tag_der_Befreiung As Date = DateSerial(J, 5, 8)
        Dim Christi_Himmelfahrt As Date = DateAdd("D", 39, O)
        Dim Pfingstmontag As Date = DateAdd("D", 50, O)
        Dim Fronleichnam As Date = DateAdd("D", 60, O)
        Dim Maria_Himmelfahrt As Date = DateSerial(J, 8, 15)
        Dim Tag_der_deutschen_Einheit As Date = DateSerial(J, 10, 3)
        Dim Reformationstag As Date = DateSerial(J, 10, 31)
        Dim Allerheiligen As Date = DateSerial(J, 11, 1)
        Dim Buss_und_Bettag As Date = DateSerial(J, 11, 22).AddDays(-(DateSerial(J, 11, 18)).ToOADate Mod 7)
        Dim EWeihnachtstag As Date = DateSerial(J, 12, 25)
        Dim ZWeihnachtstag As Date = DateSerial(J, 12, 26)

        FT_Name(EWeihnachtstag_Vorjahr) = "1. Weihnachtsfeiertag"
        FT_Name(ZWeihnachtstag_Vorjahr) = "2. Weihnachtsfeiertag"
        FT_Name(Neujahr) = "Neujahr"
        FT_Name(Heilige_3_Koenige) = "Heilige Drei Könige"
        FT_Name(Frauentag) = "Frauentag"
        FT_Name(Karfreitag) = "Karfreitag"
        FT_Name(Ostermontag) = "Ostermontag"
        FT_Name(EMai) = "Tag der Arbeit"
        If year = 2020 Then FT_Name(Tag_der_Befreiung) = "Tag der Befreiung"
        FT_Name(Christi_Himmelfahrt) = "Christi Himmelfahrt"
        FT_Name(Pfingstmontag) = "Pfingstmontag"
        FT_Name(Fronleichnam) = "Fronleichnam"
        FT_Name(Maria_Himmelfahrt) = "Mariä Himmelfahrt"
        FT_Name(Tag_der_deutschen_Einheit) = "Tag der Deutschen Einheit"
        FT_Name(Reformationstag) = "Reformationstag"
        FT_Name(Allerheiligen) = "Allerheiligen"
        FT_Name(Buss_und_Bettag) = "Buß - und Bettag"
        FT_Name(EWeihnachtstag) = "1. Weihnachtsfeiertag"
        FT_Name(ZWeihnachtstag) = "2. Weihnachtsfeiertag"

        For Each department_id In departments.Keys


            If departments(department_id).Contains("BER_") Or departments(department_id).Contains("Administration") Then
                If year = 2020 Then
                    departments_FT(department_id) = New Date() {EWeihnachtstag_Vorjahr, ZWeihnachtstag_Vorjahr, Neujahr, Frauentag,
                            Karfreitag, Ostermontag, EMai, Tag_der_Befreiung, Christi_Himmelfahrt, Pfingstmontag,
                            Tag_der_deutschen_Einheit, EWeihnachtstag, ZWeihnachtstag}
                Else
                    departments_FT(department_id) = New Date() {EWeihnachtstag_Vorjahr, ZWeihnachtstag_Vorjahr, Neujahr, Frauentag,
                            Karfreitag, Ostermontag, EMai, Christi_Himmelfahrt, Pfingstmontag,
                            Tag_der_deutschen_Einheit, EWeihnachtstag, ZWeihnachtstag}
                End If
            End If

            If departments(department_id).Contains("FRA_") Then

                departments_FT(department_id) = New Date() {EWeihnachtstag_Vorjahr, ZWeihnachtstag_Vorjahr, Neujahr,
                            Karfreitag, Ostermontag, EMai, Fronleichnam, Christi_Himmelfahrt, Pfingstmontag,
                            Tag_der_deutschen_Einheit, EWeihnachtstag, ZWeihnachtstag}
            End If

            If departments(department_id).Contains("HAM_") Then

                departments_FT(department_id) = New Date() {EWeihnachtstag_Vorjahr, ZWeihnachtstag_Vorjahr, Neujahr,
                            Karfreitag, Ostermontag, EMai, Christi_Himmelfahrt, Pfingstmontag,
                            Tag_der_deutschen_Einheit, Reformationstag, EWeihnachtstag, ZWeihnachtstag}
            End If

            If departments(department_id).Contains("HEI_") Then

                departments_FT(department_id) = New Date() {EWeihnachtstag_Vorjahr, ZWeihnachtstag_Vorjahr, Neujahr, Heilige_3_Koenige,
                            Karfreitag, Ostermontag, EMai, Fronleichnam, Christi_Himmelfahrt, Pfingstmontag,
                            Tag_der_deutschen_Einheit, Allerheiligen, EWeihnachtstag, ZWeihnachtstag}
            End If

            If departments(department_id).Contains("LEI_") Then

                departments_FT(department_id) = New Date() {EWeihnachtstag_Vorjahr, ZWeihnachtstag_Vorjahr, Neujahr,
                            Karfreitag, Ostermontag, EMai, Christi_Himmelfahrt, Pfingstmontag,
                            Tag_der_deutschen_Einheit, Reformationstag, Buss_und_Bettag, EWeihnachtstag, ZWeihnachtstag}
            End If

            If departments(department_id).Contains("MUC_") Then

                departments_FT(department_id) = New Date() {EWeihnachtstag_Vorjahr, ZWeihnachtstag_Vorjahr, Neujahr, Heilige_3_Koenige,
                            Karfreitag, Ostermontag, EMai, Fronleichnam, Christi_Himmelfahrt, Pfingstmontag, Maria_Himmelfahrt,
                            Tag_der_deutschen_Einheit, Allerheiligen, EWeihnachtstag, ZWeihnachtstag}
            End If

            If departments(department_id).Contains("BRE_") Then

                departments_FT(department_id) = New Date() {EWeihnachtstag_Vorjahr, ZWeihnachtstag_Vorjahr, Neujahr,
                            Karfreitag, Ostermontag, EMai, Christi_Himmelfahrt, Pfingstmontag,
                            Tag_der_deutschen_Einheit, Reformationstag, EWeihnachtstag, ZWeihnachtstag}
            End If

        Next

    End Sub

    Private Sub LOAD_button_Click(sender As Object, e As EventArgs) Handles LOAD_button.Click
        init_FT(Year(CDate(Strings.Right(Strings.Right(select_kw.Text, 21), 10))))
        Me.select_employees.Items.Clear()
        Erase employees_records
        'Show Loading Indicator
        If normalDPI Then
            WorkingForm.Show(Me)
        Else
            LoadingForm.Show(Me)
        End If
        'Do working in bgw_loading_DoWork
        bgw_loading.RunWorkerAsync()
    End Sub

    Private Sub bgw_loading_DoWork(sender As Object, e As DoWorkEventArgs) Handles bgw_loading.DoWork

        Dim Parsed As Dictionary(Of String, Object)
        Dim Parsed_temp As Dictionary(Of String, Object)
        Dim value As Dictionary(Of String, Object)
        Dim objhttp As Object = CreateObject("MSXML2.XMLHTTP")

        '''''''''''''''get employees_count''''''''''''''
        objhttp.Open("GET", "https://openapi.planday.com/hr/v1.0/employees?limit=0", False)
        objhttp.SetRequestHeader("Authorization", "Bearer " & api_token)
        objhttp.SetRequestHeader("X-ClientId", api_client_id)
        objhttp.Send()
        Parsed = New Web.Script.Serialization.JavaScriptSerializer().Deserialize(Of Object)(objhttp.ResponseText)

        Dim employees_count As Integer
        If Parsed("paging").ContainsKey("total") Then employees_count = Parsed("paging")("total")

        For y = 0 To CInt(employees_count / 50)

            objhttp.Open("GET", "https://openapi.planday.com/hr/v1.0/employees?limit=0&offset=" & y * 50 & "&special=BirthDate", False)
            objhttp.SetRequestHeader("Authorization", "Bearer " & api_token)
            objhttp.SetRequestHeader("X-ClientId", api_client_id)
            objhttp.Send()
            Parsed = New Web.Script.Serialization.JavaScriptSerializer().Deserialize(Of Object)(objhttp.ResponseText)

            For Each value In Parsed("data")
                For Each Values In value("departments")
                    If Values = hotel_id And
                       value("userName") <> "markus.pettinger@meininger-hotels.com" And
                        value("userName") <> "daniela.simicguel@meininger-hotels.com" Then

                        employees(Trim(value("firstName")) & " " & Trim(value("lastName"))) = value("id")
                        BarTotalCount = BarTotalCount + 1
                    End If
                Next
            Next

        Next

        If Not Me.active_employees.Checked Then
            '''''''''''''deactivated'''''''''''''''
            For Each id In deactivated
                objhttp.Open("GET", "https://openapi.planday.com/hr/v1.0/employees/" & id & "?special=BirthDate", False)
                objhttp.SetRequestHeader("Authorization", "Bearer " & api_token)
                objhttp.SetRequestHeader("X-ClientId", api_client_id)
                objhttp.Send()
                Parsed = New Web.Script.Serialization.JavaScriptSerializer().Deserialize(Of Object)(objhttp.ResponseText)

                If Parsed("data").ContainsKey("primaryDepartmentId") Then
                    If Parsed("data")("primaryDepartmentId") = hotel_id Then

                        Dim hiredFrom As String = ""
                        If Parsed("data").ContainsKey("hiredFrom") Then hiredFrom = Parsed("data")("hiredFrom")
                        Dim deactivationDate As String = ""
                        If Parsed("data").ContainsKey("deactivationDate") Then deactivationDate = Parsed("data")("deactivationDate")
                        Dim days_per_week As Integer
                        If Parsed("data").ContainsKey("custom_221250") Then days_per_week = Parsed("data")("custom_221250")("value")

                        objhttp.Open("GET", "https://openapi.planday.com/contractrules/v1.0/employees/" & id, False)
                        objhttp.SetRequestHeader("Authorization", "Bearer " & api_token)
                        objhttp.SetRequestHeader("X-ClientId", api_client_id)
                        objhttp.Send()
                        Parsed_temp = New Web.Script.Serialization.JavaScriptSerializer().Deserialize(Of Object)(objhttp.ResponseText)

                        Dim id_contract_rule As Integer
                        If objhttp.ResponseText = "{}" Then
                            id_contract_rule = 3527
                        Else
                            id_contract_rule = Parsed_temp("data")("id")
                        End If


                        Dim contract_rule As Integer
                        contract_rule = contract_rules(id_contract_rule)

                        Dim shifts_approved As Boolean
                        shifts_approved = all_shifts_approved(id, hotel_id, Format(start_monday, "yyyy-MM-dd"), Format(start_monday.AddDays(6), "yyyy-MM-dd"))
                        shifts_approved = shifts_approved Or Not shiftsApproved.Checked

                        If contract_rule > 0 And If(deactivationDate = "", False, CDate(deactivationDate) >= start_monday.AddDays(7)) And If(hiredFrom = "", False, CDate(hiredFrom) < start_monday) And shifts_approved Then
                            employees(Trim(Parsed("data")("firstName")) & " " & Trim(Parsed("data")("lastName"))) = Parsed("data")("id")
                            BarTotalCount = BarTotalCount + 1
                        End If
                    End If
                End If
            Next
        End If

        'Loading complete
        bgw_loading.ReportProgress(100)
    End Sub

    Private Sub bgw_loading_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles bgw_loading.RunWorkerCompleted
        If normalDPI Then
            WorkingForm.Close()
        Else
            'Close Loading indicator
            LoadingForm.Close()
        End If
        'Initialize ProgressBar
        ProgressBarForm.TotalCount = BarTotalCount
        'Do work in
        bgw.RunWorkerAsync()
        ProgressBarForm.ShowDialog(Me)

    End Sub

    Private Sub bgw_DoWork(sender As Object, e As DoWorkEventArgs) Handles bgw.DoWork

        Dim Parsed As Dictionary(Of String, Object)
        Dim Parsed_temp As Dictionary(Of String, Object)
        Dim value As Dictionary(Of String, Object)
        Dim objhttp As Object = CreateObject("MSXML2.XMLHTTP")

        '''''''''''''''get employees_count''''''''''''''
        objhttp.Open("GET", "https://openapi.planday.com/hr/v1.0/employees?limit=0", False)
        objhttp.SetRequestHeader("Authorization", "Bearer " & api_token)
        objhttp.SetRequestHeader("X-ClientId", api_client_id)
        objhttp.Send()
        Parsed = New Web.Script.Serialization.JavaScriptSerializer().Deserialize(Of Object)(objhttp.ResponseText)

        Dim employees_count As Integer
        If Parsed("paging").ContainsKey("total") Then employees_count = Parsed("paging")("total")

        For y = 0 To CInt(employees_count / 50)

            objhttp.Open("GET", "https://openapi.planday.com/hr/v1.0/employees?limit=0&offset=" & y * 50 & "&special=BirthDate", False)
            objhttp.SetRequestHeader("Authorization", "Bearer " & api_token)
            objhttp.SetRequestHeader("X-ClientId", api_client_id)
            objhttp.Send()
            Parsed = New Web.Script.Serialization.JavaScriptSerializer().Deserialize(Of Object)(objhttp.ResponseText)



            For Each value In Parsed("data")
                For Each Values In value("departments")
                    If Values = hotel_id And
                       value("userName") <> "markus.pettinger@meininger-hotels.com" And
                        value("userName") <> "daniela.simicguel@meininger-hotels.com" Then

                        'Do Action
                        CurrentAction = CurrentAction + 1
                        ActionName = "Processing " & Trim(value("firstName")) & " " & Trim(value("lastName")) & "... "
                        'Report current status in %
                        bgw.ReportProgress(100 * CurrentAction / BarTotalCount)

                        objhttp.Open("GET", "https://openapi.planday.com/hr/v1.0/employees/" & value("id") & "?special=BirthDate", False)
                        objhttp.SetRequestHeader("Authorization", "Bearer " & api_token)
                        objhttp.SetRequestHeader("X-ClientId", api_client_id)
                        objhttp.Send()
                        Parsed_temp = New Web.Script.Serialization.JavaScriptSerializer().Deserialize(Of Object)(objhttp.ResponseText)

                        Dim hiredFrom As String = ""
                        If Parsed_temp("data").ContainsKey("hiredFrom") Then hiredFrom = Parsed_temp("data")("hiredFrom")
                        Dim days_per_week As Integer
                        If Parsed_temp("data").ContainsKey("custom_221250") Then days_per_week = Parsed_temp("data")("custom_221250")("value")

                        objhttp.Open("GET", "https://openapi.planday.com/contractrules/v1.0/employees/" & value("id"), False)
                        objhttp.SetRequestHeader("Authorization", "Bearer " & api_token)
                        objhttp.SetRequestHeader("X-ClientId", api_client_id)
                        objhttp.Send()
                        Parsed_temp = New Web.Script.Serialization.JavaScriptSerializer().Deserialize(Of Object)(objhttp.ResponseText)

                        Dim id_contract_rule As Integer
                        If objhttp.ResponseText = "{}" Then
                            id_contract_rule = 3527
                        Else
                            id_contract_rule = Parsed_temp("data")("id")
                        End If


                        Dim contract_rule As Integer
                        contract_rule = contract_rules(id_contract_rule)

                        Dim shifts_approved As Boolean
                        shifts_approved = all_shifts_approved(value("id"), Values, Format(start_monday, "yyyy-MM-dd"), Format(start_monday.AddDays(6), "yyyy-MM-dd"))
                        Dim correction As Double
                        If Not shiftsApproved.Checked And Not shifts_approved Then correction = special_hours(value("id"), Values)
                        shifts_approved = shifts_approved Or Not shiftsApproved.Checked

                        If contract_rule > 0 And If(hiredFrom = "", False, CDate(hiredFrom) < start_monday) And shifts_approved Then
                            Dim datum As Date = start_monday
                            Dim temp As Integer

                            Do While datum <= start_monday.AddDays(6)
                                If FT_found(datum, hotel_id) And Weekday(datum, vbMonday) <> 7 Then
                                    temp = temp + 1
                                End If
                                datum = datum.AddDays(1)
                            Loop

                            current_employee = Trim(value("firstName")) & " " & Trim(value("lastName"))
                            If Len(current_employee) > 31 Then
                                current_employee = Strings.Left(current_employee, 31)
                            End If
                            Dim KW_written_overtime As Double
                            KW_written_overtime = get_overtimeKW(value("id"), Format(start_monday, "dd.MM.yyyy") & "-" & Format(start_monday.AddDays(6), "dd.MM.yyyy"))
                            Dim KW_written_overtime_PLANDAY As Double
                            KW_written_overtime_PLANDAY = get_overtimeKW_PLANDAY(value("id"), Format(start_monday, "yyyy-MM-dd") & "-" & Format(start_monday.AddDays(6), "yyyy-MM-dd")) - correction
                            Dim KW_IST As Double
                            KW_IST = Math.Round(IST_hours(value("id"), Values, contract_rule), 2, MidpointRounding.AwayFromZero)

                            If days_per_week > 0 Then contract_rule = contract_rule - contract_rule / days_per_week * temp

                            If employees_records Is Nothing Then
                                ReDim Preserve employees_records(0)
                                employees_records(0) = New EmployeeRecord(value("id"), Trim(value("firstName")) & " " & Trim(value("lastName")),
                                                                          contract_rule, KW_IST, KW_IST - contract_rule, KW_IST - contract_rule - KW_written_overtime - KW_written_overtime_PLANDAY,
                                                                          KW_written_overtime, KW_written_overtime_PLANDAY, overtime_Balance(value("id")))
                            Else
                                ReDim Preserve employees_records(UBound(employees_records) + 1)
                                employees_records(UBound(employees_records)) = New EmployeeRecord(value("id"), Trim(value("firstName")) & " " & Trim(value("lastName")),
                                                                          contract_rule, KW_IST, KW_IST - contract_rule, KW_IST - contract_rule - KW_written_overtime - KW_written_overtime_PLANDAY,
                                                                          KW_written_overtime, KW_written_overtime_PLANDAY, overtime_Balance(value("id")))

                            End If

                            KW_written_overtime = 0
                            KW_written_overtime_PLANDAY = 0
                            KW_IST = 0
                            correction = 0
                            temp = 0
                            days_per_week = 0
                        Else
                            System.Threading.Thread.Sleep(50)
                        End If

                    End If
                Next
            Next
        Next

        If Not Me.active_employees.Checked Then
            '''''''''''''deactivated'''''''''''''''
            For Each id In deactivated
                objhttp.Open("GET", "https://openapi.planday.com/hr/v1.0/employees/" & id & "?special=BirthDate", False)
                objhttp.SetRequestHeader("Authorization", "Bearer " & api_token)
                objhttp.SetRequestHeader("X-ClientId", api_client_id)
                objhttp.Send()
                Parsed = New Web.Script.Serialization.JavaScriptSerializer().Deserialize(Of Object)(objhttp.ResponseText)

                If Parsed("data").ContainsKey("primaryDepartmentId") Then
                    If Parsed("data")("primaryDepartmentId") = hotel_id Then

                        Dim hiredFrom As String = ""
                        If Parsed("data").ContainsKey("hiredFrom") Then hiredFrom = Parsed("data")("hiredFrom")
                        Dim deactivationDate As String = ""
                        If Parsed("data").ContainsKey("deactivationDate") Then deactivationDate = Parsed("data")("deactivationDate")
                        Dim days_per_week As Integer
                        If Parsed("data").ContainsKey("custom_221250") Then days_per_week = Parsed("data")("custom_221250")("value")

                        objhttp.Open("GET", "https://openapi.planday.com/contractrules/v1.0/employees/" & id, False)
                        objhttp.SetRequestHeader("Authorization", "Bearer " & api_token)
                        objhttp.SetRequestHeader("X-ClientId", api_client_id)
                        objhttp.Send()
                        Parsed_temp = New Web.Script.Serialization.JavaScriptSerializer().Deserialize(Of Object)(objhttp.ResponseText)

                        Dim id_contract_rule As Integer
                        If objhttp.ResponseText = "{}" Then
                            id_contract_rule = 3527
                        Else
                            id_contract_rule = Parsed_temp("data")("id")
                        End If


                        Dim contract_rule As Integer
                        contract_rule = contract_rules(id_contract_rule)

                        Dim shifts_approved As Boolean
                        shifts_approved = all_shifts_approved(id, hotel_id, Format(start_monday, "yyyy-MM-dd"), Format(start_monday.AddDays(6), "yyyy-MM-dd"))
                        shifts_approved = shifts_approved Or Not shiftsApproved.Checked

                        If contract_rule > 0 And If(deactivationDate = "", False, CDate(deactivationDate) >= start_monday.AddDays(7)) And If(hiredFrom = "", False, CDate(hiredFrom) < start_monday) And shifts_approved Then

                            'Do Action
                            CurrentAction = CurrentAction + 1
                            ActionName = "Processing " & Trim(Parsed("data")("firstName")) & " " & Trim(Parsed("data")("lastName")) & "... "
                            'Report current status in %
                            bgw.ReportProgress(100 * CurrentAction / BarTotalCount)

                            Dim correction As Double
                            If Not shiftsApproved.Checked And Not shifts_approved Then correction = special_hours(id, hotel_id)

                            Dim datum As Date = start_monday
                            Dim temp As Integer

                            Do While datum <= start_monday.AddDays(6)
                                If FT_found(datum, hotel_id) And Weekday(datum, vbMonday) <> 7 Then
                                    temp = temp + 1
                                End If
                                datum = datum.AddDays(1)
                            Loop

                            current_employee = Trim(Parsed("data")("firstName")) & " " & Trim(Parsed("data")("lastName"))
                            If Len(current_employee) > 31 Then
                                current_employee = Strings.Left(current_employee, 31)
                            End If
                            Dim KW_written_overtime As Double
                            KW_written_overtime = get_overtimeKW(id, Format(start_monday, "dd.MM.yyyy") & "-" & Format(start_monday.AddDays(6), "dd.MM.yyyy"))
                            Dim KW_written_overtime_PLANDAY As Double
                            KW_written_overtime_PLANDAY = get_overtimeKW_PLANDAY(id, Format(start_monday, "yyyy-MM-dd") & "-" & Format(start_monday.AddDays(6), "yyyy-MM-dd")) - correction
                            Dim KW_IST As Double
                            KW_IST = Math.Round(IST_hours(id, hotel_id, contract_rule), 2, MidpointRounding.AwayFromZero)

                            If days_per_week > 0 Then contract_rule = contract_rule - contract_rule / days_per_week * temp

                            If employees_records Is Nothing Then
                                ReDim Preserve employees_records(0)
                                employees_records(0) = New EmployeeRecord(id, Trim(Parsed("data")("firstName")) & " " & Trim(Parsed("data")("lastName")),
                                                                          contract_rule, KW_IST, KW_IST - contract_rule, KW_IST - contract_rule - KW_written_overtime - KW_written_overtime_PLANDAY,
                                                                          KW_written_overtime, KW_written_overtime_PLANDAY, overtime_Balance(id))
                            Else
                                ReDim Preserve employees_records(UBound(employees_records) + 1)
                                employees_records(UBound(employees_records)) = New EmployeeRecord(id, Trim(Parsed("data")("firstName")) & " " & Trim(Parsed("data")("lastName")),
                                                                          contract_rule, KW_IST, KW_IST - contract_rule, KW_IST - contract_rule - KW_written_overtime - KW_written_overtime_PLANDAY,
                                                                          KW_written_overtime, KW_written_overtime_PLANDAY, overtime_Balance(id))

                            End If

                            KW_written_overtime = 0
                            KW_written_overtime_PLANDAY = 0
                            KW_IST = 0
                            correction = 0
                            temp = 0
                            days_per_week = 0
                        End If
                    End If
                End If
            Next
        End If

        ImChangingStuff = False
    End Sub

    Private Sub bgw_ProgressChanged(sender As Object, e As ProgressChangedEventArgs) Handles bgw.ProgressChanged
        ProgressBarForm.NextAction(CurrentAction, ActionName)
    End Sub

    Private Sub bgw_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles bgw.RunWorkerCompleted
        CurrentAction = 0
        BarTotalCount = 0
        ProgressBarForm.Complete()
        Fill_LV()
    End Sub

    Public Function getDepartmentID(department_name As String) As String
        For Each department_id In departments.Keys
            If departments(department_id) = department_name Then
                Return department_id
                Exit Function
            End If
        Next
        Return ""
    End Function

    Public Function overtime_Balance(employee_id As String) As Double
        Dim result As Double

        Dim overtime_accounts()

        Dim Parsed As Dictionary(Of String, Object)
        Dim Values As Dictionary(Of String, Object)
        Dim value As Dictionary(Of String, Object)
        Dim objhttp As Object = CreateObject("MSXML2.XMLHTTP")

        objhttp.Open("GET", "https://openapi.planday.com/absence/v1.0/accounts?employeeId=" & employee_id & "&status=Active", False)
        objhttp.SetRequestHeader("Authorization", "Bearer " & api_token)
        objhttp.SetRequestHeader("X-ClientId", api_client_id)
        objhttp.Send()
        Parsed = New Web.Script.Serialization.JavaScriptSerializer().Deserialize(Of Object)(objhttp.ResponseText)

        For Each value In Parsed("data")


            If overtime_types.Contains(value("typeId")) Then

                If overtime_accounts Is Nothing Then
                    ReDim Preserve overtime_accounts(0)
                    overtime_accounts(0) = value("id")
                Else
                    If Not overtime_accounts.Contains(value("id")) Then
                        ReDim Preserve overtime_accounts(UBound(overtime_accounts) + 1)
                        overtime_accounts(UBound(overtime_accounts)) = value("id")
                    End If
                End If

            End If

        Next

        If Not overtime_accounts Is Nothing Then

            objhttp.Open("GET", "https://openapi.planday.com/absence/v1.0/accounts/" & overtime_accounts(0) & "/balance?balanceDate=" & Format(Now, "yyyy-MM-dd"), False)
            objhttp.SetRequestHeader("Authorization", "Bearer " & api_token)
            objhttp.SetRequestHeader("X-ClientId", api_client_id)
            objhttp.Send()
            Parsed = New Web.Script.Serialization.JavaScriptSerializer().Deserialize(Of Object)(objhttp.ResponseText)

            For Each Values In Parsed("data")("balance")
                result = result + Values("value")
            Next

        End If

        Return result


    End Function

    Public Function IST_hours(employee_id As String, department_id As String, contract_rule As Integer) As Double

        '''''
        Dim records() As ShiftRecord
        Dim breaks()
        '''''
        Dim days_per_week As Integer
        Dim paid_leave As Integer
        Dim working As Double
        Dim shift_type As String = ""
        Dim Parsed As Dictionary(Of String, Object)
        Dim value As Dictionary(Of String, Object)
        Dim objhttp As Object = CreateObject("MSXML2.XMLHTTP")

        objhttp.Open("GET", "https://openapi.planday.com/hr/v1.0/employees/" & employee_id & "?special=BirthDate", False)
        objhttp.SetRequestHeader("Authorization", "Bearer " & api_token)
        objhttp.SetRequestHeader("X-ClientId", api_client_id)
        objhttp.Send()
        Parsed = New Web.Script.Serialization.JavaScriptSerializer().Deserialize(Of Object)(objhttp.ResponseText)

        If Parsed("data").ContainsKey("custom_221250") Then days_per_week = Parsed("data")("custom_221250")("value")
        Dim i As Date = start_monday
        Dim BirthDate As String = Parsed("data")("birthDate")

        If days_per_week <> 0 Then paid_leave = CInt(contract_rule / days_per_week) * paid_leave_days(employee_id, Format(i, "yyyy-MM-dd"), Format(i.AddDays(6), "yyyy-MM-dd"))

        objhttp.Open("GET", "https://openapi.planday.com/scheduling/v1.0/shifts?employeeId=" & employee_id & "&shiftStatus=Approved" & "&from=" & Format(i, "yyyy-MM-dd") & "&to=" & Format(i.AddDays(6), "yyyy-MM-dd"), False)
        objhttp.SetRequestHeader("Authorization", "Bearer " & api_token)
        objhttp.SetRequestHeader("X-ClientId", api_client_id)
        objhttp.Send()
        Parsed = New Web.Script.Serialization.JavaScriptSerializer().Deserialize(Of Object)(objhttp.ResponseText)

        Dim shifts_count As Integer
        If Parsed("paging").ContainsKey("total") Then shifts_count = Parsed("paging")("total")

        For y = 0 To CInt(shifts_count / 50) '
            objhttp.Open("GET", "https://openapi.planday.com/scheduling/v1.0/shifts?employeeId=" & employee_id & "&shiftStatus=Approved" & "&from=" & Format(i, "yyyy-MM-dd") & "&to=" & Format(i.AddDays(6), "yyyy-MM-dd") & "&offset=" & y * 50, False)
            objhttp.SetRequestHeader("Authorization", "Bearer " & api_token)
            objhttp.SetRequestHeader("X-ClientId", api_client_id)
            objhttp.Send()
            Parsed = New Web.Script.Serialization.JavaScriptSerializer().Deserialize(Of Object)(objhttp.ResponseText)

            For Each value In Parsed("data")
                If (value.ContainsKey("startDateTime") And value.ContainsKey("endDateTime")) Then
                    '''''
                    Erase breaks

                    breaks = return_breaks(value("startDateTime"), value("endDateTime"), unter_18(BirthDate, Strings.Left(value("startDateTime"), 10)))
                    '''''
                    If value.ContainsKey("shiftTypeId") Then
                        shift_type = shift_types(value("shiftTypeId"))
                    Else
                        shift_type = "Normal"
                    End If

                    Select Case shift_type

                        Case "Sick leave - unpaid"
                        Case "No show"
                        'Case "FREI"
                        Case "Time off in lieu"
                        Case "Holiday"
                        Case "Anderes Haus übernimmt Schicht"
                        Case "Bank holiday (off)"
                        Case "FREI // Kurzarbeit Ausgleich"
                        Case Else
                            '''''''''
                            If records Is Nothing Then
                                ReDim Preserve records(0)
                                records(0) = New ShiftRecord
                                records(0).Shift_date = CDate(value("date"))
                                records(0).Shift_type = shift_type
                                records(0).Shift_from = Strings.Right(value("startDateTime"), 5)
                                records(0).Shift_to = Strings.Right(value("endDateTime"), 5)
                                If Not breaks Is Nothing Then
                                    records(0).Shift_break_1_from = Strings.Right(Strings.Left(breaks(0), 16), 5)
                                    records(0).Shift_break_1_to = Strings.Right(Strings.Right(breaks(0), 16), 5)
                                    If breaks.Length > 1 Then
                                        records(0).Shift_break_2_from = Strings.Right(Strings.Left(breaks(1), 16), 5)
                                        records(0).Shift_break_2_to = Strings.Right(Strings.Right(breaks(1), 16), 5)
                                    End If
                                End If
                                records(0).Shift_length = shift_length(value("startDateTime"), value("endDateTime"), unter_18(BirthDate, Strings.Left(value("startDateTime"), 10)))
                            Else
                                ReDim Preserve records(UBound(records) + 1)
                                records(UBound(records)) = New ShiftRecord
                                records(UBound(records)).Shift_date = CDate(value("date"))
                                records(UBound(records)).Shift_type = shift_type
                                records(UBound(records)).Shift_from = Strings.Right(value("startDateTime"), 5)
                                records(UBound(records)).Shift_to = Strings.Right(value("endDateTime"), 5)
                                If Not breaks Is Nothing Then
                                    records(UBound(records)).Shift_break_1_from = Strings.Right(Strings.Left(breaks(0), 16), 5)
                                    records(UBound(records)).Shift_break_1_to = Strings.Right(Strings.Right(breaks(0), 16), 5)
                                    If breaks.Length > 1 Then
                                        records(UBound(records)).Shift_break_2_from = Strings.Right(Strings.Left(breaks(1), 16), 5)
                                        records(UBound(records)).Shift_break_2_to = Strings.Right(Strings.Right(breaks(1), 16), 5)
                                    End If
                                End If
                                records(UBound(records)).Shift_length = shift_length(value("startDateTime"), value("endDateTime"), unter_18(BirthDate, Strings.Left(value("startDateTime"), 10)))
                            End If
                            ''''''''''''''''
                            working = working + shift_length(value("startDateTime"), value("endDateTime"), unter_18(BirthDate, Strings.Left(value("startDateTime"), 10)))

                    End Select
                End If
            Next

        Next
        ''''''''''
        If records Is Nothing Then
            ReDim Preserve records(0)
            records(0) = New ShiftRecord
            records(0).Shift_type = "Holiday"
            records(0).Shift_length = paid_leave
            If days_per_week <> 0 Then records(0).Holiday_days = paid_leave / CInt(contract_rule / days_per_week)
        Else
            ReDim Preserve records(UBound(records) + 1)
            records(UBound(records)) = New ShiftRecord
            records(UBound(records)).Shift_type = "Holiday"
            records(UBound(records)).Shift_length = paid_leave
            If days_per_week <> 0 Then records(UBound(records)).Holiday_days = paid_leave / CInt(contract_rule / days_per_week)
        End If
        ''''''''''''
        employee_kw_hours(current_employee) = records

        Return (working + paid_leave)


    End Function

    Public Function get_overtimeKW_PLANDAY(employee_id As String, kw_string As String) As Double
        Dim kw_overtime As Double
        Dim kw_startDate As Date
        Dim kw_endDate As Date
        Dim overtime_accounts()

        Dim Parsed As Dictionary(Of String, Object)
        Dim Values As Dictionary(Of String, Object)
        Dim value As Dictionary(Of String, Object)
        Dim objhttp As Object = CreateObject("MSXML2.XMLHTTP")

        kw_startDate = CDate(Strings.Left(kw_string, 10))
        kw_endDate = CDate(Strings.Right(kw_string, 10))

        objhttp.Open("GET", "https://openapi.planday.com/absence/v1.0/accounts?employeeId=" & employee_id & "&status=Active", False)
        objhttp.SetRequestHeader("Authorization", "Bearer " & api_token)
        objhttp.SetRequestHeader("X-ClientId", api_client_id)
        objhttp.Send()
        Parsed = New Web.Script.Serialization.JavaScriptSerializer().Deserialize(Of Object)(objhttp.ResponseText)

        For Each value In Parsed("data")

            If overtime_types.Contains(value("typeId")) Then

                If overtime_accounts Is Nothing Then
                    ReDim Preserve overtime_accounts(0)
                    overtime_accounts(0) = value("id")
                Else
                    If Not overtime_accounts.Contains(value("id")) Then
                        ReDim Preserve overtime_accounts(UBound(overtime_accounts) + 1)
                        overtime_accounts(UBound(overtime_accounts)) = value("id")
                    End If
                End If

            End If

        Next

        If Not overtime_accounts Is Nothing Then

            objhttp.Open("GET", "https://openapi.planday.com/absence/v1.0/accounts/" & overtime_accounts(0) & "/transactions?date=" & Format(Now, "yyyy-MM-dd"), False)
            objhttp.SetRequestHeader("Authorization", "Bearer " & api_token)
            objhttp.SetRequestHeader("X-ClientId", api_client_id)
            objhttp.Send()
            Parsed = New Web.Script.Serialization.JavaScriptSerializer().Deserialize(Of Object)(objhttp.ResponseText)


            For Each value In Parsed("data")
                If value.ContainsKey("note") Then
                    If value("note").contains(kw_string) Then

                        For Each Values1 In value("costs")
                            kw_overtime = kw_overtime + Values1("value")
                        Next

                    End If
                End If
                If value("type") = "Entry" And CDate(value("date")) >= kw_startDate And CDate(value("date")) <= kw_endDate Then

                    For Each Values In value("costs")
                        kw_overtime = kw_overtime + Values("value")
                    Next

                End If
            Next

        End If

        Return kw_overtime


    End Function

    Public Function get_overtimeKW(employee_id As String, kw_string As String) As Double
        Dim result As Double
        Dim overtime_accounts()
        Dim Parsed As Dictionary(Of String, Object)
        Dim Values As Dictionary(Of String, Object)
        Dim value As Dictionary(Of String, Object)
        Dim objhttp As Object = CreateObject("MSXML2.XMLHTTP")

        objhttp.Open("GET", "https://openapi.planday.com/absence/v1.0/accounts?employeeId=" & employee_id & "&status=Active", False)
        objhttp.SetRequestHeader("Authorization", "Bearer " & api_token)
        objhttp.SetRequestHeader("X-ClientId", api_client_id)
        objhttp.Send()
        Parsed = New Web.Script.Serialization.JavaScriptSerializer().Deserialize(Of Object)(objhttp.ResponseText)

        For Each value In Parsed("data")

            If overtime_types.Contains(value("typeId")) Then


                If overtime_accounts Is Nothing Then
                    ReDim Preserve overtime_accounts(0)
                    overtime_accounts(0) = value("id")
                Else
                    If Not overtime_accounts.Contains(value("id")) Then
                        ReDim Preserve overtime_accounts(UBound(overtime_accounts) + 1)
                        overtime_accounts(UBound(overtime_accounts)) = value("id")
                    End If
                End If

            End If

        Next

        If overtime_accounts Is Nothing Then
            Return 0
            Exit Function
        End If

        objhttp.Open("GET", "https://openapi.planday.com/absence/v1.0/accounts/" & overtime_accounts(0) & "/transactions?date=" & Format(Now, "yyyy-MM-dd") & "&externalId=automatic_overtime", False)
        objhttp.SetRequestHeader("Authorization", "Bearer " & api_token)
        objhttp.SetRequestHeader("X-ClientId", api_client_id)
        objhttp.Send()
        Parsed = New Web.Script.Serialization.JavaScriptSerializer().Deserialize(Of Object)(objhttp.ResponseText)

        For Each value In Parsed("data")
            If Strings.Right(value("note"), 21) = kw_string Then
                For Each Values In value("costs")
                    result = result + Values("value")
                Next
            End If
        Next

        Return result

    End Function

    Public Function FT_found_calculate(datum As Date, department_id As String) As Boolean
        If (DateAndTime.Day(datum) = 25 Or DateAndTime.Day(datum) = 26) And Month(datum) = 12 Then
            Return False
        Else
            Return departments_FT(department_id).Contains(datum)
        End If
    End Function

    Public Function FT_found(datum As Date, department_id As String) As Boolean
        Return departments_FT(department_id).Contains(datum)
    End Function

    Public Function special_hours(employee_id As String, department_id As String) As Double

        Dim Parsed As Dictionary(Of String, Object)
        Dim value As Dictionary(Of String, Object)
        Dim objhttp As Object = CreateObject("MSXML2.XMLHTTP")
        Dim shift_type As String = ""
        Dim hours As Double
        Dim shift_length_hours As Double
        Dim shift_start As Double
        Dim shift_end As Double
        Dim shift_start_date As Date
        Dim shift_end_date As Date
        Dim days_per_week As Integer

        objhttp.Open("GET", "https://openapi.planday.com/hr/v1.0/employees/" & employee_id & "?special=BirthDate", False)
        objhttp.SetRequestHeader("Authorization", "Bearer " & api_token)
        objhttp.SetRequestHeader("X-ClientId", api_client_id)
        objhttp.Send()
        Parsed = New Web.Script.Serialization.JavaScriptSerializer().Deserialize(Of Object)(objhttp.ResponseText)

        If Parsed("data").ContainsKey("custom_221250") Then days_per_week = Parsed("data")("custom_221250")("value")
        Dim i As Date = start_monday
        Dim BirthDate As String = Parsed("data")("birthDate")


        objhttp.Open("GET", "https://openapi.planday.com/scheduling/v1.0/shifts?employeeId=" & employee_id & "&from=" & Format(i, "yyyy-MM-dd") & "&to=" & Format(i.AddDays(6), "yyyy-MM-dd"), False)
        objhttp.SetRequestHeader("Authorization", "Bearer " & api_token)
        objhttp.SetRequestHeader("X-ClientId", api_client_id)
        objhttp.Send()
        Parsed = New Web.Script.Serialization.JavaScriptSerializer().Deserialize(Of Object)(objhttp.ResponseText)

        Dim shifts_count As Integer
        If Parsed("paging").ContainsKey("total") Then shifts_count = Parsed("paging")("total")

        For y = 0 To CInt(shifts_count / 50) '
            objhttp.Open("GET", "https://openapi.planday.com/scheduling/v1.0/shifts?employeeId=" & employee_id & "&from=" & Format(i, "yyyy-MM-dd") & "&to=" & Format(i.AddDays(6), "yyyy-MM-dd") & "&offset=" & y * 50, False)
            objhttp.SetRequestHeader("Authorization", "Bearer " & api_token)
            objhttp.SetRequestHeader("X-ClientId", api_client_id)
            objhttp.Send()
            Parsed = New Web.Script.Serialization.JavaScriptSerializer().Deserialize(Of Object)(objhttp.ResponseText)

            For Each value In Parsed("data")
                If (value.ContainsKey("startDateTime") And value.ContainsKey("endDateTime")) Then

                    If value.ContainsKey("shiftTypeId") Then
                        shift_type = shift_types(value("shiftTypeId"))
                    Else
                        shift_type = "Normal"
                    End If

                    Select Case shift_type

                        Case "Overtime"

                            If value("status") <> "Approved" Then hours = hours + shift_length(value("startDateTime"), value("endDateTime"), unter_18(BirthDate, Strings.Left(value("startDateTime"), 10)))

                        Case "Sick leave - Bank holiday"

                            If value("status") <> "Approved" Then hours = hours + shift_length(value("startDateTime"), value("endDateTime"), unter_18(BirthDate, Strings.Left(value("startDateTime"), 10)))

                        Case "Time off in lieu"

                            If value("status") <> "Approved" Then hours = hours - shift_length(value("startDateTime"), value("endDateTime"), unter_18(BirthDate, Strings.Left(value("startDateTime"), 10)))

                        Case "Sick leave - unpaid"

                            shift_start = TimeValue(Strings.Right(value("startDateTime"), 5)).ToOADate * 24
                            shift_end = TimeValue(Strings.Right(value("endDateTime"), 5)).ToOADate * 24
                            shift_start_date = CDate(Strings.Left(value("startDateTime"), Len(value("startDateTime")) - 6))
                            shift_end_date = CDate(Strings.Left(value("endDateTime"), Len(value("endDateTime")) - 6))

                            If shift_start > shift_end Then
                                shift_length_hours = shift_end + 24 - shift_start

                                If shift_end_date = LastDayOfMonth("SUN", DateSerial(Year(start_monday), 3, 1)) And shift_end > 3 Then
                                    shift_length_hours = shift_length_hours - 1
                                End If

                                If shift_end_date = LastDayOfMonth("SUN", DateSerial(Year(start_monday), 10, 1)) And shift_end > 3 Then
                                    shift_length_hours = shift_length_hours + 1
                                End If

                            Else
                                shift_length_hours = shift_end - shift_start

                                If shift_start_date = LastDayOfMonth("SUN", DateSerial(Year(start_monday), 3, 1)) And shift_end > 3 And shift_start < 3 Then
                                    shift_length_hours = shift_length_hours - 1
                                End If

                                If shift_start_date = LastDayOfMonth("SUN", DateSerial(Year(start_monday), 10, 1)) And shift_end > 3 And shift_start < 3 Then
                                    shift_length_hours = shift_length_hours + 1
                                End If

                            End If

                            If value("status") <> "Approved" Then hours = hours - shift_length_hours

                    End Select
                End If
            Next

        Next

        special_hours = hours

    End Function

    Public Function all_shifts_approved(employee_id As String, department_id As String, Optional start_period As String = "", Optional end_period As String = "") As Boolean
        Dim result As Boolean
        Dim begin_date As String
        Dim final_date As String

        If start_period = "" And end_period = "" Then
            begin_date = start_
            final_date = end_
        Else
            begin_date = start_period
            final_date = end_period
        End If

        Dim Parsed As Dictionary(Of String, Object)
        Dim value As Dictionary(Of String, Object)
        Dim objhttp As Object = CreateObject("MSXML2.XMLHTTP")

        objhttp.Open("GET", "https://openapi.planday.com/scheduling/v1.0/shifts?departmentId=" & department_id & "&employeeId=" & employee_id & "&from=" & begin_date & "&to=" & final_date, False)
        objhttp.SetRequestHeader("Authorization", "Bearer " & api_token)
        objhttp.SetRequestHeader("X-ClientId", api_client_id)
        objhttp.Send()
        Parsed = New Web.Script.Serialization.JavaScriptSerializer().Deserialize(Of Object)(objhttp.ResponseText)

        result = True
        For Each value In Parsed("data")
            If value("status") <> "Approved" Then
                result = False
                Exit For
            End If
        Next

        Return result

    End Function

    Public Sub adjust_overtime(employee_id As String, time_value As String, comment As String)
        Dim overtime_accounts()
        Dim Parsed As Dictionary(Of String, Object)
        Dim value As Dictionary(Of String, Object)
        Dim objhttp As Object = CreateObject("MSXML2.XMLHTTP")

        objhttp.Open("GET", "https://openapi.planday.com/absence/v1.0/accounts?employeeId=" & employee_id & "&status=Active", False)
        objhttp.SetRequestHeader("Authorization", "Bearer " & api_token)
        objhttp.SetRequestHeader("X-ClientId", api_client_id)
        objhttp.Send()
        Parsed = New Web.Script.Serialization.JavaScriptSerializer().Deserialize(Of Object)(objhttp.ResponseText)

        For Each value In Parsed("data")


            If overtime_types.Contains(value("typeId")) Then

                If overtime_accounts Is Nothing Then
                    ReDim Preserve overtime_accounts(0)
                    overtime_accounts(0) = value("id")
                Else
                    If Not overtime_accounts.Contains(value("id")) Then
                        ReDim Preserve overtime_accounts(UBound(overtime_accounts) + 1)
                        overtime_accounts(UBound(overtime_accounts)) = value("id")
                    End If
                End If

            End If

        Next

        If overtime_accounts Is Nothing Then
            Exit Sub
        End If

        Dim JSON As String
        JSON = "{""externalId"": ""automatic_overtime"",""type"": ""Adjustment"",""date"": """ & Format(Now, "yyyy-MM-dd") & """,""amounts"":[{""value"": " & Replace(time_value, ",", ".") & ",""unit"":{""type"": ""Hours""}}],""note"": """ & comment & """}"
        objhttp.Open("POST", "https://openapi.planday.com/absence/v1.0/accounts/" & overtime_accounts(0) & "/transactions", False)
        objhttp.SetRequestHeader("Content-Type", "application/json")
        objhttp.SetRequestHeader("Authorization", "Bearer " & api_token)
        objhttp.SetRequestHeader("X-ClientId", api_client_id)

        On Error Resume Next
        objhttp.Send(JSON)

    End Sub

    Public Function calculate_XMAS_125(var_start As String, var_end As String, department_id As String) As Double
        Dim result As Double
        Dim start_surcharges As Double
        Dim end_surcharges As Double
        Dim start_date As Date
        Dim end_date As Date
        Dim start_time As Double
        Dim end_time As Double
        Dim start_day As Integer
        Dim end_day As Integer

        start_surcharges = 14.0#

        start_date = CDate(Strings.Left(var_start, 10))
        end_date = CDate(Strings.Left(var_end, 10))

        start_time = TimeValue(Replace(var_start, "T", " ")).ToOADate * 24
        end_time = TimeValue(Replace(var_end, "T", " ")).ToOADate * 24
        start_day = Weekday(CDate(Replace(var_start, "T", " ")), vbMonday)
        end_day = Weekday(CDate(Replace(var_end, "T", " ")), vbMonday)

        If start_day = end_day Then

            If DateAndTime.Day(start_date) = 31 And Month(start_date) = 12 Then
                If start_time <= start_surcharges And end_time > start_surcharges Then
                    result = end_time - start_surcharges
                End If

                If start_time <= start_surcharges And end_time <= start_surcharges Then
                    'do nothing
                End If

                If start_time > start_surcharges Then
                    result = end_time - start_time
                End If
            End If

        Else

            If DateAndTime.Day(start_date) = 31 And Month(start_date) = 12 Then
                If start_time <= start_surcharges Then
                    result = 24 - start_surcharges
                Else
                    result = 24 - start_time
                End If

            End If

        End If

        Return result

    End Function
    Public Function calculate_XMAS_150(var_start As String, var_end As String, department_id As String) As Double
        Dim result As Double
        Dim start_surcharges As Double
        Dim end_surcharges As Double
        Dim start_date As Date
        Dim end_date As Date
        Dim start_time As Double
        Dim end_time As Double
        Dim start_day As Integer
        Dim end_day As Integer

        start_surcharges = 14.0#

        start_date = CDate(Strings.Left(var_start, 10))
        end_date = CDate(Strings.Left(var_end, 10))

        start_time = TimeValue(Replace(var_start, "T", " ")).ToOADate * 24
        end_time = TimeValue(Replace(var_end, "T", " ")).ToOADate * 24
        start_day = Weekday(CDate(Replace(var_start, "T", " ")), vbMonday)
        end_day = Weekday(CDate(Replace(var_end, "T", " ")), vbMonday)

        If start_day = end_day Then
            If (DateAndTime.Day(start_date) = 25 Or DateAndTime.Day(start_date) = 26) And Month(start_date) = 12 Then
                result = end_time - start_time
            End If

            If DateAndTime.Day(start_date) = 24 And Month(start_date) = 12 Then
                If start_time <= start_surcharges And end_time > start_surcharges Then
                    result = end_time - start_surcharges
                End If

                If start_time <= start_surcharges And end_time <= start_surcharges Then
                    'do nothing
                End If

                If start_time > start_surcharges Then
                    result = end_time - start_time
                End If
            End If

        Else
            If (DateAndTime.Day(start_date) = 24 Or DateAndTime.Day(start_date) = 25) And Month(start_date) = 12 Then
                If start_time <= start_surcharges Then
                    result = 24 - start_surcharges + end_time
                Else
                    result = 24 - start_time + end_time
                End If

            End If

            If DateAndTime.Day(start_date) = 26 And Month(start_date) = 12 Then
                If start_time <= start_surcharges Then
                    result = 24 - start_surcharges
                Else
                    result = 24 - start_time
                End If

            End If

        End If

        Return result

    End Function

    Public Function calculate_sunday_FT(var_start As String, var_end As String, department_id As String) As Double
        Dim result As Double
        Dim start_date As Date
        Dim end_date As Date
        Dim start_time As Double
        Dim end_time As Double
        Dim start_day As Integer
        Dim end_day As Integer

        start_date = CDate(Strings.Left(var_start, 10))
        end_date = CDate(Strings.Left(var_end, 10))

        start_time = TimeValue(Replace(var_start, "T", " ")).ToOADate * 24
        end_time = TimeValue(Replace(var_end, "T", " ")).ToOADate * 24
        start_day = Weekday(CDate(Replace(var_start, "T", " ")), vbMonday)
        end_day = Weekday(CDate(Replace(var_end, "T", " ")), vbMonday)

        If start_day = end_day Then
            If start_day = 7 Or FT_found(start_date, department_id) Then
                If (DateAndTime.Day(start_date) = 25 Or DateAndTime.Day(start_date) = 26) And Month(start_date) = 12 And Year(start_date) > 2019 Then
                    'do nothing
                Else
                    result = end_time - start_time
                End If
            Else
                'do nothing
            End If

        Else
            If start_day = 7 Or FT_found(start_date, department_id) Then
                If end_day = 7 Or FT_found(end_date, department_id) Then
                    If (DateAndTime.Day(start_date) = 25 Or DateAndTime.Day(start_date) = 26) And Month(start_date) = 12 And Year(start_date) > 2019 Then
                        'do nothing
                    Else
                        result = 24 - start_time + end_time
                    End If
                Else
                    If (DateAndTime.Day(start_date) = 25 Or DateAndTime.Day(start_date) = 26) And Month(start_date) = 12 And Year(start_date) > 2019 Then
                        'do nothing
                    Else
                        result = 24 - start_time
                    End If
                End If

            Else

                If end_day = 7 Or FT_found(end_date, department_id) Then
                    If DateAndTime.Day(start_date) = 24 And Month(start_date) = 12 And Year(start_date) > 2019 Then
                        'do nothing
                    Else
                        result = end_time
                    End If
                Else
                    'do nothing
                End If

            End If

        End If
        Return result

    End Function

    Public Function calculate_night(var_start As String, var_end As String, department_id As String, senior As Boolean) As Double
        Dim result As Double
        Dim start_surcharges As Double
        Dim end_surcharges As Double
        Dim start_date As Date
        Dim end_date As Date
        Dim start_time As Double
        Dim end_time As Double
        Dim start_day As Integer
        Dim end_day As Integer

        If senior Then
            start_surcharges = 20.0#
        Else
            start_surcharges = 22.0#
        End If
        end_surcharges = 6.0#

        start_date = CDate(Strings.Left(var_start, 10))
        end_date = CDate(Strings.Left(var_end, 10))

        start_time = TimeValue(Replace(var_start, "T", " ")).ToOADate * 24
        end_time = TimeValue(Replace(var_end, "T", " ")).ToOADate * 24
        start_day = Weekday(CDate(Replace(var_start, "T", " ")), vbMonday)
        end_day = Weekday(CDate(Replace(var_end, "T", " ")), vbMonday)


        If start_day = end_day Then
            If start_day = 7 Or FT_found(start_date, department_id) Then
                'do nothing
            Else
                If start_time >= end_surcharges And end_time <= start_surcharges Then
                    'do nothing
                End If

                If start_time >= end_surcharges And end_time > start_surcharges Then
                    result = 24 - end_time - start_surcharges
                End If

                If start_time < end_surcharges And end_time <= start_surcharges Then
                    result = end_surcharges - start_time
                End If

                If start_time < end_surcharges And end_time > start_surcharges Then
                    result = end_surcharges - start_time + end_time - start_surcharges
                End If

                If start_time < end_surcharges And end_time < end_surcharges Then
                    result = end_time - start_time
                End If

                If start_time < end_surcharges And end_time >= end_surcharges Then
                    result = end_surcharges - start_time
                End If

                If start_time <= start_surcharges And end_time > start_surcharges Then
                    result = end_time - start_surcharges
                End If

                If start_time <= start_surcharges And end_time <= start_surcharges Then
                    'do nothing
                End If

                If start_time > start_surcharges Then
                    result = end_time - start_time
                End If

            End If


        Else
            If start_day = 7 Or FT_found(start_date, department_id) Then
                If end_day = 7 Or FT_found(end_date, department_id) Then
                    'do nothing
                Else
                    If end_time < end_surcharges Then
                        result = end_time
                    Else
                        result = end_surcharges
                    End If
                End If
            Else
                If start_time <= start_surcharges Then

                    If end_day = 7 Or FT_found(end_date, department_id) Then
                        result = 24 - start_surcharges
                    Else
                        If end_time < end_surcharges Then
                            result = 24 - start_surcharges + end_time
                        Else
                            If end_time <= start_surcharges Then
                                result = 24 - start_surcharges + end_surcharges
                            Else
                                result = 24 - start_surcharges + end_surcharges + end_time - start_surcharges
                            End If
                        End If
                    End If

                    If start_time < end_surcharges Then
                        result = result + end_surcharges - start_time
                    End If

                Else
                    If end_day = 7 Or FT_found(end_date, department_id) Then
                        result = 24 - start_time
                    Else
                        If end_time < end_surcharges Then
                            result = 24 - start_time + end_time
                        Else
                            If end_time <= start_surcharges Then
                                result = 24 - start_time + end_surcharges
                            Else
                                result = 24 - start_time + end_surcharges + end_time - start_surcharges
                            End If
                        End If
                    End If

                End If
            End If



        End If
        Return result

    End Function

    Public Function paid_leave_days(employee_id As String, Optional start_d As String = "", Optional end_d As String = "") As Double
        Dim begin_date As String
        Dim final_date As String
        Dim result As Double
        Dim end_balance As Double
        Dim start_balance As Double
        Dim temp As Double

        If start_d = "" And end_d = "" Then
            begin_date = Format(CDate(start_).AddDays(-1), "yyyy-MM-dd")
            final_date = end_
        Else
            begin_date = Format(CDate(start_d).AddDays(-1), "yyyy-MM-dd")
            final_date = end_d
        End If


        Dim vacation_accounts(), vacation_accounts_next_year()
        Dim Parsed As Dictionary(Of String, Object)
        Dim value As Dictionary(Of String, Object)
        Dim objhttp As Object = CreateObject("MSXML2.XMLHTTP")

        objhttp.Open("GET", "https://openapi.planday.com/absence/v1.0/accounts?employeeId=" & employee_id, False)
        objhttp.SetRequestHeader("Authorization", "Bearer " & api_token)
        objhttp.SetRequestHeader("X-ClientId", api_client_id)
        objhttp.Send()
        Parsed = New Web.Script.Serialization.JavaScriptSerializer().Deserialize(Of Object)(objhttp.ResponseText)

        For Each value In Parsed("data")


            If vacation_types.Contains(value("typeId")) Then


                If Strings.Left(begin_date, 4) = Strings.Left(value("validityPeriod")("start"), 4) Then
                    If vacation_accounts Is Nothing Then
                        ReDim Preserve vacation_accounts(0)
                        vacation_accounts(0) = value("id")
                    Else
                        If Not vacation_accounts.Contains(value("id")) Then
                            ReDim Preserve vacation_accounts(UBound(vacation_accounts) + 1)
                            vacation_accounts(UBound(vacation_accounts)) = value("id")
                        End If
                    End If

                End If

                If Strings.Left(begin_date, 4) <> Strings.Left(final_date, 4) Then

                    If Strings.Left(final_date, 4) = Strings.Left(value("validityPeriod")("start"), 4) Then
                        If vacation_accounts_next_year Is Nothing Then
                            ReDim Preserve vacation_accounts_next_year(0)
                            vacation_accounts_next_year(0) = value("id")
                        Else
                            If Not vacation_accounts_next_year.Contains(value("id")) Then
                                ReDim Preserve vacation_accounts_next_year(UBound(vacation_accounts_next_year) + 1)
                                vacation_accounts_next_year(UBound(vacation_accounts_next_year)) = value("id")
                            End If
                        End If

                    End If

                End If

            End If

        Next

        If vacation_accounts Is Nothing Then
            result = 0
            Exit Function
        End If



        If Year(CDate(begin_date)) <> Year(CDate(final_date)) Then

            If Month(CDate(begin_date)) = 12 And DateAndTime.Day(CDate(begin_date)) = 31 Then
                begin_date = Format(CDate(begin_date).AddDays(1), "yyyy-MM-dd")
                GoTo weiter
            End If

            For Each Item In vacation_accounts
                objhttp.Open("GET", "https://openapi.planday.com/absence/v1.0/accounts/" & Item & "/balance?balanceDate=" & Format(DateSerial(Year(CDate(begin_date)), 12, 31), "yyyy-MM-dd"), False)
                objhttp.SetRequestHeader("Authorization", "Bearer " & api_token)
                objhttp.SetRequestHeader("X-ClientId", api_client_id)
                objhttp.Send()
                Parsed = New Web.Script.Serialization.JavaScriptSerializer().Deserialize(Of Object)(objhttp.ResponseText)

                For Each value In Parsed("data")("balance")
                    end_balance = end_balance + value("value")
                Next
            Next

            For Each Item In vacation_accounts
                objhttp.Open("GET", "https://openapi.planday.com/absence/v1.0/accounts/" & Item & "/balance?balanceDate=" & begin_date, False)
                objhttp.SetRequestHeader("Authorization", "Bearer " & api_token)
                objhttp.SetRequestHeader("X-ClientId", api_client_id)
                objhttp.Send()
                Parsed = New Web.Script.Serialization.JavaScriptSerializer().Deserialize(Of Object)(objhttp.ResponseText)

                For Each value In Parsed("data")("balance")
                    start_balance = start_balance + value("value")
                Next
            Next

            temp = start_balance - end_balance
            start_balance = 0
            end_balance = 0

            If vacation_accounts_next_year Is Nothing Then
                result = 0
                Exit Function
            End If

            For Each Item In vacation_accounts_next_year
                objhttp.Open("GET", "https://openapi.planday.com/absence/v1.0/accounts/" & Item & "/balance?balanceDate=" & final_date, False)
                objhttp.SetRequestHeader("Authorization", "Bearer " & api_token)
                objhttp.SetRequestHeader("X-ClientId", api_client_id)
                objhttp.Send()
                Parsed = New Web.Script.Serialization.JavaScriptSerializer().Deserialize(Of Object)(objhttp.ResponseText)

                For Each value In Parsed("data")("balance")
                    end_balance = end_balance + value("value")
                Next
            Next

            For Each Item In vacation_accounts_next_year
                objhttp.Open("GET", "https://openapi.planday.com/absence/v1.0/accounts/" & Item & "/balance?balanceDate=" & Format(DateSerial(Year(CDate(final_date)), 1, 1), "yyyy-MM-dd"), False)
                objhttp.SetRequestHeader("Authorization", "Bearer " & api_token)
                objhttp.SetRequestHeader("X-ClientId", api_client_id)
                objhttp.Send()
                Parsed = New Web.Script.Serialization.JavaScriptSerializer().Deserialize(Of Object)(objhttp.ResponseText)

                For Each value In Parsed("data")("balance")
                    start_balance = start_balance + value("value")
                Next
            Next

            result = Math.Round(temp + start_balance - end_balance, 2)

        Else
weiter:
            For Each Item In vacation_accounts
                objhttp.Open("GET", "https://openapi.planday.com/absence/v1.0/accounts/" & Item & "/balance?balanceDate=" & final_date, False)
                objhttp.SetRequestHeader("Authorization", "Bearer " & api_token)
                objhttp.SetRequestHeader("X-ClientId", api_client_id)
                objhttp.Send()
                Parsed = New Web.Script.Serialization.JavaScriptSerializer().Deserialize(Of Object)(objhttp.ResponseText)

                For Each value In Parsed("data")("balance")
                    end_balance = end_balance + value("value")
                Next
            Next

            For Each Item In vacation_accounts
                objhttp.Open("GET", "https://openapi.planday.com/absence/v1.0/accounts/" & Item & "/balance?balanceDate=" & begin_date, False)
                objhttp.SetRequestHeader("Authorization", "Bearer " & api_token)
                objhttp.SetRequestHeader("X-ClientId", api_client_id)
                objhttp.Send()
                Parsed = New Web.Script.Serialization.JavaScriptSerializer().Deserialize(Of Object)(objhttp.ResponseText)

                For Each value In Parsed("data")("balance")
                    start_balance = start_balance + value("value")
                Next
            Next
            result = Math.Round(start_balance - end_balance, 2)
        End If
        Return result

    End Function

    Public Function get_ownDepartment() As String

        Dim result As String = ""
        Dim Parsed As Dictionary(Of String, Object)
        Dim Parsed_temp As Dictionary(Of String, Object)
        Dim value As Dictionary(Of String, Object)
        Dim objhttp As Object = CreateObject("MSXML2.XMLHTTP")

        Dim user As String
        Dim oShell As Object = CreateObject("WScript.Shell")
        user = oShell.ExpandEnvironmentStrings("%UserName%")
        user = LCase(user)

        If user = "v.vankov" Or user = "vankov" Then
            user = "ventzislav.vankov"
            Return "ALL"
            Exit Function
        End If

        If user.Contains("reception") Then

            If user.Contains("am") Then result = "BER_AM_Meininger Airport Hotel BBI GmbH"
            If user.Contains("sp") Then result = "BER_SP_Meininger 10 City Hostel B.-M. GmbH"
            If user.Contains("ap") Then result = "BER_AP_Meininger Hot.Ber.Eas.Sid.Gal.GmbH"
            If user.Contains("wp") Then result = "BER_WP_Meininger Berlin Hauptbahnhof GmbH"
            If user.Contains("os") Then result = "BER_OS_Meininger Oranienburger Str. GmbH"
            If user.Contains("ts") Then result = "BER_TS_Meininger Hotel Berlin Tiergar.GmbH"
            If user.Contains("bc") Then result = "FRA_BC_Meininger Airport Frankfurt GmbH"
            If user.Contains("ea") Then result = "FRA_EA_Meininger 10 Frankfurt GmbH"
            If user.Contains("ga") Then result = "HAM_GA_Meininger 10 Hamburg GmbH"
            If user.Contains("cb") Then result = "HEI_CB_Meininger Hotel Heidelberg GmbH"
            If user.Contains("ab") Then result = "LEI_AB_Meininger Hotel Leipzig HBF GmbH"
            If user.Contains("ls") Then result = "MUC_LS_Meininger 10 Hostel+Reiseverm. GmbH"
            If user.Contains("la") Then result = "MUC_LA_Meininger Hotel Muen. Olympiap.GmbH"
            Return result
        Else

            '''''''employee_records'''''''''
            '''''''''''''''get employees_count''''''''''''''
            objhttp.Open("GET", "https://openapi.planday.com/hr/v1.0/employees?limit=0", False)
            objhttp.SetRequestHeader("Authorization", "Bearer " & api_token)
            objhttp.SetRequestHeader("X-ClientId", api_client_id)
            objhttp.Send()
            Parsed = New Web.Script.Serialization.JavaScriptSerializer().Deserialize(Of Object)(objhttp.ResponseText)

            Dim employees_count As Integer
            If Parsed("paging").ContainsKey("total") Then employees_count = Parsed("paging")("total")

            For y = 0 To CInt(employees_count / 50)

                objhttp.Open("GET", "https://openapi.planday.com/hr/v1.0/employees?limit=0&offset=" & y * 50 & "&special=BirthDate", False)
                objhttp.SetRequestHeader("Authorization", "Bearer " & api_token)
                objhttp.SetRequestHeader("X-ClientId", api_client_id)
                objhttp.Send()
                Parsed = New Web.Script.Serialization.JavaScriptSerializer().Deserialize(Of Object)(objhttp.ResponseText)


                For Each value In Parsed("data")

                    If LCase(value("email")) = user & "@meininger-hotels.com" Then
                        objhttp.Open("GET", "https://openapi.planday.com/hr/v1.0/employees/" & value("id") & "?special=BirthDate", False)
                        objhttp.SetRequestHeader("Authorization", "Bearer " & api_token)
                        objhttp.SetRequestHeader("X-ClientId", api_client_id)
                        objhttp.Send()
                        Parsed_temp = New Web.Script.Serialization.JavaScriptSerializer().Deserialize(Of Object)(objhttp.ResponseText)

                        If Parsed_temp("data").ContainsKey("departments") Then
                            For Each dep In Parsed_temp("data")("departments")
                                If departments(dep) = "Administration" Or departments(dep) = "HR" Then
                                    Return "ALL"
                                    Exit Function
                                End If
                            Next
                            If Parsed_temp("data").ContainsKey("primaryDepartmentId") Then
                                Return departments(Parsed_temp("data")("primaryDepartmentId"))
                                Exit Function
                            End If
                        End If
                    End If
                Next
            Next
            Return "NONE"
        End If

    End Function

    Public Function return_breaks(startTime As String, endTime As String, Under_18 As Boolean) As String()

        Dim temp() As String
        Dim schicht_laenge As Double
        Dim shift_start As Double
        Dim shift_end As Double
        Dim shift_start_date As Date
        Dim shift_end_date As Date

        shift_start = TimeValue(Strings.Right(startTime, 5)).ToOADate * 24
        shift_end = TimeValue(Strings.Right(endTime, 5)).ToOADate * 24
        shift_start_date = CDate(Strings.Left(startTime, Len(startTime) - 6))
        shift_end_date = CDate(Strings.Left(endTime, Len(endTime) - 6))

        If shift_start > shift_end Then
            schicht_laenge = shift_end + 24 - shift_start

            If shift_end_date = LastDayOfMonth("SUN", DateSerial(Year(start_monday), 3, 1)) And shift_end > 3 Then
                schicht_laenge = schicht_laenge - 1
            End If

            If shift_end_date = LastDayOfMonth("SUN", DateSerial(Year(start_monday), 10, 1)) And shift_end > 3 Then
                schicht_laenge = schicht_laenge + 1
            End If

        Else
            schicht_laenge = shift_end - shift_start

            If shift_start_date = LastDayOfMonth("SUN", DateSerial(Year(start_monday), 3, 1)) And shift_end > 3 And shift_start < 3 Then
                schicht_laenge = schicht_laenge - 1
            End If

            If shift_start_date = LastDayOfMonth("SUN", DateSerial(Year(start_monday), 10, 1)) And shift_end > 3 And shift_start < 3 Then
                schicht_laenge = schicht_laenge + 1
            End If

        End If

        If Under_18 Then
            If schicht_laenge > 4.5 Then
                If schicht_laenge > 6 Then
                    '+1 after 2.5 hrs
                    If temp Is Nothing Then
                        ReDim Preserve temp(0)
                        temp(0) = Format(DateAdd("n", 150, CDate(Replace(startTime, "T", " "))), "yyyy-MM-ddTHH:mm") & "-" & Format(DateAdd("n", 210, CDate(Replace(startTime, "T", " "))), "yyyy-MM-ddTHH:mm")
                    Else
                        ReDim Preserve temp(UBound(temp) + 1)
                        temp(UBound(temp)) = Format(DateAdd("n", 150, CDate(Replace(startTime, "T", " "))), "yyyy-MM-ddTHH:mm") & "-" & Format(DateAdd("n", 210, CDate(Replace(startTime, "T", " "))), "yyyy-MM-ddTHH:mm")
                    End If
                Else
                    '+0.5 after 2.5 hrs
                    If temp Is Nothing Then
                        ReDim Preserve temp(0)
                        temp(0) = Format(DateAdd("n", 150, CDate(Replace(startTime, "T", " "))), "yyyy-MM-ddTHH:mm") & "-" & Format(DateAdd("n", 180, CDate(Replace(startTime, "T", " "))), "yyyy-MM-ddTHH:mm")
                    Else
                        ReDim Preserve temp(UBound(temp) + 1)
                        temp(UBound(temp)) = Format(DateAdd("n", 150, CDate(Replace(startTime, "T", " "))), "yyyy-MM-ddTHH:mm") & "-" & Format(DateAdd("n", 180, CDate(Replace(startTime, "T", " "))), "yyyy-MM-ddTHH:mm")
                    End If
                End If
            End If
        Else
            If schicht_laenge > 6 Then
                If schicht_laenge > 9.5 Then
                    '+0.5 after 3 hrs
                    If temp Is Nothing Then
                        ReDim Preserve temp(0)
                        temp(0) = Format(DateAdd("n", 180, CDate(Replace(startTime, "T", " "))), "yyyy-MM-ddTHH:mm") & "-" & Format(DateAdd("n", 210, CDate(Replace(startTime, "T", " "))), "yyyy-MM-ddTHH:mm")
                    Else
                        ReDim Preserve temp(UBound(temp) + 1)
                        temp(UBound(temp)) = Format(DateAdd("n", 180, CDate(Replace(startTime, "T", " "))), "yyyy-MM-ddTHH:mm") & "-" & Format(DateAdd("n", 210, CDate(Replace(startTime, "T", " "))), "yyyy-MM-ddTHH:mm")
                    End If
                    '+0.25 after 5 hrs
                    If temp Is Nothing Then
                        ReDim Preserve temp(0)
                        temp(0) = Format(DateAdd("n", 300, CDate(Replace(startTime, "T", " "))), "yyyy-MM-ddTHH:mm") & "-" & Format(DateAdd("n", 315, CDate(Replace(startTime, "T", " "))), "yyyy-MM-ddTHH:mm")
                    Else
                        ReDim Preserve temp(UBound(temp) + 1)
                        temp(UBound(temp)) = Format(DateAdd("n", 300, CDate(Replace(startTime, "T", " "))), "yyyy-MM-ddTHH:mm") & "-" & Format(DateAdd("n", 315, CDate(Replace(startTime, "T", " "))), "yyyy-MM-ddTHH:mm")
                    End If
                Else
                    '+0.5 after 3 hrs
                    If temp Is Nothing Then
                        ReDim Preserve temp(0)
                        temp(0) = Format(DateAdd("n", 180, CDate(Replace(startTime, "T", " "))), "yyyy-MM-ddTHH:mm") & "-" & Format(DateAdd("n", 210, CDate(Replace(startTime, "T", " "))), "yyyy-MM-ddTHH:mm")
                    Else
                        ReDim Preserve temp(UBound(temp) + 1)
                        temp(UBound(temp)) = Format(DateAdd("n", 180, CDate(Replace(startTime, "T", " "))), "yyyy-MM-ddTHH:mm") & "-" & Format(DateAdd("n", 210, CDate(Replace(startTime, "T", " "))), "yyyy-MM-ddTHH:mm")
                    End If
                End If
            End If

        End If

        return_breaks = temp

    End Function

    Public Function shift_length(startTime As String, endTime As String, Under_18 As Boolean) As Double
        Dim result As Double
        Dim shift_start As Double
        Dim shift_end As Double
        Dim shift_start_date As Date
        Dim shift_end_date As Date

        shift_start = TimeValue(Strings.Right(startTime, 5)).ToOADate * 24
        shift_end = TimeValue(Strings.Right(endTime, 5)).ToOADate * 24
        shift_start_date = CDate(Strings.Left(startTime, Len(startTime) - 6))
        shift_end_date = CDate(Strings.Left(endTime, Len(endTime) - 6))

        If shift_start > shift_end Then
            result = shift_end + 24 - shift_start

            If shift_end_date = LastDayOfMonth("SUN", DateSerial(current_year, 3, 1)) And shift_end > 3 Then
                result = result - 1
            End If

            If shift_end_date = LastDayOfMonth("SUN", DateSerial(current_year, 10, 1)) And shift_end > 3 Then
                result = result + 1
            End If

        Else
            result = shift_end - shift_start

            If shift_start_date = LastDayOfMonth("SUN", DateSerial(current_year, 3, 1)) And shift_end > 3 And shift_start < 3 Then
                result = result - 1
            End If

            If shift_start_date = LastDayOfMonth("SUN", DateSerial(current_year, 10, 1)) And shift_end > 3 And shift_start < 3 Then
                result = result + 1
            End If

        End If

        If Under_18 Then
            If result > 4.5 Then
                If result > 6 Then
                    result = result - 1
                Else
                    result = result - 0.5
                End If
            End If
        Else
            If result > 6 Then
                If result > 9.5 Then
                    result = result - 0.75
                Else
                    result = result - 0.5
                End If
            End If

        End If

        Return result
    End Function

    Public Function unter_18(BirthDate As String, on_date As String) As Boolean
        Dim iYear As Integer
        Dim iMonth As Integer
        Dim d As Integer
        Dim dt As Date
        Dim result As Boolean

        If Not IsDate(BirthDate) Or Not IsDate(on_date) Then Exit Function

        dt = CDate(BirthDate)
        If dt > Now Then Exit Function

        iYear = Year(dt)
        iMonth = Month(dt)
        d = DateAndTime.Day(dt)

        iYear = Year(CDate(on_date)) - iYear
        iMonth = Month(CDate(on_date)) - iMonth
        d = DateAndTime.Day(CDate(on_date)) - d

        If Math.Sign(d) = -1 Then
            d = 30 - Math.Abs(d)
            iMonth = iMonth - 1
        End If

        If Math.Sign(iMonth) = -1 Then
            iMonth = 12 - Math.Abs(iMonth)
            iYear = iYear - 1
        End If

        If iYear < 18 Then
            result = True
        Else
            result = False
        End If

        Return result
    End Function

    Public Sub calculate_quarter(employee_id, department_id, senior, BirthDate)

        Dim working_days()
        Dim Parsed As Dictionary(Of String, Object)
        Dim value As Dictionary(Of String, Object)
        Dim objhttp As Object = CreateObject("MSXML2.XMLHTTP")
        Dim shift_type As String = ""

        objhttp.Open("GET", "https://openapi.planday.com/scheduling/v1.0/shifts?employeeId=" & employee_id & "&shiftStatus=Approved" & "&from=" & quarter_start_ & "&to=" & quarter_end_, False)
        objhttp.SetRequestHeader("Authorization", "Bearer " & api_token)
        objhttp.SetRequestHeader("X-ClientId", api_client_id)
        objhttp.Send()
        Parsed = New Web.Script.Serialization.JavaScriptSerializer().Deserialize(Of Object)(objhttp.ResponseText)

        Dim shifts_count As Integer
        If Parsed("paging").ContainsKey("total") Then shifts_count = Parsed("paging")("total")

        For y = 0 To CInt(shifts_count / 50) '
            objhttp.Open("GET", "https://openapi.planday.com/scheduling/v1.0/shifts?employeeId=" & employee_id & "&shiftStatus=Approved" & "&from=" & quarter_start_ & "&to=" & quarter_end_ & "&offset=" & y * 50, False)
            objhttp.SetRequestHeader("Authorization", "Bearer " & api_token)
            objhttp.SetRequestHeader("X-ClientId", api_client_id)
            objhttp.Send()
            Parsed = New Web.Script.Serialization.JavaScriptSerializer().Deserialize(Of Object)(objhttp.ResponseText)

            For Each value In Parsed("data")
                If (value.ContainsKey("startDateTime") And value.ContainsKey("endDateTime")) Then

                    If value.ContainsKey("shiftTypeId") Then
                        shift_type = shift_types(value("shiftTypeId"))
                    Else
                        shift_type = "Normal"
                    End If

                    Select Case shift_type

                        Case "Sick leave - unpaid"
                        Case "No show"
                        Case "FREI"
                        Case "Time off in lieu"
                        Case "Sick leave - paid"

                            If working_days Is Nothing Then
                                ReDim Preserve working_days(0)
                                working_days(0) = CDate(Strings.Left(value("startDateTime"), 10))
                            Else
                                If Not working_days.Contains(CDate(Strings.Left(value("startDateTime"), 10))) Then
                                    ReDim Preserve working_days(UBound(working_days) + 1)
                                    working_days(UBound(working_days)) = CDate(Strings.Left(value("startDateTime"), 10))
                                End If
                            End If

                            quarter_working_hours = quarter_working_hours + shift_length(value("startDateTime"), value("endDateTime"), unter_18(BirthDate, Strings.Left(value("startDateTime"), 10)))

                        Case "Holiday"
                        Case "Sick leave - Bank holiday"

                            If working_days Is Nothing Then
                                ReDim Preserve working_days(0)
                                working_days(0) = CDate(Strings.Left(value("startDateTime"), 10))
                            Else
                                If Not working_days.Contains(CDate(Strings.Left(value("startDateTime"), 10))) Then
                                    ReDim Preserve working_days(UBound(working_days) + 1)
                                    working_days(UBound(working_days)) = CDate(Strings.Left(value("startDateTime"), 10))
                                End If
                            End If

                            quarter_working_hours = quarter_working_hours + shift_length(value("startDateTime"), value("endDateTime"), unter_18(BirthDate, Strings.Left(value("startDateTime"), 10)))

                        Case "Anderes Haus übernimmt Schicht"
                        Case "Bank holiday (off)"
                        Case "Sick leave - child"

                            If working_days Is Nothing Then
                                ReDim Preserve working_days(0)
                                working_days(0) = CDate(Strings.Left(value("startDateTime"), 10))
                            Else
                                If Not working_days.Contains(CDate(Strings.Left(value("startDateTime"), 10))) Then
                                    ReDim Preserve working_days(UBound(working_days) + 1)
                                    working_days(UBound(working_days)) = CDate(Strings.Left(value("startDateTime"), 10))
                                End If
                            End If

                            quarter_working_hours = quarter_working_hours + shift_length(value("startDateTime"), value("endDateTime"), unter_18(BirthDate, Strings.Left(value("startDateTime"), 10)))

                        Case "Unpaid leave"
                        Case "FREI // Kurzarbeit Ausgleich"
                        Case "Parental leave"
                        Case Else

                            If working_days Is Nothing Then
                                ReDim Preserve working_days(0)
                                working_days(0) = CDate(Strings.Left(value("startDateTime"), 10))
                            Else
                                If Not working_days.Contains(CDate(Strings.Left(value("startDateTime"), 10))) Then
                                    ReDim Preserve working_days(UBound(working_days) + 1)
                                    working_days(UBound(working_days)) = CDate(Strings.Left(value("startDateTime"), 10))
                                End If
                            End If

                            quarter_working_hours = quarter_working_hours + shift_length(value("startDateTime"), value("endDateTime"), unter_18(BirthDate, Strings.Left(value("startDateTime"), 10)))

                    End Select
                End If
            Next

        Next

        quarter_working_days = If(working_days Is Nothing, 0, working_days.Length)

    End Sub

    Public Sub calculate_last_3_months(employee_id, department_id, senior, BirthDate)

        Dim breaks()
        Dim Parsed As Dictionary(Of String, Object)
        Dim value As Dictionary(Of String, Object)
        Dim objhttp As Object = CreateObject("MSXML2.XMLHTTP")
        Dim shift_type As String = ""

        objhttp.Open("GET", "https://openapi.planday.com/scheduling/v1.0/shifts?employeeId=" & employee_id & "&shiftStatus=Approved" & "&from=" & last_3_months_start_ & "&to=" & last_3_months_end_, False)
        objhttp.SetRequestHeader("Authorization", "Bearer " & api_token)
        objhttp.SetRequestHeader("X-ClientId", api_client_id)
        objhttp.Send()
        Parsed = New Web.Script.Serialization.JavaScriptSerializer().Deserialize(Of Object)(objhttp.ResponseText)

        Dim shifts_count As Integer
        If Parsed("paging").ContainsKey("total") Then shifts_count = Parsed("paging")("total")

        For y = 0 To CInt(shifts_count / 50)
            objhttp.Open("GET", "https://openapi.planday.com/scheduling/v1.0/shifts?employeeId=" & employee_id & "&shiftStatus=Approved" & "&from=" & last_3_months_start_ & "&to=" & last_3_months_end_ & "&offset=" & y * 50, False)
            objhttp.SetRequestHeader("Authorization", "Bearer " & api_token)
            objhttp.SetRequestHeader("X-ClientId", api_client_id)
            objhttp.Send()
            Parsed = New Web.Script.Serialization.JavaScriptSerializer().Deserialize(Of Object)(objhttp.ResponseText)

            For Each value In Parsed("data")
                If (value.ContainsKey("startDateTime") And value.ContainsKey("endDateTime")) Then

                    Erase breaks

                    breaks = return_breaks(value("startDateTime"), value("endDateTime"), unter_18(BirthDate, Strings.Left(value("startDateTime"), 10)))

                    If value.ContainsKey("shiftTypeId") Then
                        shift_type = shift_types(value("shiftTypeId"))
                    Else
                        shift_type = "Normal"
                    End If

                    Select Case shift_type

                        Case "Sick leave - unpaid"
                        Case "No show"
                        Case "FREI"
                        Case "Time off in lieu"
                        Case "Sick leave - paid"
                        Case "Holiday"
                        Case "Sick leave - Bank holiday"
                        Case "Anderes Haus übernimmt Schicht"
                        Case "Bank holiday (off)"
                        Case "Sick leave - child"
                        Case "Unpaid leave"
                        Case "FREI // Kurzarbeit Ausgleich"
                        Case "Parental leave"
                        Case Else

                            last_3_months_sunday_FT_hours = last_3_months_sunday_FT_hours + calculate_sunday_FT(value("startDateTime"), value("endDateTime"), department_id)
                            If Not breaks Is Nothing Then
                                For Each break In breaks
                                    last_3_months_sunday_FT_hours = last_3_months_sunday_FT_hours - calculate_sunday_FT(Strings.Left(break, 16), Strings.Right(break, 16), department_id)
                                Next
                            End If

                            last_3_months_night_hours = last_3_months_night_hours + calculate_night(value("startDateTime"), value("endDateTime"), department_id, senior)
                            If Not breaks Is Nothing Then
                                For Each break In breaks
                                    last_3_months_night_hours = last_3_months_night_hours - calculate_night(Strings.Left(break, 16), Strings.Right(break, 16), department_id, senior)
                                Next
                            End If

                    End Select
                End If
            Next

        Next
    End Sub

    Public Sub calculate_All_XMAS(employee_id, department_id, senior, BirthDate)

        Dim row_ As Integer
        Dim breaks()
        Dim Parsed As Dictionary(Of String, Object)
        Dim value As Dictionary(Of String, Object)
        Dim objhttp As Object = CreateObject("MSXML2.XMLHTTP")
        Dim temp_sunday As Double
        Dim temp_night As Double
        Dim temp_150 As Double
        Dim temp_125 As Double

        objhttp.open("GET", "https://openapi.planday.com/scheduling/v1.0/shifts?employeeId=" & employee_id & "&shiftStatus=Approved" & "&from=" & start_ & "&to=" & end_, False)
        objhttp.setRequestHeader("Authorization", "Bearer " & api_token)
        objhttp.setRequestHeader("X-ClientId", api_client_id)
        objhttp.send()
        Parsed = New Web.Script.Serialization.JavaScriptSerializer().Deserialize(Of Object)(objhttp.responseText)

        Dim shifts_count As Integer
        If Parsed("paging").ContainsKey("total") Then shifts_count = Parsed("paging")("total")

        For y = 0 To CInt(shifts_count / 50) '
            objhttp.open("GET", "https://openapi.planday.com/scheduling/v1.0/shifts?employeeId=" & employee_id & "&shiftStatus=Approved" & "&from=" & start_ & "&to=" & end_ & "&offset=" & y * 50, False)
            objhttp.setRequestHeader("Authorization", "Bearer " & api_token)
            objhttp.setRequestHeader("X-ClientId", api_client_id)
            objhttp.send()
            Parsed = New Web.Script.Serialization.JavaScriptSerializer().Deserialize(Of Object)(objhttp.responseText)

            For Each value In Parsed("data")
                If (value.ContainsKey("startDateTime") And value.ContainsKey("endDateTime")) Then

                    Erase breaks

                    breaks = return_breaks(value("startDateTime"), value("endDateTime"), unter_18(BirthDate, Strings.Left(value("startDateTime"), 10)))

                    Dim shift_type As String = ""

                    If value.ContainsKey("shiftTypeId") Then
                        shift_type = shift_types(value("shiftTypeId"))
                    Else
                        shift_type = "Normal"
                    End If
                    '''''''''''''''''''''''''''''

                    If Not SheetExists(current_employee, report_workbooks(current_workbook)) Then
                        report_workbooks(current_workbook).Worksheets.Add(current_employee)
                        report_workbooks(current_workbook).Worksheet(current_employee).Range("A1").Value = "date"
                        report_workbooks(current_workbook).Worksheet(current_employee).Range("B1").Value = "shift type"
                        report_workbooks(current_workbook).Worksheet(current_employee).Range("C1").Value = "from"
                        report_workbooks(current_workbook).Worksheet(current_employee).Range("D1").Value = "to"
                        report_workbooks(current_workbook).Worksheet(current_employee).Range("E1").Value = "break1 from"
                        report_workbooks(current_workbook).Worksheet(current_employee).Range("F1").Value = "break1 to"
                        report_workbooks(current_workbook).Worksheet(current_employee).Range("G1").Value = "break2 from"
                        report_workbooks(current_workbook).Worksheet(current_employee).Range("H1").Value = "break2 to"
                        report_workbooks(current_workbook).Worksheet(current_employee).Range("I1").Value = "sunday"
                        report_workbooks(current_workbook).Worksheet(current_employee).Range("J1").Value = "night"
                        report_workbooks(current_workbook).Worksheet(current_employee).Range("K1").Value = "festive_150"
                        report_workbooks(current_workbook).Worksheet(current_employee).Range("L1").Value = "festive_125"
                        report_workbooks(current_workbook).Worksheet(current_employee).Range("M1").Value = "working"
                        report_workbooks(current_workbook).Worksheet(current_employee).Range("N1").Value = "paid_sick_leave"
                        report_workbooks(current_workbook).Worksheet(current_employee).Range("O1").Value = "sick_leave_sunday"
                        report_workbooks(current_workbook).Worksheet(current_employee).Range("P1").Value = "sick_leave_night"

                        row_ = 2
                    End If

                    report_workbooks(current_workbook).Worksheet(current_employee).Range("A" & row_).Value = CDate(value("date"))
                    report_workbooks(current_workbook).Worksheet(current_employee).Range("A" & row_).Style.NumberFormat.Format = "[$-de-DE]ddd dd/mm/yyyy"
                    If Weekday(CDate(value("date")), vbMonday) = 7 Or FT_found(CDate(value("date")), department_id) Then
                        report_workbooks(current_workbook).Worksheet(current_employee).Range("A" & row_).Style.Font.FontColor = XLColor.Red
                    End If
                    report_workbooks(current_workbook).Worksheet(current_employee).Range("B" & row_).Value = shift_type
                    report_workbooks(current_workbook).Worksheet(current_employee).Range("C" & row_).Value = Strings.Right(value("startDateTime"), 5)
                    report_workbooks(current_workbook).Worksheet(current_employee).Range("C" & row_).Style.NumberFormat.Format = "hh:mm"
                    report_workbooks(current_workbook).Worksheet(current_employee).Range("D" & row_).Value = Strings.Right(value("endDateTime"), 5)
                    report_workbooks(current_workbook).Worksheet(current_employee).Range("D" & row_).Style.NumberFormat.Format = "hh:mm"
                    If Not breaks Is Nothing Then
                        report_workbooks(current_workbook).Worksheet(current_employee).Range("E" & row_).Value = Strings.Right(Strings.Left(breaks(0), 16), 5)
                        report_workbooks(current_workbook).Worksheet(current_employee).Range("E" & row_).Style.NumberFormat.Format = "hh:mm"
                        report_workbooks(current_workbook).Worksheet(current_employee).Range("F" & row_).Value = Strings.Right(Strings.Right(breaks(0), 16), 5)
                        report_workbooks(current_workbook).Worksheet(current_employee).Range("F" & row_).Style.NumberFormat.Format = "hh:mm"
                        If breaks.Length > 1 Then
                            report_workbooks(current_workbook).Worksheet(current_employee).Range("G" & row_).Value = Strings.Right(Strings.Left(breaks(1), 16), 5)
                            report_workbooks(current_workbook).Worksheet(current_employee).Range("G" & row_).Style.NumberFormat.Format = "hh:mm"
                            report_workbooks(current_workbook).Worksheet(current_employee).Range("H" & row_).Value = Strings.Right(Strings.Right(breaks(1), 16), 5)
                            report_workbooks(current_workbook).Worksheet(current_employee).Range("H" & row_).Style.NumberFormat.Format = "hh:mm"
                        End If
                    End If
                    ''''''''''''''''''''''''''''''''
                    Select Case shift_type

                        Case "Sick leave - unpaid"
                        Case "No show"
                        Case "FREI"
                        Case "Time off in lieu"
                        Case "Sick leave - paid"

                            sunday_FT_hours_krank = sunday_FT_hours_krank + calculate_sunday_FT(value("startDateTime"), value("endDateTime"), department_id) + calculate_XMAS_150(value("startDateTime"), value("endDateTime"), department_id) + calculate_XMAS_125(value("startDateTime"), value("endDateTime"), department_id)
                            If Not breaks Is Nothing Then
                                For Each break In breaks
                                    ''''
                                    temp_sunday = temp_sunday + calculate_sunday_FT(Strings.Left(break, 16), Strings.Right(break, 16), department_id) + calculate_XMAS_150(Strings.Left(break, 16), Strings.Right(break, 16), department_id) + calculate_XMAS_125(Strings.Left(break, 16), Strings.Right(break, 16), department_id)
                                    ''''
                                    sunday_FT_hours_krank = sunday_FT_hours_krank - calculate_sunday_FT(Strings.Left(break, 16), Strings.Right(break, 16), department_id) - calculate_XMAS_150(Strings.Left(break, 16), Strings.Right(break, 16), department_id) - calculate_XMAS_125(Strings.Left(break, 16), Strings.Right(break, 16), department_id)
                                Next
                            End If

                            night_hours_krank = night_hours_krank + calculate_night(value("startDateTime"), value("endDateTime"), department_id, senior)
                            If Not breaks Is Nothing Then
                                For Each break In breaks
                                    ''''
                                    temp_night = temp_night + calculate_night(Strings.Left(break, 16), Strings.Right(break, 16), department_id, senior)
                                    ''''
                                    night_hours_krank = night_hours_krank - calculate_night(Strings.Left(break, 16), Strings.Right(break, 16), department_id, senior)
                                Next
                            End If

                            working_hours_krank = working_hours_krank + shift_length(value("startDateTime"), value("endDateTime"), unter_18(BirthDate, Strings.Left(value("startDateTime"), 10)))
                            '''''''''''
                            report_workbooks(current_workbook).Worksheet(current_employee).Range("N" & row_).Value = shift_length(value("startDateTime"), value("endDateTime"), unter_18(BirthDate, Strings.Left(value("startDateTime"), 10)))
                            report_workbooks(current_workbook).Worksheet(current_employee).Range("O" & row_).Value = calculate_sunday_FT(value("startDateTime"), value("endDateTime"), department_id) + calculate_XMAS_150(value("startDateTime"), value("endDateTime"), department_id) + calculate_XMAS_125(value("startDateTime"), value("endDateTime"), department_id) - temp_sunday
                            report_workbooks(current_workbook).Worksheet(current_employee).Range("P" & row_).Value = calculate_night(value("startDateTime"), value("endDateTime"), department_id, senior) - temp_night
                            temp_sunday = 0
                            temp_night = 0
                            temp_150 = 0
                            temp_125 = 0
                    '''''''''''
                        Case "Holiday"
                        Case "Sick leave - Bank holiday"

                            sunday_FT_hours_krank = sunday_FT_hours_krank + calculate_sunday_FT(value("startDateTime"), value("endDateTime"), department_id) + calculate_XMAS_150(value("startDateTime"), value("endDateTime"), department_id) + calculate_XMAS_125(value("startDateTime"), value("endDateTime"), department_id)
                            If Not breaks Is Nothing Then
                                For Each break In breaks
                                    ''''
                                    temp_sunday = temp_sunday + calculate_sunday_FT(Strings.Left(break, 16), Strings.Right(break, 16), department_id) + calculate_XMAS_150(Strings.Left(break, 16), Strings.Right(break, 16), department_id) + calculate_XMAS_125(Strings.Left(break, 16), Strings.Right(break, 16), department_id)
                                    ''''
                                    sunday_FT_hours_krank = sunday_FT_hours_krank - calculate_sunday_FT(Strings.Left(break, 16), Strings.Right(break, 16), department_id) - calculate_XMAS_150(Strings.Left(break, 16), Strings.Right(break, 16), department_id) - calculate_XMAS_125(Strings.Left(break, 16), Strings.Right(break, 16), department_id)
                                Next
                            End If

                            night_hours_krank = night_hours_krank + calculate_night(value("startDateTime"), value("endDateTime"), department_id, senior)
                            If Not breaks Is Nothing Then
                                For Each break In breaks
                                    ''''
                                    temp_night = temp_night + calculate_night(Strings.Left(break, 16), Strings.Right(break, 16), department_id, senior)
                                    ''''
                                    night_hours_krank = night_hours_krank - calculate_night(Strings.Left(break, 16), Strings.Right(break, 16), department_id, senior)
                                Next
                            End If

                            working_hours_krank = working_hours_krank + shift_length(value("startDateTime"), value("endDateTime"), unter_18(BirthDate, Strings.Left(value("startDateTime"), 10)))
                            '''''''''''
                            report_workbooks(current_workbook).Worksheet(current_employee).Range("N" & row_).Value = shift_length(value("startDateTime"), value("endDateTime"), unter_18(BirthDate, Strings.Left(value("startDateTime"), 10)))
                            report_workbooks(current_workbook).Worksheet(current_employee).Range("O" & row_).Value = calculate_sunday_FT(value("startDateTime"), value("endDateTime"), department_id) + calculate_XMAS_150(value("startDateTime"), value("endDateTime"), department_id) + calculate_XMAS_125(value("startDateTime"), value("endDateTime"), department_id) - temp_sunday
                            report_workbooks(current_workbook).Worksheet(current_employee).Range("P" & row_).Value = calculate_night(value("startDateTime"), value("endDateTime"), department_id, senior) - temp_night
                            temp_sunday = 0
                            temp_night = 0
                            temp_150 = 0
                            temp_125 = 0
                    '''''''''''
                        Case "Anderes Haus übernimmt Schicht"
                        Case "Bank holiday (off)"
                        Case "Sick leave - child"

                            sunday_FT_hours_krank = sunday_FT_hours_krank + calculate_sunday_FT(value("startDateTime"), value("endDateTime"), department_id) + calculate_XMAS_150(value("startDateTime"), value("endDateTime"), department_id) + calculate_XMAS_125(value("startDateTime"), value("endDateTime"), department_id)
                            If Not breaks Is Nothing Then
                                For Each break In breaks
                                    ''''
                                    temp_sunday = temp_sunday + calculate_sunday_FT(Strings.Left(break, 16), Strings.Right(break, 16), department_id) + calculate_XMAS_150(Strings.Left(break, 16), Strings.Right(break, 16), department_id) + calculate_XMAS_125(Strings.Left(break, 16), Strings.Right(break, 16), department_id)
                                    ''''
                                    sunday_FT_hours_krank = sunday_FT_hours_krank - calculate_sunday_FT(Strings.Left(break, 16), Strings.Right(break, 16), department_id) - calculate_XMAS_150(Strings.Left(break, 16), Strings.Right(break, 16), department_id) - calculate_XMAS_125(Strings.Left(break, 16), Strings.Right(break, 16), department_id)
                                Next
                            End If

                            night_hours_krank = night_hours_krank + calculate_night(value("startDateTime"), value("endDateTime"), department_id, senior)
                            If Not breaks Is Nothing Then
                                For Each break In breaks
                                    ''''
                                    temp_night = temp_night + calculate_night(Strings.Left(break, 16), Strings.Right(break, 16), department_id, senior)
                                    ''''
                                    night_hours_krank = night_hours_krank - calculate_night(Strings.Left(break, 16), Strings.Right(break, 16), department_id, senior)
                                Next
                            End If

                            working_hours_krank = working_hours_krank + shift_length(value("startDateTime"), value("endDateTime"), unter_18(BirthDate, Strings.Left(value("startDateTime"), 10)))
                            '''''''''''
                            report_workbooks(current_workbook).Worksheet(current_employee).Range("N" & row_).Value = shift_length(value("startDateTime"), value("endDateTime"), unter_18(BirthDate, Strings.Left(value("startDateTime"), 10)))
                            report_workbooks(current_workbook).Worksheet(current_employee).Range("O" & row_).Value = calculate_sunday_FT(value("startDateTime"), value("endDateTime"), department_id) + calculate_XMAS_150(value("startDateTime"), value("endDateTime"), department_id) + calculate_XMAS_125(value("startDateTime"), value("endDateTime"), department_id) - temp_sunday
                            report_workbooks(current_workbook).Worksheet(current_employee).Range("P" & row_).Value = calculate_night(value("startDateTime"), value("endDateTime"), department_id, senior) - temp_night
                            temp_sunday = 0
                            temp_night = 0
                            temp_150 = 0
                            temp_125 = 0
                    '''''''''''
                        Case "Unpaid leave"
                        Case "FREI // Kurzarbeit Ausgleich"
                        Case "Parental leave"
                        Case Else

                            sunday_FT_hours = sunday_FT_hours + calculate_sunday_FT(value("startDateTime"), value("endDateTime"), department_id)
                            XMAS_150 = XMAS_150 + calculate_XMAS_150(value("startDateTime"), value("endDateTime"), department_id)
                            XMAS_125 = XMAS_125 + calculate_XMAS_125(value("startDateTime"), value("endDateTime"), department_id)

                            If Not breaks Is Nothing Then
                                For Each break In breaks
                                    ''''
                                    temp_sunday = temp_sunday + calculate_sunday_FT(Strings.Left(break, 16), Strings.Right(break, 16), department_id)
                                    temp_150 = temp_150 + calculate_XMAS_150(Strings.Left(break, 16), Strings.Right(break, 16), department_id)
                                    temp_125 = temp_125 + calculate_XMAS_125(Strings.Left(break, 16), Strings.Right(break, 16), department_id)
                                    ''''
                                    sunday_FT_hours = sunday_FT_hours - calculate_sunday_FT(Strings.Left(break, 16), Strings.Right(break, 16), department_id)
                                    XMAS_150 = XMAS_150 - calculate_XMAS_150(Strings.Left(break, 16), Strings.Right(break, 16), department_id)
                                    XMAS_125 = XMAS_125 - calculate_XMAS_125(Strings.Left(break, 16), Strings.Right(break, 16), department_id)
                                Next
                            End If

                            night_hours = night_hours + calculate_night(value("startDateTime"), value("endDateTime"), department_id, senior)
                            If Not breaks Is Nothing Then
                                For Each break In breaks
                                    ''''
                                    temp_night = temp_night + calculate_night(Strings.Left(break, 16), Strings.Right(break, 16), department_id, senior)
                                    ''''
                                    night_hours = night_hours - calculate_night(Strings.Left(break, 16), Strings.Right(break, 16), department_id, senior)
                                Next
                            End If

                            working_hours = working_hours + shift_length(value("startDateTime"), value("endDateTime"), unter_18(BirthDate, Strings.Left(value("startDateTime"), 10)))
                            '''''''''''
                            report_workbooks(current_workbook).Worksheet(current_employee).Range("M" & row_).Value = shift_length(value("startDateTime"), value("endDateTime"), unter_18(BirthDate, Strings.Left(value("startDateTime"), 10)))
                            report_workbooks(current_workbook).Worksheet(current_employee).Range("I" & row_).Value = calculate_sunday_FT(value("startDateTime"), value("endDateTime"), department_id) - temp_sunday
                            report_workbooks(current_workbook).Worksheet(current_employee).Range("J" & row_).Value = calculate_night(value("startDateTime"), value("endDateTime"), department_id, senior) - temp_night
                            report_workbooks(current_workbook).Worksheet(current_employee).Range("K" & row_).Value = calculate_XMAS_150(value("startDateTime"), value("endDateTime"), department_id) - temp_150
                            report_workbooks(current_workbook).Worksheet(current_employee).Range("L" & row_).Value = calculate_XMAS_125(value("startDateTime"), value("endDateTime"), department_id) - temp_125
                            temp_sunday = 0
                            temp_night = 0
                            temp_150 = 0
                            temp_125 = 0
                            '''''''''''
                    End Select
                    ''''
                    row_ = row_ + 1
                    ''''
                End If
            Next
        Next
        '''''
        Dim i As Integer
        If row_ > 0 Then
            Dim table As IXLTable = report_workbooks(current_workbook).Worksheet(current_employee).Range("$A$1:$P$" & row_ - 1).CreateTable(current_employee)
            table.Theme = XLTableTheme.TableStyleLight8
            table.ShowTotalsRow = True
            table.Field(0).TotalsRowLabel = "Total"
            report_workbooks(current_workbook).Worksheet(current_employee).Range("A" & row_).Style.Font.FontColor = XLColor.Black
            For i = 8 To table.ColumnCount - 1
                table.Field(i).TotalsRowFunction = XLTotalsRowFunction.Sum
            Next

            report_workbooks(current_workbook).Worksheet(current_employee).Columns("A:A").Width = 12.64
            Dim max_length As Integer = 0
            For Each cell In table.DataRange.Column(2).Cells
                If Len(cell.GetFormattedString) > max_length Then max_length = Len(cell.GetFormattedString)
            Next
            report_workbooks(current_workbook).Worksheet(current_employee).Columns("B:B").Width = CalculateColumnWidth(max_length)
            report_workbooks(current_workbook).Worksheet(current_employee).Columns("C:D").Width = 6.55
            report_workbooks(current_workbook).Worksheet(current_employee).Columns("E:H").Width = 12.82
            report_workbooks(current_workbook).Worksheet(current_employee).Columns("I:M").Width = 9.18
            report_workbooks(current_workbook).Worksheet(current_employee).Columns("N:P").Width = 16.91
        End If
        ''''''
    End Sub
    Public Sub calculate_All(employee_id, department_id, senior, BirthDate)

        Dim row_ As Integer
        Dim breaks()
        Dim Parsed As Dictionary(Of String, Object)
        Dim value As Dictionary(Of String, Object)
        Dim objhttp As Object = CreateObject("MSXML2.XMLHTTP")
        Dim temp_sunday As Double
        Dim temp_night As Double

        objhttp.open("GET", "https://openapi.planday.com/scheduling/v1.0/shifts?employeeId=" & employee_id & "&shiftStatus=Approved" & "&from=" & start_ & "&to=" & end_, False)
        objhttp.setRequestHeader("Authorization", "Bearer " & api_token)
        objhttp.setRequestHeader("X-ClientId", api_client_id)
        objhttp.send()
        Parsed = New Web.Script.Serialization.JavaScriptSerializer().Deserialize(Of Object)(objhttp.responseText)

        Dim shifts_count As Integer
        If Parsed("paging").ContainsKey("total") Then shifts_count = Parsed("paging")("total")

        For y = 0 To CInt(shifts_count / 50) '
            objhttp.open("GET", "https://openapi.planday.com/scheduling/v1.0/shifts?employeeId=" & employee_id & "&shiftStatus=Approved" & "&from=" & start_ & "&to=" & end_ & "&offset=" & y * 50, False)
            objhttp.setRequestHeader("Authorization", "Bearer " & api_token)
            objhttp.setRequestHeader("X-ClientId", api_client_id)
            objhttp.send()
            Parsed = New Web.Script.Serialization.JavaScriptSerializer().Deserialize(Of Object)(objhttp.responseText)

            For Each value In Parsed("data")
                If (value.ContainsKey("startDateTime") And value.ContainsKey("endDateTime")) Then

                    Erase breaks

                    breaks = return_breaks(value("startDateTime"), value("endDateTime"), unter_18(BirthDate, Strings.Left(value("startDateTime"), 10)))

                    Dim shift_type As String = ""

                    If value.ContainsKey("shiftTypeId") Then
                        shift_type = shift_types(value("shiftTypeId"))
                    Else
                        shift_type = "Normal"
                    End If
                    '''''''''''''''''''''''''''''

                    If Not SheetExists(current_employee, report_workbooks(current_workbook)) Then
                        report_workbooks(current_workbook).Worksheets.Add(current_employee)
                        report_workbooks(current_workbook).Worksheet(current_employee).Range("A1").Value = "date"
                        report_workbooks(current_workbook).Worksheet(current_employee).Range("B1").Value = "shift type"
                        report_workbooks(current_workbook).Worksheet(current_employee).Range("C1").Value = "from"
                        report_workbooks(current_workbook).Worksheet(current_employee).Range("D1").Value = "to"
                        report_workbooks(current_workbook).Worksheet(current_employee).Range("E1").Value = "break1 from"
                        report_workbooks(current_workbook).Worksheet(current_employee).Range("F1").Value = "break1 to"
                        report_workbooks(current_workbook).Worksheet(current_employee).Range("G1").Value = "break2 from"
                        report_workbooks(current_workbook).Worksheet(current_employee).Range("H1").Value = "break2 to"
                        report_workbooks(current_workbook).Worksheet(current_employee).Range("I1").Value = "sunday"
                        report_workbooks(current_workbook).Worksheet(current_employee).Range("J1").Value = "night"
                        report_workbooks(current_workbook).Worksheet(current_employee).Range("K1").Value = "working"
                        report_workbooks(current_workbook).Worksheet(current_employee).Range("L1").Value = "paid_sick_leave"
                        report_workbooks(current_workbook).Worksheet(current_employee).Range("M1").Value = "sick_leave_sunday"
                        report_workbooks(current_workbook).Worksheet(current_employee).Range("N1").Value = "sick_leave_night"

                        row_ = 2
                    End If

                    report_workbooks(current_workbook).Worksheet(current_employee).Range("A" & row_).Value = CDate(value("date"))
                    report_workbooks(current_workbook).Worksheet(current_employee).Range("A" & row_).Style.NumberFormat.Format = "[$-de-DE]ddd dd/mm/yyyy"
                    If Weekday(CDate(value("date")), vbMonday) = 7 Or FT_found(CDate(value("date")), department_id) Then
                        report_workbooks(current_workbook).Worksheet(current_employee).Range("A" & row_).Style.Font.FontColor = XLColor.Red
                    End If
                    report_workbooks(current_workbook).Worksheet(current_employee).Range("B" & row_).Value = shift_type
                    report_workbooks(current_workbook).Worksheet(current_employee).Range("C" & row_).Value = Strings.Right(value("startDateTime"), 5)
                    report_workbooks(current_workbook).Worksheet(current_employee).Range("C" & row_).Style.NumberFormat.Format = "hh:mm"
                    report_workbooks(current_workbook).Worksheet(current_employee).Range("D" & row_).Value = Strings.Right(value("endDateTime"), 5)
                    report_workbooks(current_workbook).Worksheet(current_employee).Range("D" & row_).Style.NumberFormat.Format = "hh:mm"
                    If Not breaks Is Nothing Then
                        report_workbooks(current_workbook).Worksheet(current_employee).Range("E" & row_).Value = Strings.Right(Strings.Left(breaks(0), 16), 5)
                        report_workbooks(current_workbook).Worksheet(current_employee).Range("E" & row_).Style.NumberFormat.Format = "hh:mm"
                        report_workbooks(current_workbook).Worksheet(current_employee).Range("F" & row_).Value = Strings.Right(Strings.Right(breaks(0), 16), 5)
                        report_workbooks(current_workbook).Worksheet(current_employee).Range("F" & row_).Style.NumberFormat.Format = "hh:mm"
                        If breaks.Length > 1 Then
                            report_workbooks(current_workbook).Worksheet(current_employee).Range("G" & row_).Value = Strings.Right(Strings.Left(breaks(1), 16), 5)
                            report_workbooks(current_workbook).Worksheet(current_employee).Range("G" & row_).Style.NumberFormat.Format = "hh:mm"
                            report_workbooks(current_workbook).Worksheet(current_employee).Range("H" & row_).Value = Strings.Right(Strings.Right(breaks(1), 16), 5)
                            report_workbooks(current_workbook).Worksheet(current_employee).Range("H" & row_).Style.NumberFormat.Format = "hh:mm"
                        End If
                    End If
                    ''''''''''''''''''''''''''''''''
                    Select Case shift_type

                        Case "Sick leave - unpaid"
                        Case "No show"
                        Case "FREI"
                        Case "Time off in lieu"
                        Case "Sick leave - paid"

                            sunday_FT_hours_krank = sunday_FT_hours_krank + calculate_sunday_FT(value("startDateTime"), value("endDateTime"), department_id)
                            If Not breaks Is Nothing Then
                                For Each break In breaks
                                    ''''
                                    temp_sunday = temp_sunday + calculate_sunday_FT(Strings.Left(break, 16), Strings.Right(break, 16), department_id)
                                    ''''
                                    sunday_FT_hours_krank = sunday_FT_hours_krank - calculate_sunday_FT(Strings.Left(break, 16), Strings.Right(break, 16), department_id)
                                Next
                            End If

                            night_hours_krank = night_hours_krank + calculate_night(value("startDateTime"), value("endDateTime"), department_id, senior)
                            If Not breaks Is Nothing Then
                                For Each break In breaks
                                    ''''
                                    temp_night = temp_night + calculate_night(Strings.Left(break, 16), Strings.Right(break, 16), department_id, senior)
                                    ''''
                                    night_hours_krank = night_hours_krank - calculate_night(Strings.Left(break, 16), Strings.Right(break, 16), department_id, senior)
                                Next
                            End If

                            working_hours_krank = working_hours_krank + shift_length(value("startDateTime"), value("endDateTime"), unter_18(BirthDate, Strings.Left(value("startDateTime"), 10)))
                            '''''''''''
                            report_workbooks(current_workbook).Worksheet(current_employee).Range("L" & row_).Value = shift_length(value("startDateTime"), value("endDateTime"), unter_18(BirthDate, Strings.Left(value("startDateTime"), 10)))
                            report_workbooks(current_workbook).Worksheet(current_employee).Range("M" & row_).Value = calculate_sunday_FT(value("startDateTime"), value("endDateTime"), department_id) - temp_sunday
                            report_workbooks(current_workbook).Worksheet(current_employee).Range("N" & row_).Value = calculate_night(value("startDateTime"), value("endDateTime"), department_id, senior) - temp_night
                            temp_sunday = 0
                            temp_night = 0
                    '''''''''''
                        Case "Holiday"
                        Case "Sick leave - Bank holiday"

                            sunday_FT_hours_krank = sunday_FT_hours_krank + calculate_sunday_FT(value("startDateTime"), value("endDateTime"), department_id)
                            If Not breaks Is Nothing Then
                                For Each break In breaks
                                    ''''
                                    temp_sunday = temp_sunday + calculate_sunday_FT(Strings.Left(break, 16), Strings.Right(break, 16), department_id)
                                    ''''
                                    sunday_FT_hours_krank = sunday_FT_hours_krank - calculate_sunday_FT(Strings.Left(break, 16), Strings.Right(break, 16), department_id)
                                Next
                            End If

                            night_hours_krank = night_hours_krank + calculate_night(value("startDateTime"), value("endDateTime"), department_id, senior)
                            If Not breaks Is Nothing Then
                                For Each break In breaks
                                    ''''
                                    temp_night = temp_night + calculate_night(Strings.Left(break, 16), Strings.Right(break, 16), department_id, senior)
                                    ''''
                                    night_hours_krank = night_hours_krank - calculate_night(Strings.Left(break, 16), Strings.Right(break, 16), department_id, senior)
                                Next
                            End If

                            working_hours_krank = working_hours_krank + shift_length(value("startDateTime"), value("endDateTime"), unter_18(BirthDate, Strings.Left(value("startDateTime"), 10)))
                            '''''''''''
                            report_workbooks(current_workbook).Worksheet(current_employee).Range("L" & row_).Value = shift_length(value("startDateTime"), value("endDateTime"), unter_18(BirthDate, Strings.Left(value("startDateTime"), 10)))
                            report_workbooks(current_workbook).Worksheet(current_employee).Range("M" & row_).Value = calculate_sunday_FT(value("startDateTime"), value("endDateTime"), department_id) - temp_sunday
                            report_workbooks(current_workbook).Worksheet(current_employee).Range("N" & row_).Value = calculate_night(value("startDateTime"), value("endDateTime"), department_id, senior) - temp_night
                            temp_sunday = 0
                            temp_night = 0
                    '''''''''''
                        Case "Anderes Haus übernimmt Schicht"
                        Case "Bank holiday (off)"
                        Case "Sick leave - child"

                            sunday_FT_hours_krank = sunday_FT_hours_krank + calculate_sunday_FT(value("startDateTime"), value("endDateTime"), department_id)
                            If Not breaks Is Nothing Then
                                For Each break In breaks
                                    ''''
                                    temp_sunday = temp_sunday + calculate_sunday_FT(Strings.Left(break, 16), Strings.Right(break, 16), department_id)
                                    ''''
                                    sunday_FT_hours_krank = sunday_FT_hours_krank - calculate_sunday_FT(Strings.Left(break, 16), Strings.Right(break, 16), department_id)
                                Next
                            End If

                            night_hours_krank = night_hours_krank + calculate_night(value("startDateTime"), value("endDateTime"), department_id, senior)
                            If Not breaks Is Nothing Then
                                For Each break In breaks
                                    ''''
                                    temp_night = temp_night + calculate_night(Strings.Left(break, 16), Strings.Right(break, 16), department_id, senior)
                                    ''''
                                    night_hours_krank = night_hours_krank - calculate_night(Strings.Left(break, 16), Strings.Right(break, 16), department_id, senior)
                                Next
                            End If

                            working_hours_krank = working_hours_krank + shift_length(value("startDateTime"), value("endDateTime"), unter_18(BirthDate, Strings.Left(value("startDateTime"), 10)))
                            '''''''''''
                            report_workbooks(current_workbook).Worksheet(current_employee).Range("L" & row_).Value = shift_length(value("startDateTime"), value("endDateTime"), unter_18(BirthDate, Strings.Left(value("startDateTime"), 10)))
                            report_workbooks(current_workbook).Worksheet(current_employee).Range("M" & row_).Value = calculate_sunday_FT(value("startDateTime"), value("endDateTime"), department_id) - temp_sunday
                            report_workbooks(current_workbook).Worksheet(current_employee).Range("N" & row_).Value = calculate_night(value("startDateTime"), value("endDateTime"), department_id, senior) - temp_night
                            temp_sunday = 0
                            temp_night = 0
                    '''''''''''
                        Case "Unpaid leave"
                        Case "FREI // Kurzarbeit Ausgleich"
                        Case "Parental leave"
                        Case Else

                            sunday_FT_hours = sunday_FT_hours + calculate_sunday_FT(value("startDateTime"), value("endDateTime"), department_id)
                            If Not breaks Is Nothing Then
                                For Each break In breaks
                                    ''''
                                    temp_sunday = temp_sunday + calculate_sunday_FT(Strings.Left(break, 16), Strings.Right(break, 16), department_id)
                                    ''''
                                    sunday_FT_hours = sunday_FT_hours - calculate_sunday_FT(Strings.Left(break, 16), Strings.Right(break, 16), department_id)
                                Next
                            End If

                            night_hours = night_hours + calculate_night(value("startDateTime"), value("endDateTime"), department_id, senior)
                            If Not breaks Is Nothing Then
                                For Each break In breaks
                                    ''''
                                    temp_night = temp_night + calculate_night(Strings.Left(break, 16), Strings.Right(break, 16), department_id, senior)
                                    ''''
                                    night_hours = night_hours - calculate_night(Strings.Left(break, 16), Strings.Right(break, 16), department_id, senior)
                                Next
                            End If

                            working_hours = working_hours + shift_length(value("startDateTime"), value("endDateTime"), unter_18(BirthDate, Strings.Left(value("startDateTime"), 10)))
                            '''''''''''
                            report_workbooks(current_workbook).Worksheet(current_employee).Range("K" & row_).Value = shift_length(value("startDateTime"), value("endDateTime"), unter_18(BirthDate, Strings.Left(value("startDateTime"), 10)))
                            report_workbooks(current_workbook).Worksheet(current_employee).Range("I" & row_).Value = calculate_sunday_FT(value("startDateTime"), value("endDateTime"), department_id) - temp_sunday
                            report_workbooks(current_workbook).Worksheet(current_employee).Range("J" & row_).Value = calculate_night(value("startDateTime"), value("endDateTime"), department_id, senior) - temp_night
                            temp_sunday = 0
                            temp_night = 0
                            '''''''''''
                    End Select
                    ''''
                    row_ = row_ + 1
                    ''''
                End If
            Next
        Next
        '''''
        Dim i As Integer
        If row_ > 0 Then
            Dim table As IXLTable = report_workbooks(current_workbook).Worksheet(current_employee).Range("$A$1:$N$" & row_ - 1).CreateTable(current_employee)
            table.Theme = XLTableTheme.TableStyleLight8
            table.ShowTotalsRow = True
            table.Field(0).TotalsRowLabel = "Total"
            report_workbooks(current_workbook).Worksheet(current_employee).Range("A" & row_).Style.Font.FontColor = XLColor.Black
            For i = 8 To table.ColumnCount - 1
                table.Field(i).TotalsRowFunction = XLTotalsRowFunction.Sum
            Next

            report_workbooks(current_workbook).Worksheet(current_employee).Columns("A:A").Width = 12.64
            Dim max_length As Integer = 0
            For Each cell In table.DataRange.Column(2).Cells
                If Len(cell.GetFormattedString) > max_length Then max_length = Len(cell.GetFormattedString)
            Next
            report_workbooks(current_workbook).Worksheet(current_employee).Columns("B:B").Width = CalculateColumnWidth(max_length)
            report_workbooks(current_workbook).Worksheet(current_employee).Columns("C:D").Width = 6.55
            report_workbooks(current_workbook).Worksheet(current_employee).Columns("E:H").Width = 12.82
            report_workbooks(current_workbook).Worksheet(current_employee).Columns("I:K").Width = 9.18
            report_workbooks(current_workbook).Worksheet(current_employee).Columns("L:N").Width = 16.91
        End If
        ''''''
    End Sub

    Public Sub format_white_borders(range As IXLCells)
        For Each cell In range
            With cell
                .Style.Border.LeftBorder = XLBorderStyleValues.Thin
                .Style.Border.LeftBorderColor = XLColor.White
                .Style.Border.RightBorder = XLBorderStyleValues.Thin
                .Style.Border.RightBorderColor = XLColor.White
            End With
        Next
    End Sub

    Public Sub format_grey(range As IXLRange)
        range.Style.Fill.BackgroundColor = XLColor.FromArgb(191, 191, 191)
    End Sub

    Public Sub format_yellow(range As IXLRange)
        range.Style.Fill.BackgroundColor = XLColor.Yellow
        range.Style.Border.LeftBorder = XLBorderStyleValues.Thin
        range.Style.Border.RightBorder = XLBorderStyleValues.Thin
    End Sub

    Public Sub format_blue(range As IXLRange)
        range.Style.Fill.BackgroundColor = XLColor.FromArgb(189, 215, 238)
        range.Style.Border.LeftBorder = XLBorderStyleValues.Thin
        range.Style.Border.RightBorder = XLBorderStyleValues.Thin
    End Sub

    Public Sub format_blue_noBorders(range As IXLRange)
        range.Style.Fill.BackgroundColor = XLColor.FromArgb(189, 215, 238)
    End Sub

    Public Function LastDayOfMonth(Which_Day As String, Which_Date As String) As Date

        Dim i As Integer

        Dim iDay As Integer

        Dim iDaysInMonth As Integer

        Dim FullDateNew As Date

        Dim result As Date

        Which_Date = CDate(Which_Date)



        Select Case UCase(Which_Day)

            Case "MON"

                iDay = 1

            Case "TUE"

                iDay = 2

            Case "WED"

                iDay = 3

            Case "THU"

                iDay = 4

            Case "FRI"

                iDay = 5

            Case "SAT"

                iDay = 6

            Case "SUN"

                iDay = 7

        End Select



        iDaysInMonth = DateAndTime.Day(DateAdd("d", -1, DateSerial(Year(Which_Date), Month(Which_Date) + 1, 1)))



        FullDateNew = DateSerial(Year(Which_Date), Month(Which_Date), iDaysInMonth)



        For i = 0 To iDaysInMonth

            If Weekday(FullDateNew.AddDays(-i), vbMonday) = iDay Then

                result = FullDateNew.AddDays(-i)

                Exit For

            End If

        Next i
        Return result
    End Function

    Public Function SheetExists(SheetName As String, Optional wb As XLWorkbook = Nothing) As Boolean

        Dim result As Boolean
        If wb Is Nothing Then Return False
        For Each sheet In wb.Worksheets
            If sheet.Name = SheetName Then
                result = True
                Exit For
            Else
                result = False
            End If
        Next
        Return result
    End Function

    Private Sub Fill_LV()
        ImChangingStuff = True
        Me.select_employees.ListViewItemSorter = Nothing
        CheckBox1.Checked = True
        CheckBox1.Font = New Font(CheckBox1.Font, FontStyle.Bold)

        If Not (employees_records Is Nothing) Then
            For Each element In employees_records
                Dim item As ListViewItem = Me.select_employees.Items.Add(element.Mitarbeiter)
                item.Checked = True
                item.Font = New Font(item.Font, FontStyle.Bold)
                item.UseItemStyleForSubItems = False

                item.SubItems.Add(element.Soll)
                item.SubItems.Add(element.Ist).ForeColor = Color.Blue

                item.SubItems.Add(element.Überstunden).Font = New Font(item.Font, FontStyle.Bold)
                If element.Überstunden = 0 Then item.SubItems(3).Text = ""

                If element.zuÜbertragen = 0 Then
                    item.SubItems.Add("")
                    item.SubItems(3).Font = New Font(item.Font, FontStyle.Regular)
                    item.Checked = False
                    item.Font = New Font(item.Font, FontStyle.Regular)
                    CheckBox1.Checked = False
                    CheckBox1.Font = New Font(item.Font, FontStyle.Regular)
                End If

                If element.zuÜbertragen > 0 Then
                    item.SubItems.Add(element.zuÜbertragen).Font = New Font(item.Font, FontStyle.Bold)
                    item.SubItems(4).ForeColor = Color.FromArgb(0, 140, 60)
                End If

                If element.zuÜbertragen < 0 Then
                    item.SubItems.Add(element.zuÜbertragen).Font = New Font(item.Font, FontStyle.Bold)
                    item.SubItems(4).ForeColor = Color.Red
                End If

                item.SubItems.Add(element.Übertragen)
                If element.Übertragen = 0 Then item.SubItems(5).Text = ""

                item.SubItems.Add(element.vonPlanday)
                If element.vonPlanday = 0 Then item.SubItems(6).Text = ""

                item.SubItems.Add(element.Kontostand)
            Next
        End If

        get_FT()
        Label6.Visible = True
        loaded.Text = "KW" & fGetKW(CDate(Strings.Right(select_kw.SelectedItem, 10)))

        Dim lvcs As New ListViewColumnSorter
        lvcs.SortingOrder = SortOrder.Descending
        lvcs.ColumnIndex = 4
        Me.select_employees.Columns(4).Tag = SortOrder.Descending
        Me.select_employees.ListViewItemSorter = lvcs

        ImChangingStuff = False
    End Sub

    Private Function get_FT()
        Dim FT As New Dictionary(Of Date, String)
        Dim department_id As String
        Dim start_kw As Date
        Dim end_kw As Date
        Dim DateToolTip = New ToolTip()

        FT_label.Visible = False
        FT_1.Text = ""
        DateToolTip.SetToolTip(FT_1, "")
        FT_2.Text = ""
        DateToolTip.SetToolTip(FT_2, "")

        department_id = hotel_id

        start_kw = CDate(Strings.Left(Strings.Right(select_kw.SelectedItem, 21), 10))
        end_kw = CDate(Strings.Right(select_kw.SelectedItem, 10))


        Do While start_kw <= end_kw
            If departments_FT(hotel_id).Contains(start_kw) Then FT(Strings.Format(start_kw, "dd.MM.yyyy")) = FT_Name(start_kw)
            start_kw = start_kw.AddDays(1)
        Loop

        If FT.Count = 1 Then
            FT_1.Text = FT.Keys(0)
            DateToolTip.SetToolTip(FT_1, FT.Values.ElementAt(0))
        End If

        If FT.Count = 2 Then
            FT_1.Text = FT.Keys(0)
            DateToolTip.SetToolTip(FT_1, FT.Values.ElementAt(0))
            FT_2.Text = FT.Keys(1)
            DateToolTip.SetToolTip(FT_2, FT.Values.ElementAt(1))
        End If
        If FT.Count > 0 Then FT_label.Visible = True

        Return FT
    End Function

    Public Function employeeID_fromRecords(name As String) As String
        For Each record In employees_records
            If name = record.Mitarbeiter Then
                Return record.employee_id
                Exit Function
            End If
        Next
        Return ""
    End Function

    Private Sub select_hotel_SelectedIndexChanged(sender As Object, e As EventArgs) Handles select_hotel.SelectedIndexChanged
        hotel_id = getDepartmentID(select_hotel.SelectedItem)
    End Sub

    Private Sub select_kw_SelectedIndexChanged(sender As Object, e As EventArgs) Handles select_kw.SelectedIndexChanged
        start_monday = CDate(Strings.Left(Strings.Right(select_kw.SelectedItem, 21), 10))
        current_year = Year(start_monday)
    End Sub

    Private Sub bgw_write_overtime_DoWork(sender As Object, e As DoWorkEventArgs) Handles bgw_write_overtime.DoWork
        Dim employee_id As String

        For i = 0 To select_employees.Items.Count - 1
            If GetItemChecked(i) Then
                employee_id = employeeID_fromRecords(GetItem(i).Text)

                If Boolean_write Then
                    If GetItem(i).SubItems(4).Text <> "" Then
                        adjust_overtime(employee_id, GetItem(i).SubItems(4).Text, "Overtime " & Format(start_monday, "dd.MM.yyyy") & "-" & Format(start_monday.AddDays(6), "dd.MM.yyyy"))
                        SetSubitemText(i, 5, If(GetItem(i).SubItems(5).Text = "", 0, CDbl(GetItem(i).SubItems(5).Text)) + If(GetItem(i).SubItems(4).Text = "", 0, CDbl(GetItem(i).SubItems(4).Text)))
                        SetSubitemText(i, 7, If(GetItem(i).SubItems(7).Text = "", 0, CDbl(GetItem(i).SubItems(7).Text)) + If(GetItem(i).SubItems(4).Text = "", 0, CDbl(GetItem(i).SubItems(4).Text)))
                        SetSubitemText(i, 4, "")
                    End If
                Else
                    If GetItem(i).SubItems(5).Text <> "" Then
                        adjust_overtime(employee_id, 0 - CDbl(GetItem(i).SubItems(5).Text), "Cancellation " & Format(start_monday, "dd.MM.yyyy") & "-" & Format(start_monday.AddDays(6), "dd.MM.yyyy"))
                        SetSubitemText(i, 4, If(GetItem(i).SubItems(4).Text = "", 0, CDbl(GetItem(i).SubItems(4).Text)) + If(GetItem(i).SubItems(5).Text = "", 0, CDbl(GetItem(i).SubItems(5).Text)))
                        SetSubitemText(i, 7, If(GetItem(i).SubItems(7).Text = "", 0, CDbl(GetItem(i).SubItems(7).Text)) - If(GetItem(i).SubItems(5).Text = "", 0, CDbl(GetItem(i).SubItems(5).Text)))
                        SetSubitemText(i, 5, "")
                    End If
                End If
                SetItemChecked(i, False)
            End If
        Next
        bgw_write_overtime.ReportProgress(100)
    End Sub

    Private Sub bgw_write_overtime_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles bgw_write_overtime.RunWorkerCompleted
        If normalDPI Then
            WorkingForm.Close()
        Else
            LoadingForm.Close()
        End If
        Dim lvcs As New ListViewColumnSorter
        lvcs.SortingOrder = SortOrder.Descending
        lvcs.ColumnIndex = 4
        Me.select_employees.Columns(4).Tag = SortOrder.Descending
        Me.select_employees.ListViewItemSorter = lvcs
    End Sub

    Private Sub EXCEL_button_Click(sender As Object, e As EventArgs) Handles EXCEL_button.Click
        If Me.select_employees.Items.Count = 0 Then Exit Sub
        If Me.select_employees.CheckedItems.Count = 0 Then
            MsgBox("Kein Mitarbeiter ausgewählt!")
            Exit Sub
        End If
        ExportPDF.ToolTip1.SetToolTip(ExportPDF.CREATE_button, "EXCEL Datei erstellen und speichern")
        ExportPDF.Text = "Export EXCEL"
        ExportPDF.ShowDialog(Me)
    End Sub

    Private Sub PDF_button_Click(sender As Object, e As EventArgs) Handles PDF_button.Click
        If Me.select_employees.Items.Count = 0 Then Exit Sub
        If Me.select_employees.CheckedItems.Count = 0 Then
            MsgBox("Kein Mitarbeiter ausgewählt!")
            Exit Sub
        End If
        ExportPDF.ToolTip1.SetToolTip(ExportPDF.CREATE_button, "PDF Datei erstellen und speichern")
        ExportPDF.Text = "Export PDF"
        ExportPDF.ShowDialog(Me)
    End Sub

    Function CalculateColumnWidth(ByVal textLength As Integer) As Double
        Dim font = New System.Drawing.Font("Calibri", 11)
        Dim digitMaximumWidth As Single = 0

        Using graphics = System.Drawing.Graphics.FromHwnd(Process.GetCurrentProcess().MainWindowHandle)

            For i = 0 To 10 - 1
                Dim digitWidth = graphics.MeasureString(i.ToString(), font).Width
                If digitWidth > digitMaximumWidth Then digitMaximumWidth = digitWidth
            Next
        End Using

        Return Math.Truncate((textLength * digitMaximumWidth + 5.0) / digitMaximumWidth * 256.0) / 256.0
    End Function

End Class
