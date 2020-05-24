Imports System.ComponentModel
Imports PLANDAY.OvertimeForm

Public Class PayrollPeriodForm

    Private Sub select_year_SelectedIndexChanged(sender As Object, e As EventArgs) Handles select_year.SelectedIndexChanged
        Jan_Start.Value = payroll_periods(select_year.Text).January.StartDate
        Jan_End.Value = payroll_periods(select_year.Text).January.EndDate

        Feb_Start.Value = payroll_periods(select_year.Text).February.StartDate
        Feb_End.Value = payroll_periods(select_year.Text).February.EndDate

        Mar_Start.Value = payroll_periods(select_year.Text).March.StartDate
        Mar_End.Value = payroll_periods(select_year.Text).March.EndDate

        Apr_Start.Value = payroll_periods(select_year.Text).April.StartDate
        Apr_End.Value = payroll_periods(select_year.Text).April.EndDate

        May_Start.Value = payroll_periods(select_year.Text).May.StartDate
        May_End.Value = payroll_periods(select_year.Text).May.EndDate

        Jun_Start.Value = payroll_periods(select_year.Text).June.StartDate
        Jun_End.Value = payroll_periods(select_year.Text).June.EndDate

        Jul_Start.Value = payroll_periods(select_year.Text).July.StartDate
        Jul_End.Value = payroll_periods(select_year.Text).July.EndDate

        Aug_Start.Value = payroll_periods(select_year.Text).August.StartDate
        Aug_End.Value = payroll_periods(select_year.Text).August.EndDate

        Sep_Start.Value = payroll_periods(select_year.Text).September.StartDate
        Sep_End.Value = payroll_periods(select_year.Text).September.EndDate

        Oct_Start.Value = payroll_periods(select_year.Text).October.StartDate
        Oct_End.Value = payroll_periods(select_year.Text).October.EndDate

        Nov_Start.Value = payroll_periods(select_year.Text).November.StartDate
        Nov_End.Value = payroll_periods(select_year.Text).November.EndDate

        Dec_Start.Value = payroll_periods(select_year.Text).December.StartDate
        Dec_End.Value = payroll_periods(select_year.Text).December.EndDate
    End Sub

    Private Sub CLOSE_button_Click(sender As Object, e As EventArgs) Handles CLOSE_button.Click
        Me.Close()
    End Sub

    Private Sub SELECT_button_Click(sender As Object, e As EventArgs) Handles SELECT_button.Click
        ExportForm.YEAR_button.Text = payroll_periods(select_year.Text).Year
        If payroll_periods(select_year.Text).Year = 2019 Then
            ExportForm.select_month.Items.Clear()
            For i = 7 To 12
                ExportForm.select_month.Items.Add(System.Globalization.CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(i))
            Next
            ExportForm.select_month.SelectedIndex = 0
        End If
        OvertimeForm.set_params()
        Me.Close()
    End Sub

    Private Sub PayrollPeriodForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        select_year.Select()
    End Sub
End Class