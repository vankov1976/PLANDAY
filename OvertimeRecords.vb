Public Class OvertimeRecords

    Private Sub select_employee_SelectedValueChanged(sender As Object, e As EventArgs) Handles select_employee.SelectedValueChanged
        If OvertimeForm.ImChangingStuff Then Exit Sub
        OvertimeForm.load_employee_records(select_employee.Text)
    End Sub

    Private Sub KW_Records_ItemSelectionChanged(sender As Object, e As ListViewItemSelectionChangedEventArgs) Handles KW_Records.ItemSelectionChanged
        e.Item.Selected = Nothing
    End Sub

    Private Sub select_employee_KeyDown(sender As Object, e As KeyEventArgs) Handles select_employee.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Close()
        End If
    End Sub

    Private Sub KW_Records_KeyDown(sender As Object, e As KeyEventArgs) Handles KW_Records.KeyDown
        select_employee.Select()

        If e.KeyCode = Keys.Escape Then
            Me.Close()
        End If
    End Sub

End Class