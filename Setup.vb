Public Class Setup

    Private first_load_done As Boolean
    Private Sub Setup_Load(sender As Object, e As EventArgs) Handles Me.Load
        If first_load_done Then Exit Sub
        first_load_done = True
        Dim index As Integer
        Dim row As String()
        If OvertimeForm.PLANDAY_portal = "DE" Then

            For Each shift In My.Settings.DE_shifts
                index = My.Settings.DE_shifts.IndexOf(shift)
                row = New String() {My.Settings.DE_shifts_names.Item(index), My.Settings.DE_shifts_surcharges.Item(index), My.Settings.DE_shifts_overtime.Item(index), My.Settings.DE_shifts_sick_paid.Item(index), My.Settings.DE_shifts_sick_surcharges.Item(index)}
                DataGridView1.Rows.Add(row)
            Next
            For Each rw As DataGridViewRow In DataGridView1.Rows
                disableCell(rw.Cells(1), rw.Cells(4).Value)
            Next

            For Each department In My.Settings.DE_departments
                index = My.Settings.DE_departments.IndexOf(department)
                If My.Settings.DE_departments_bundesland.Item(index) = "not_set" Then
                    row = New String() {My.Settings.DE_departments_names.Item(index)}
                Else
                    row = New String() {My.Settings.DE_departments_names.Item(index), My.Settings.DE_departments_bundesland.Item(index)}
                End If
                DataGridView2.Rows.Add(row)
            Next

        End If
        If OvertimeForm.PLANDAY_portal = "HQ" Then

            For Each shift In My.Settings.HQ_shifts
                index = My.Settings.HQ_shifts.IndexOf(shift)
                row = New String() {My.Settings.HQ_shifts_names.Item(index), My.Settings.HQ_shifts_surcharges.Item(index), My.Settings.HQ_shifts_overtime.Item(index), My.Settings.HQ_shifts_sick_paid.Item(index), My.Settings.HQ_shifts_sick_surcharges.Item(index)}
                DataGridView1.Rows.Add(row)
            Next
            For Each rw As DataGridViewRow In DataGridView1.Rows
                disableCell(rw.Cells(1), rw.Cells(4).Value)
            Next

            For Each department In My.Settings.HQ_departments
                index = My.Settings.HQ_departments.IndexOf(department)
                If My.Settings.HQ_departments_bundesland.Item(index) = "not_set" Then
                    row = New String() {My.Settings.HQ_departments_names.Item(index)}
                Else
                    row = New String() {My.Settings.HQ_departments_names.Item(index), My.Settings.HQ_departments_bundesland.Item(index)}
                End If
                DataGridView2.Rows.Add(row)
            Next

        End If

        'Assign Click event to the DataGridView Cell.
        AddHandler DataGridView1.CellContentClick, AddressOf DataGridView_CellClick
    End Sub
    Private Sub DataGridView_CellClick(ByVal sender As Object, ByVal e As DataGridViewCellEventArgs)
        Dim index As Integer
        'Check to ensure that the row CheckBox is clicked.
        If e.RowIndex >= 0 AndAlso e.ColumnIndex > 0 Then

            'Reference the GridView Row.
            Dim row As DataGridViewRow = DataGridView1.Rows(e.RowIndex)
            If OvertimeForm.PLANDAY_portal = "DE" Then
                index = My.Settings.DE_shifts_names.IndexOf(row.Cells(0).Value)
            End If

            If OvertimeForm.PLANDAY_portal = "HQ" Then
                index = My.Settings.HQ_shifts_names.IndexOf(row.Cells(0).Value)
            End If

            'Set the CheckBox selection.
            If Not row.Cells(e.ColumnIndex).ReadOnly Then row.Cells(e.ColumnIndex).Value = Convert.ToBoolean(row.Cells(e.ColumnIndex).EditedFormattedValue)

                Select Case e.ColumnIndex
                    Case 1
                        If OvertimeForm.PLANDAY_portal = "DE" Then
                        If Not row.Cells(e.ColumnIndex).ReadOnly Then My.Settings.DE_shifts_surcharges.Item(index) = row.Cells(e.ColumnIndex).Value
                    End If
                        If OvertimeForm.PLANDAY_portal = "HQ" Then
                        If Not row.Cells(e.ColumnIndex).ReadOnly Then My.Settings.HQ_shifts_surcharges.Item(index) = row.Cells(e.ColumnIndex).Value
                    End If
                    Case 2
                        If OvertimeForm.PLANDAY_portal = "DE" Then
                        My.Settings.DE_shifts_overtime.Item(index) = row.Cells(e.ColumnIndex).Value
                    End If
                        If OvertimeForm.PLANDAY_portal = "HQ" Then
                        My.Settings.HQ_shifts_overtime.Item(index) = row.Cells(e.ColumnIndex).Value
                    End If
                    Case 3
                        If OvertimeForm.PLANDAY_portal = "DE" Then
                        My.Settings.DE_shifts_sick_paid.Item(index) = row.Cells(e.ColumnIndex).Value
                    End If
                        If OvertimeForm.PLANDAY_portal = "HQ" Then
                        My.Settings.HQ_shifts_sick_paid.Item(index) = row.Cells(e.ColumnIndex).Value
                    End If
                    Case 4
                        If OvertimeForm.PLANDAY_portal = "DE" Then
                        My.Settings.DE_shifts_sick_surcharges.Item(index) = row.Cells(e.ColumnIndex).Value
                        disableCell(row.Cells(1), row.Cells(e.ColumnIndex).Value)
                        End If
                        If OvertimeForm.PLANDAY_portal = "HQ" Then
                        My.Settings.HQ_shifts_sick_surcharges.Item(index) = row.Cells(e.ColumnIndex).Value
                        disableCell(row.Cells(1), row.Cells(e.ColumnIndex).Value)
                        End If
                End Select


            End If
    End Sub

    Private Sub disableCell(cell As DataGridViewCheckBoxCell, disable As Boolean)
        'toggle read-only state
        cell.ReadOnly = disable
    End Sub

    Private Sub DataGridView2_MouseMove(sender As Object, e As MouseEventArgs) Handles DataGridView2.MouseMove
        Dim ht As DataGridView.HitTestInfo = DataGridView2.HitTest(e.X, e.Y)

        ' Can use this to auto-activate the control
        If Not ht Is DataGridView.HitTestInfo.Nowhere Then
            If ht.ColumnIndex >= 0 AndAlso ht.ColumnIndex < DataGridView2.Columns.Count Then
                Dim col As DataGridViewColumn = DataGridView2.Columns(ht.ColumnIndex)

                If TypeOf (col) Is DataGridViewComboBoxColumn Then
                    If Not ht.RowIndex = DataGridView2.NewRowIndex AndAlso ht.RowIndex >= 0 _
             AndAlso ht.RowIndex < DataGridView2.RowCount Then
                        DataGridView2.CurrentCell = DataGridView2.Rows(ht.RowIndex).Cells(ht.ColumnIndex)
                        DataGridView2.BeginEdit(True)
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub DataGridView2_EditingControlShowing(sender As Object, e As DataGridViewEditingControlShowingEventArgs) Handles DataGridView2.EditingControlShowing
        ' Only for a DatagridComboBoxColumn at ColumnIndex 1.
        If DataGridView2.CurrentCell.ColumnIndex = 1 Then
            Dim combo As ComboBox = CType(e.Control, ComboBox)
            If (combo IsNot Nothing) Then
                ' Remove an existing event-handler, if present, to avoid 
                ' adding multiple handlers when the editing control is reused.
                RemoveHandler combo.SelectionChangeCommitted, New EventHandler(AddressOf ComboBox_SelectionChangeCommitted)

                ' Add the event handler. 
                AddHandler combo.SelectionChangeCommitted, New EventHandler(AddressOf ComboBox_SelectionChangeCommitted)
            End If
        End If
    End Sub

    Private Sub ComboBox_SelectionChangeCommitted(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim index As Integer
        Dim combo As ComboBox = CType(sender, ComboBox)
        Dim row As DataGridViewRow = DataGridView2.Rows(DataGridView2.CurrentCell.RowIndex)

        If OvertimeForm.PLANDAY_portal = "DE" Then
            index = My.Settings.DE_departments_names.IndexOf(row.Cells(0).Value)
        End If

        If OvertimeForm.PLANDAY_portal = "HQ" Then
            index = My.Settings.HQ_departments_names.IndexOf(row.Cells(0).Value)
        End If

        Debug.Print("Row: {0}, Value: {1}", DataGridView2.CurrentCell.RowIndex, combo.SelectedItem)
        If OvertimeForm.PLANDAY_portal = "DE" Then
            My.Settings.DE_departments_bundesland.Item(index) = combo.SelectedItem
        End If
        If OvertimeForm.PLANDAY_portal = "HQ" Then
            My.Settings.HQ_departments_bundesland.Item(index) = combo.SelectedItem
        End If
    End Sub

End Class