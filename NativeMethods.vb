Imports System.Runtime.InteropServices
Class NativeMethods
    Private Const LVM_FIRST As Integer = &H1000
    Private Const LVM_SETITEMSTATE As Integer = LVM_FIRST + 43

    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Auto)>
    Public Structure LVITEM
        Public mask As Integer
        Public iItem As Integer
        Public iSubItem As Integer
        Public state As Integer
        Public stateMask As Integer
        <MarshalAs(UnmanagedType.LPTStr)>
        Public pszText As String
        Public cchTextMax As Integer
        Public iImage As Integer
        Public lParam As IntPtr
        Public iIndent As Integer
        Public iGroupId As Integer
        Public cColumns As Integer
        Public puColumns As IntPtr
    End Structure

    <DllImport("user32.dll", EntryPoint:="SendMessage", CharSet:=CharSet.Auto)>
    Public Shared Function SendMessageLVItem(ByVal hWnd As IntPtr, ByVal msg As Integer, ByVal wParam As Integer, ByRef lvi As LVITEM) As IntPtr

    End Function

    Public Shared Sub SelectAllItems(ByVal list As ListView)
        NativeMethods.SetItemState(list, -1, 2, 2)
    End Sub

    Public Shared Sub DeselectAllItems(ByVal list As ListView)
        NativeMethods.SetItemState(list, -1, 2, 0)
    End Sub

    Public Shared Sub SetItemState(ByVal list As ListView, ByVal itemIndex As Integer, ByVal mask As Integer, ByVal value As Integer)
        Dim lvItem As LVITEM = New LVITEM()
        lvItem.stateMask = mask
        lvItem.state = value
        SendMessageLVItem(list.Handle, LVM_SETITEMSTATE, itemIndex, lvItem)
    End Sub
End Class
