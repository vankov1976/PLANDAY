Imports System.Runtime.InteropServices

Public Class GetDPI
    <DllImport("gdi32.dll", CharSet:=CharSet.Auto, SetLastError:=True, ExactSpelling:=True)>
    Public Shared Function GetDeviceCaps(ByVal hDC As IntPtr, ByVal nIndex As Integer) As Integer

    End Function

    Public Enum DeviceCap
        LOGPIXELSX = 88
        LOGPIXELSY = 90
    End Enum
End Class