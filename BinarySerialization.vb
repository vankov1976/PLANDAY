Imports System.IO

Public Class BinarySerialization
    Sub WriteToBinaryFile(Of T)(ByVal filePath As String, ByVal objectToWrite As T, ByVal Optional append As Boolean = False)
        Using stream As Stream = File.Open(filePath, If(append, FileMode.Append, FileMode.Create))
            Dim binaryFormatter = New System.Runtime.Serialization.Formatters.Binary.BinaryFormatter()
            binaryFormatter.Serialize(stream, objectToWrite)
        End Using
    End Sub

    Function ReadFromBinaryFile(Of T)(ByVal filePath As String) As T
        Using stream As Stream = File.Open(filePath, FileMode.Open)
            Dim binaryFormatter = New System.Runtime.Serialization.Formatters.Binary.BinaryFormatter()
            Return CType(binaryFormatter.Deserialize(stream), T)
        End Using
    End Function

    Private Sub Out_()

        'BinarySerialization.WriteToBinaryFile(Of PayrollPeriods)("C:person.bin", test)
    End Sub

    Private Sub In_()
        'Dim test As PayrollPeriods = ReadFromBinaryFile(Of PayrollPeriods)("C:person.bin")
    End Sub
End Class
