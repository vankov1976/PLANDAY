
Namespace My
    ' The following events are available for MyApplication:
    ' Startup: Raised when the application starts, before the startup form is created.
    ' Shutdown: Raised after all application forms are closed.  This event is not raised if the application terminates abnormally.
    ' UnhandledException: Raised if the application encounters an unhandled exception.
    ' StartupNextInstance: Raised when launching a single-instance application and the application is already active. 
    ' NetworkAvailabilityChanged: Raised when the network connection is connected or disconnected.
    Partial Friend Class MyApplication
        Private Sub AppStart(ByVal sender As Object, ByVal e As Microsoft.VisualBasic.ApplicationServices.StartupEventArgs) Handles Me.Startup
            AddHandler AppDomain.CurrentDomain.AssemblyResolve, AddressOf ResolveAssemblies
        End Sub

        Private Function ResolveAssemblies(sender As Object, e As System.ResolveEventArgs) As Reflection.Assembly
            Dim ressourceName = "PLANDAY." + New Reflection.AssemblyName(e.Name).Name + ".dll"
            If ressourceName = "PLANDAY.itext.licensekey.dll" Then Exit Function
            Using stream = Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream(ressourceName)
                Dim assemblyData(CInt(stream.Length)) As Byte
                stream.Read(assemblyData, 0, assemblyData.Length)
                Return Reflection.Assembly.Load(assemblyData)
            End Using
        End Function
    End Class
End Namespace
