Public Class EmployeeRecord
    Private m_employee_id As String
    Private m_Mitarbeiter As String
    Private m_Soll As Double
    Private m_Ist As Double
    Private m_Überstunden As Double
    Private m_zuÜbertragen As Double
    Private m_Übertragen As Double
    Private m_vonPlanday As Double
    Private m_Kontostand As Double

    Public Sub New(id As String, employee As String, must As Double, actual As Double,
                   overtime As Double, to_write As Double, written As Double, written_fromPlanday As Double, balance As Double)
        m_employee_id = id
        m_Mitarbeiter = employee
        m_Soll = must
        m_Ist = actual
        m_Überstunden = overtime
        m_zuÜbertragen = to_write
        m_Übertragen = written
        m_vonPlanday = written_fromPlanday
        m_Kontostand = balance
    End Sub

    Public Property employee_id() As String

        Get
            Return m_employee_id
        End Get

        Set(Value As String)
            m_employee_id = Value
        End Set

    End Property

    Public Property Mitarbeiter() As String

        Get
            Return m_Mitarbeiter
        End Get

        Set(Value As String)
            m_Mitarbeiter = Value
        End Set

    End Property

    Public Property Soll() As Double

        Get
            Return m_Soll
        End Get

        Set(Value As Double)
            m_Soll = Value
        End Set

    End Property

    Public Property Ist() As Double

        Get
            Return m_Ist
        End Get

        Set(Value As Double)
            m_Ist = Value
        End Set

    End Property

    Public Property Überstunden() As Double

        Get
            Return m_Überstunden
        End Get

        Set(Value As Double)
            m_Überstunden = Value
        End Set

    End Property

    Public Property zuÜbertragen() As Double

        Get
            Return m_zuÜbertragen
        End Get

        Set(Value As Double)
            m_zuÜbertragen = Value
        End Set

    End Property

    Public Property Übertragen() As Double

        Get
            Return m_Übertragen
        End Get

        Set(Value As Double)
            m_Übertragen = Value
        End Set

    End Property

    Public Property vonPlanday() As Double

        Get
            Return m_vonPlanday
        End Get

        Set(Value As Double)
            m_vonPlanday = Value
        End Set

    End Property

    Public Property Kontostand() As Double

        Get
            Return m_Kontostand
        End Get

        Set(Value As Double)
            m_Kontostand = Value
        End Set

    End Property

End Class
