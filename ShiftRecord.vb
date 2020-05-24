Public Class ShiftRecord
    Private m_shift_date As Date
    Private m_shift_type As String
    Private m_shift_from As String
    Private m_shift_to As String
    Private m_shift_break_1_from As String = ""
    Private m_shift_break_1_to As String = ""
    Private m_shift_break_2_from As String = ""
    Private m_shift_break_2_to As String = ""
    Private m_shift_length As Double
    Private m_holiday_days As Double
    Public Sub New()
    End Sub
    Public Sub New(shift_date As Date, shift_type As String, shift_from As String, shift_to As String, shift_break_1_from As String, shift_break_1_to As String, shift_break_2_from As String, shift_break_2_to As String, shift_length As Double, holiday_days As Double)
        m_shift_date = shift_date
        m_shift_type = shift_type
        m_shift_from = shift_from
        m_shift_to = shift_to
        m_shift_break_1_from = shift_break_1_from
        m_shift_break_1_to = shift_break_1_to
        m_shift_break_2_from = shift_break_2_from
        m_shift_break_2_to = shift_break_2_to
        m_shift_length = shift_length
        m_holiday_days = holiday_days
    End Sub

    Public Property Shift_date As Date
        Get
            Return m_shift_date
        End Get
        Set(value As Date)
            m_shift_date = value
        End Set
    End Property

    Public Property Shift_type As String
        Get
            Return m_shift_type
        End Get
        Set(value As String)
            m_shift_type = value
        End Set
    End Property

    Public Property Shift_from As String
        Get
            Return m_shift_from
        End Get
        Set(value As String)
            m_shift_from = value
        End Set
    End Property

    Public Property Shift_to As String
        Get
            Return m_shift_to
        End Get
        Set(value As String)
            m_shift_to = value
        End Set
    End Property

    Public Property Shift_break_1_from As String
        Get
            Return m_shift_break_1_from
        End Get
        Set(value As String)
            m_shift_break_1_from = value
        End Set
    End Property

    Public Property Shift_break_1_to As String
        Get
            Return m_shift_break_1_to
        End Get
        Set(value As String)
            m_shift_break_1_to = value
        End Set
    End Property

    Public Property Shift_break_2_from As String
        Get
            Return m_shift_break_2_from
        End Get
        Set(value As String)
            m_shift_break_2_from = value
        End Set
    End Property

    Public Property Shift_break_2_to As String
        Get
            Return m_shift_break_2_to
        End Get
        Set(value As String)
            m_shift_break_2_to = value
        End Set
    End Property

    Public Property Shift_length As Double
        Get
            Return m_shift_length
        End Get
        Set(value As Double)
            m_shift_length = value
        End Set
    End Property

    Public Property Holiday_days As Double
        Get
            Return m_holiday_days
        End Get
        Set(value As Double)
            m_holiday_days = value
        End Set
    End Property
End Class
