<Serializable>
Public Class PayrollPeriods
    Private m_Year As Integer
    Private m_January As MonthPeriod
    Private m_February As MonthPeriod
    Private m_March As MonthPeriod
    Private m_April As MonthPeriod
    Private m_May As MonthPeriod
    Private m_June As MonthPeriod
    Private m_July As MonthPeriod
    Private m_August As MonthPeriod
    Private m_September As MonthPeriod
    Private m_October As MonthPeriod
    Private m_November As MonthPeriod
    Private m_December As MonthPeriod
    Private m_Periods As MonthPeriod()

    Public Sub New(year As Integer, january As MonthPeriod, february As MonthPeriod, march As MonthPeriod, april As MonthPeriod, may As MonthPeriod, june As MonthPeriod, july As MonthPeriod, august As MonthPeriod, september As MonthPeriod, october As MonthPeriod, november As MonthPeriod, december As MonthPeriod)
        m_Year = year
        m_January = january
        m_February = february
        m_March = march
        m_April = april
        m_May = may
        m_June = june
        m_July = july
        m_August = august
        m_September = september
        m_October = october
        m_November = november
        m_December = december
        m_Periods = New MonthPeriod() {january, february, march, april, may, june, july, august, september, october, november, december}
    End Sub

    Public Property Year As Integer
        Get
            Return m_Year
        End Get
        Set(value As Integer)
            m_Year = value
        End Set
    End Property

    Public Property January As MonthPeriod
        Get
            Return m_January
        End Get
        Set(value As MonthPeriod)
            m_January = value
        End Set
    End Property

    Public Property February As MonthPeriod
        Get
            Return m_February
        End Get
        Set(value As MonthPeriod)
            m_February = value
        End Set
    End Property

    Public Property March As MonthPeriod
        Get
            Return m_March
        End Get
        Set(value As MonthPeriod)
            m_March = value
        End Set
    End Property

    Public Property April As MonthPeriod
        Get
            Return m_April
        End Get
        Set(value As MonthPeriod)
            m_April = value
        End Set
    End Property

    Public Property May As MonthPeriod
        Get
            Return m_May
        End Get
        Set(value As MonthPeriod)
            m_May = value
        End Set
    End Property

    Public Property June As MonthPeriod
        Get
            Return m_June
        End Get
        Set(value As MonthPeriod)
            m_June = value
        End Set
    End Property

    Public Property July As MonthPeriod
        Get
            Return m_July
        End Get
        Set(value As MonthPeriod)
            m_July = value
        End Set
    End Property

    Public Property August As MonthPeriod
        Get
            Return m_August
        End Get
        Set(value As MonthPeriod)
            m_August = value
        End Set
    End Property

    Public Property September As MonthPeriod
        Get
            Return m_September
        End Get
        Set(value As MonthPeriod)
            m_September = value
        End Set
    End Property

    Public Property October As MonthPeriod
        Get
            Return m_October
        End Get
        Set(value As MonthPeriod)
            m_October = value
        End Set
    End Property

    Public Property November As MonthPeriod
        Get
            Return m_November
        End Get
        Set(value As MonthPeriod)
            m_November = value
        End Set
    End Property

    Public Property December As MonthPeriod
        Get
            Return m_December
        End Get
        Set(value As MonthPeriod)
            m_December = value
        End Set
    End Property

    Public Property Periods As MonthPeriod()
        Get
            Return m_Periods
        End Get
        Set(value As MonthPeriod())
            m_Periods = value
        End Set
    End Property
End Class
