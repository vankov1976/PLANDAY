<Serializable>
Public Class MonthPeriod
    Private m_StartDate As Date
    Private m_EndDate As Date

    Public Sub New(startDate As Date, endDate As Date)
        m_StartDate = startDate
        m_EndDate = endDate
    End Sub

    Public Property StartDate As Date
        Get
            Return m_StartDate
        End Get
        Set(value As Date)
            m_StartDate = value
        End Set
    End Property

    Public Property EndDate As Date
        Get
            Return m_EndDate
        End Get
        Set(value As Date)
            m_EndDate = value
        End Set
    End Property
End Class
