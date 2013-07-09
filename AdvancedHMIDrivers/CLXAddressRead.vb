Public Class CLXAddressRead
    Inherits MfgControl.AdvancedHMI.Drivers.CLXAddress

    Private m_TransactionNumber As UInt16
    Public Property TransactionNumber As UInt16
        Get
            Return m_TransactionNumber
        End Get
        Set(value As UInt16)
            m_TransactionNumber = value
        End Set
    End Property

    Private m_InternalRequest As Boolean
    Public Property InternalRequest As Boolean
        Get
            Return m_InternalRequest
        End Get
        Set(value As Boolean)
            m_InternalRequest = value
        End Set
    End Property

    Private m_AsynMode As Boolean
    Public Property AsyncMode As Boolean
        Get
            Return m_AsynMode
        End Get
        Set(value As Boolean)
            m_AsynMode = value
        End Set
    End Property
End Class
