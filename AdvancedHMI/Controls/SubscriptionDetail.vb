Friend Class SubscriptionDetail
    Friend PLCAddress As String
    Friend NotificationID As Integer
    Friend Callback As AdvancedHMIDrivers.IComComponent.ReturnValues

    Public Sub New()
    End Sub

    Public Sub New(ByVal plcAddress As String, ByVal notificationID As Integer, ByVal callback As AdvancedHMIDrivers.IComComponent.ReturnValues)
        Me.PLCAddress = plcAddress
        Me.NotificationID = notificationID
        Me.Callback = callback
    End Sub

End Class
