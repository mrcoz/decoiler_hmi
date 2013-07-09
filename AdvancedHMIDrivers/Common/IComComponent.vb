Public Interface IComComponent
    Delegate Sub ReturnValues(ByVal values As String())
    Function Subscribe(ByVal plcAddress As String, ByVal numberOfElements As Int16, ByVal pollRate As Integer, ByVal callback As ReturnValues) As Integer
    Function Unsubscribe(ByVal id As Integer) As Integer
    Function ReadAny(ByVal startAddress As String) As String
    Function ReadAny(ByVal startAddress As String, ByVal numberOfElements As Integer) As String()
    Function ReadSynchronous(ByVal startAddress As String, ByVal numberOfElements As Integer) As String()
    Function WriteData(ByVal startAddress As String, ByVal dataToWrite As String) As String
    Property DisableSubscriptions() As Boolean
End Interface
