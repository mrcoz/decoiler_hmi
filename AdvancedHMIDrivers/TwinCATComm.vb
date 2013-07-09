﻿Public Class TwinCATComm
    Inherits System.ComponentModel.Component
    Implements IComComponent

    Public Event ComError As EventHandler

    Private Shared DLL(10) As MfgControl.AdvancedHMI.Drivers.ADSforTwinCAT
    Private MyDLLInstance As Integer

    Private tmrPollList As New List(Of Windows.Forms.Timer)

    Private Structure PolledAddressInfo
        Friend Symbol As MfgControl.AdvancedHMI.Drivers.ADSforTwinCAT.SymbolInfo
        Friend dlgCallBack As IComComponent.ReturnValues
        Friend PollRate As Integer
        Friend ID As Integer
    End Structure

    Private PolledAddressList As New List(Of PolledAddressInfo)


#Region "Properties"
    Private m_TargetIPAddress As String = "192.168.2.6"   '* this is a default value
    Public Property TargetIPAddress() As String
        Get
            Return m_TargetIPAddress.ToString
        End Get
        Set(ByVal value As String)
            m_TargetIPAddress = value

            '* If a new instance needs to be created, such as a different AMS Address
            CreateDLLInstance()


            If DLL(MyDLLInstance) IsNot Nothing Then
                DLL(MyDLLInstance).TargetIPAddress = value
            End If
        End Set
    End Property

    '* USing a small "a" ensures the designer will set this property first
    '* If anything else is set before, a new instance will not get created
    Private m_TargetAMSNetID As String = "0.0.0.0.0.0"
    Public Property aTargetAMSNetID() As String
        Get
            Return m_TargetAMSNetID
        End Get
        Set(ByVal value As String)
            m_TargetAMSNetID = value

            '* If a new instance needs to be created, such as a different AMS Address
            CreateDLLInstance()


            If DLL(MyDLLInstance) IsNot Nothing Then
                DLL(MyDLLInstance).TargetAMSNetID = value
            End If
        End Set
    End Property

    Private m_TargetAMSPort As UInt16 = 801
    Public Property TargetAMSPort() As UInt16
        Get
            Return m_TargetAMSPort
        End Get
        Set(ByVal value As UInt16)
            m_TargetAMSPort = value

            '* If a new instance needs to be created, such as a different AMS Address
            CreateDLLInstance()


            If DLL(MyDLLInstance) IsNot Nothing Then
                DLL(MyDLLInstance).TargetAMSPort = value
            End If
        End Set
    End Property

    '*****************************************************************
    '* The username and password used to register with the AMS router
    Private m_UserName As String = "Administrator"
    Public Property UserName() As String
        Get
            Return m_UserName
        End Get
        Set(ByVal value As String)
            m_UserName = value

            '* If a new instance needs to be created, such as a different AMS Address
            CreateDLLInstance()


            If DLL(MyDLLInstance) IsNot Nothing Then
                DLL(MyDLLInstance).UserName = value
            End If

        End Set
    End Property

    Private m_Password As String = "1"
    Public Property Password() As String
        Get
            Return m_Password
        End Get
        Set(ByVal value As String)
            m_Password = value

            '* If a new instance needs to be created, such as a different AMS Address
            CreateDLLInstance()


            If DLL(MyDLLInstance) IsNot Nothing Then
                DLL(MyDLLInstance).Password = value
            End If
        End Set
    End Property

    '**************************************************
    '* Its purpose is to fetch
    '* the main form in order to synchronize the
    '* notification thread/event
    '**************************************************
    Private m_SynchronizingObject As Windows.Forms.Form
    '* do not let this property show up in the property window
    ' <System.ComponentModel.Browsable(False)> _
    Public Property SynchronizingObject() As Windows.Forms.Form
        Get
            'If Me.Site.DesignMode Then

            Dim host1 As System.ComponentModel.Design.IDesignerHost
            Dim obj1 As Object
            If (m_SynchronizingObject Is Nothing) AndAlso MyBase.DesignMode Then
                host1 = CType(Me.GetService(GetType(System.ComponentModel.Design.IDesignerHost)), System.ComponentModel.Design.IDesignerHost)
                If host1 IsNot Nothing Then
                    obj1 = host1.RootComponent
                    m_SynchronizingObject = CType(obj1, System.ComponentModel.ISynchronizeInvoke)
                End If
            End If
            'End If
            Return m_SynchronizingObject
        End Get

        Set(ByVal Value As Windows.Forms.Form)
            If Not Value Is Nothing Then
                m_SynchronizingObject = Value
            End If
        End Set
    End Property

    Private m_DisableSubscriptions As Boolean
    Public Property DisableSubScriptions() As Boolean Implements IComComponent.DisableSubscriptions
        Get
            Return m_DisableSubscriptions
        End Get
        Set(ByVal value As Boolean)
            m_DisableSubscriptions = value

            If value Then
                '* Stop the poll timers
                For i As Int16 = 0 To tmrPollList.Count - 1
                    tmrPollList(i).Enabled = False
                Next
            Else
                '* Start the poll timers
                For i As Int16 = 0 To tmrPollList.Count - 1
                    tmrPollList(i).Enabled = True
                Next
            End If
        End Set
    End Property
#End Region

#Region "ConstructorDestructor"
    Private components As System.ComponentModel.IContainer

    Public Sub New(ByVal container As System.ComponentModel.IContainer)
        MyClass.New()

        'Required for Windows.Forms Class Composition Designer support
        container.Add(Me)
    End Sub

    Public Sub New()
        MyBase.New()

        'CreateDLLInstance()
    End Sub

    'Component overrides dispose to clean up the component list.
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        '* The handle linked to the DataLink Layer has to be removed, otherwise it causes a problem when a form is closed
        If DLL(MyDLLInstance) IsNot Nothing Then RemoveHandler DLL(MyDLLInstance).DataReceived, Dr

        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub
#End Region
    '***************************************************************
    '* Create the Data Link Layer Instances
    '* if the AMS NET Address is the same, then resuse a common instance
    '***************************************************************
    Private Sub CreateDLLInstance()
        '* Still default, so ignore
        If m_TargetAMSNetID = "0.0.0.0.0.0" Then Exit Sub

        If DLL(0) IsNot Nothing Then
            '* At least one DLL instance already exists,
            '* so check to see if it has the same IP address
            '* if so, reuse the instance, otherwise create a new one
            Dim i As Integer
            While DLL(i) IsNot Nothing AndAlso DLL(i).TargetAMSNetID <> m_TargetAMSNetID AndAlso i < 11
                i += 1
            End While
            MyDLLInstance = i
        End If

        If DLL(MyDLLInstance) Is Nothing Then
            DLL(MyDLLInstance) = New MfgControl.AdvancedHMI.Drivers.ADSforTwinCAT
        End If
        DLL(MyDLLInstance).TargetAMSNetID = m_TargetAMSNetID
        DLL(MyDLLInstance).TargetAMSPort = m_TargetAMSPort
        'If DLL(MyDLLInstance).TargetIPAddress <> m_TargetIPAddress Then DLL(MyDLLInstance).TargetIPAddress = m_TargetIPAddress
        DLL(MyDLLInstance).TargetIPAddress = m_TargetIPAddress

        DLL(MyDLLInstance).UserName = m_UserName
        DLL(MyDLLInstance).Password = m_Password
        'End If

        AddHandler DLL(MyDLLInstance).DataReceived, Dr
        AddHandler DLL(MyDLLInstance).ComError, AddressOf ComErrorFromDLL
    End Sub

    Private Sub ComErrorFromDLL(ByVal sender As Object, ByVal e As MfgControl.AdvancedHMI.Drivers.Common.PlcComErrorEventArgs)
        RaiseEvent ComError(sender, e)
    End Sub

    '*******************************************************************
    '*******************************************************************
    Private CurrentID As Integer
    Public Function Subscribe(ByVal plcAddress As String, ByVal numberOfElements As Int16, ByVal pollRate As Integer, ByVal callback As IComComponent.ReturnValues) As Integer Implements IComComponent.Subscribe
        Dim i As Integer

        '*********************************
        '* Get the symbol information
        '*********************************
        Try
            If DLL(MyDLLInstance) IsNot Nothing And m_TargetAMSNetID <> "0.0.0.0.0.0" Then
                i = DLL(MyDLLInstance).GetSymbolInfo(plcAddress)
            Else
                Throw New MfgControl.AdvancedHMI.Drivers.Common.PLCDriverException(-3, "DLL Layer not yet created")
            End If
        Catch ex As MfgControl.AdvancedHMI.Drivers.Common.PLCDriverException
            If ex.ErrorCode = 1808 Then
                Throw New MfgControl.AdvancedHMI.Drivers.Common.PLCDriverException(ex.ErrorCode, """" & plcAddress & """ PLC address not found. ADS Error " & ex.ErrorCode)
            Else
                Throw
            End If
        End Try

        '************************
        '* Error returned
        '***********************
        If i < 0 Then
            Throw New MfgControl.AdvancedHMI.Drivers.Common.PLCDriverException(i, "Failed to get symbol")
        End If

        '***********************************************************
        '* Check if there was already a subscription made for this
        '***********************************************************
        Dim Sym As MfgControl.AdvancedHMI.Drivers.ADSforTwinCAT.SymbolInfo = DLL(MyDLLInstance).UsedSymbols(i)
        Dim SubExists As Boolean
        Dim k As Integer

        If DLL(MyDLLInstance).UsedSymbols.Count > 0 Then
            For k = 0 To PolledAddressList.Count - 1
                If PolledAddressList(k).Symbol.Name = plcAddress And PolledAddressList(k).dlgCallBack = callback Then
                    SubExists = True
                    Exit For
                End If
            Next
        End If


        If (Not SubExists) Then
            '* The ID is used as a reference for removing polled addresses
            CurrentID += 1

            Dim tmpPA As New PolledAddressInfo

            tmpPA.PollRate = pollRate
            tmpPA.dlgCallBack = callback
            tmpPA.ID = CurrentID
            tmpPA.Symbol = Sym

            PolledAddressList.Add(tmpPA)

            PolledAddressList.Sort(AddressOf SortPolledAddresses)



            '********************************************************************
            '* Check to see if there already exists a timer for this poll rate
            '********************************************************************
            Dim j As Integer = 0
            While j < tmrPollList.Count AndAlso tmrPollList(j) IsNot Nothing AndAlso tmrPollList(j).Interval <> pollRate
                j += 1
            End While

            If j >= tmrPollList.Count Then
                '* Add new timer
                Dim tmrTemp As New Windows.Forms.Timer
                If pollRate > 0 Then
                    tmrTemp.Interval = pollRate
                Else
                    tmrTemp.Interval = 250
                End If

                tmrPollList.Add(tmrTemp)
                AddHandler tmrPollList(j).Tick, AddressOf PollUpdate

                tmrTemp.Enabled = True
            End If

            Return tmpPA.ID
        Else
            '* Return the subscription that already exists
            Return PolledAddressList(k).ID
        End If
        'Return -1
    End Function

    '***************************************************************
    '* Used to sort polled addresses by File Number and element
    '* This helps in optimizing reading
    '**************************************************************
    Private Function SortPolledAddresses(ByVal A1 As PolledAddressInfo, ByVal A2 As PolledAddressInfo) As Integer
        If A1.Symbol.IndexGroup = A2.Symbol.IndexGroup Then
            If A1.Symbol.IndexOffset > A2.Symbol.IndexOffset Then
                Return 1
            ElseIf A1.Symbol.IndexOffset = A2.Symbol.IndexOffset Then
                Return 0
            Else
                Return -1
            End If
        Else

            If A1.Symbol.IndexGroup > A2.Symbol.IndexGroup Then
                Return 1
            Else
                Return -1
            End If
        End If
    End Function

    '**************************************************************
    '* Perform the reads for the variables added for notification
    '* Attempt to optimize by grouping reads
    '**************************************************************
    Private InternalRequest As Boolean '* This is used to dinstinquish when to send data back to notification request
    Private Sub PollUpdate(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'If m_DisableSubscriptions Then Exit Sub

        Dim intTimerIndex As Integer = tmrPollList.IndexOf(sender)

        '* Stop the poll timer
        tmrPollList(intTimerIndex).Enabled = False

        Dim value(0) As String
        Dim i As Integer
        While i < PolledAddressList.Count
            Dim values() As Byte
            Dim NumberToRead As Int16 = 1
            Dim BytesToRead As Integer = PolledAddressList(i).Symbol.ByteCount
            Try
                '* TODO : optimize in as few reads as possible - try group reading
                '* Reading consecuctive addresses only seems to work on %M (&H4020), Variable Memory (&H4040)
                Dim AddressSpan As Integer = 200
                If PolledAddressList(i).Symbol.IndexGroup <> &H4040 And PolledAddressList(i).Symbol.IndexGroup <> &H4020 Then
                    '* Do a Sum Up Read
                    AddressSpan = 0
                    Dim IndexGroups As New System.Collections.Generic.List(Of UInt32)
                    Dim IndexOffsets As New System.Collections.Generic.List(Of UInt32)
                    Dim ReadLength As New System.Collections.Generic.List(Of UInt32)

                    NumberToRead = 0
                    Dim j, ReadBytes As Integer
                    ReadBytes = 0
                    While (j + i) < PolledAddressList.Count AndAlso (ReadBytes < 200 And PolledAddressList(j + i).Symbol.IndexGroup <> &H4040 And PolledAddressList(j + i).Symbol.IndexGroup <> &H4020)
                        IndexGroups.Add(PolledAddressList(j + i).Symbol.IndexGroup)
                        IndexOffsets.Add(PolledAddressList(j + i).Symbol.IndexOffset)
                        ReadLength.Add(PolledAddressList(j + i).Symbol.ByteCount)
                        ReadBytes += PolledAddressList(j + i).Symbol.ByteCount
                        NumberToRead += 1
                        j += 1
                    End While

                    values = DLL(MyDLLInstance).ReadGroup(IndexGroups, IndexOffsets, ReadLength)

                Else
                    '***************************************************************
                    '* Read multiple values when they are close together in memory
                    '***************************************************************
                    While (i + NumberToRead) < PolledAddressList.Count AndAlso (PolledAddressList(i).Symbol.IndexGroup = PolledAddressList(i + NumberToRead).Symbol.IndexGroup) AndAlso (CInt(PolledAddressList(i + NumberToRead).Symbol.IndexOffset) - PolledAddressList(i).Symbol.IndexOffset) <= AddressSpan
                        'BytesToRead += PolledAddressList(i + NumberToRead).Symbol.ByteCount
                        BytesToRead = (PolledAddressList(i + NumberToRead).Symbol.IndexOffset - PolledAddressList(i).Symbol.IndexOffset) + PolledAddressList(i + NumberToRead).Symbol.ByteCount
                        NumberToRead += 1
                    End While

                    'Try
                    values = DLL(MyDLLInstance).ReadAny(PolledAddressList(i).Symbol.IndexGroup, PolledAddressList(i).Symbol.IndexOffset, BytesToRead)
                    'Catch ex As Exception
                    'TODO
                    'Dim dbg As Integer = 0
                    'value(0) = ex.Message
                    'End Try
                End If

                Dim ArrayIndex As Int16 = 0
                '* Extract the values for each symbol and send to subscriber
                For j As Integer = i To i + NumberToRead - 1
                    Select Case PolledAddressList(j).Symbol.DataType
                        Case MfgControl.AdvancedHMI.Drivers.ADSforTwinCAT.ADSDataType.Int8
                            value(0) = values(ArrayIndex)
                        Case MfgControl.AdvancedHMI.Drivers.ADSforTwinCAT.ADSDataType.UInt8
                            value(0) = values(ArrayIndex)
                        Case MfgControl.AdvancedHMI.Drivers.ADSforTwinCAT.ADSDataType.Int16
                            value(0) = BitConverter.ToInt16(values, ArrayIndex)
                        Case MfgControl.AdvancedHMI.Drivers.ADSforTwinCAT.ADSDataType.UInt16
                            value(0) = BitConverter.ToUInt16(values, ArrayIndex)
                        Case MfgControl.AdvancedHMI.Drivers.ADSforTwinCAT.ADSDataType.Int32
                            value(0) = BitConverter.ToInt32(values, ArrayIndex)
                        Case MfgControl.AdvancedHMI.Drivers.ADSforTwinCAT.ADSDataType.UInt32
                            value(0) = BitConverter.ToUInt32(values, ArrayIndex)
                        Case MfgControl.AdvancedHMI.Drivers.ADSforTwinCAT.ADSDataType.Int64
                            value(0) = BitConverter.ToInt64(values, ArrayIndex)
                        Case MfgControl.AdvancedHMI.Drivers.ADSforTwinCAT.ADSDataType.UInt64
                            value(0) = BitConverter.ToUInt64(values, ArrayIndex)
                        Case MfgControl.AdvancedHMI.Drivers.ADSforTwinCAT.ADSDataType.Real32
                            value(0) = BitConverter.ToSingle(values, ArrayIndex)
                        Case MfgControl.AdvancedHMI.Drivers.ADSforTwinCAT.ADSDataType.Real64
                            value(0) = BitConverter.ToDouble(values, ArrayIndex)
                        Case MfgControl.AdvancedHMI.Drivers.ADSforTwinCAT.ADSDataType.StringType
                            Dim NullPos As Integer = ArrayIndex
                            While ArrayIndex < values.Length And values(NullPos) <> 0
                                NullPos += 1
                            End While
                            value(0) = System.Text.Encoding.ASCII.GetString(values, ArrayIndex, NullPos - ArrayIndex).TrimEnd(Chr(0))
                        Case Else '* Bool type (33)
                            value(0) = values(ArrayIndex)
                    End Select

                    'm_SynchronizingObject.BeginInvoke(PolledAddressList(j).dlgCallBack, CObj(value))
                    m_SynchronizingObject.Invoke(PolledAddressList(j).dlgCallBack, CObj(value))

                    If j < (i + NumberToRead - 1) Then
                        If AddressSpan = 0 Then
                            '* A Sum Up Read was done
                            ArrayIndex += PolledAddressList(j).Symbol.ByteCount
                        Else
                            '* A block of memory was read, so point to the next value needed
                            ArrayIndex += PolledAddressList(j + 1).Symbol.IndexOffset - PolledAddressList(j).Symbol.IndexOffset
                        End If
                    End If
                Next
            Catch ex As Exception
                Dim y = 0
            End Try

            i += NumberToRead
        End While

        '* Start the poll timer
        tmrPollList(intTimerIndex).Enabled = True
    End Sub

    Public Function UnSubscribe(ByVal ID As Integer) As Integer Implements IComComponent.Unsubscribe
        Dim i As Integer = 0
        While i < PolledAddressList.Count AndAlso PolledAddressList(i).ID <> ID
            i += 1
        End While

        If i < PolledAddressList.Count Then
            PolledAddressList.RemoveAt(i)
            If PolledAddressList.Count = 0 Then
                '* No more items to be polled, so remove all polling timers 28-NOV-10
                For j As Integer = 0 To tmrPollList.Count - 1
                    'tmrPollList(j).Enabled = False
                    tmrPollList(j).Enabled = False
                    tmrPollList.Remove(tmrPollList(j))
                Next
            End If
        End If
    End Function

    Public Function ReadSynchronous(ByVal startAddress As String, ByVal numberOfElements As Integer) As String() Implements IComComponent.ReadSynchronous
        Return ReadAny(startAddress, numberOfElements, False)
    End Function

    Public Function ReadAny(ByVal startAddress As String, ByVal numberOfElements As Integer) As String() Implements IComComponent.ReadAny
        Return ReadAny(startAddress, numberOfElements, False)
    End Function


    Public Function ReadAny(ByVal startAddress As String) As String Implements IComComponent.ReadAny
        Return DLL(MyDLLInstance).ReadAny(startAddress)
    End Function

    Public Function ReadAny(ByVal startAddress As String, ByVal numberOfElements As Integer, ByVal AsyncModeIn As Boolean) As String()
        Return Nothing
    End Function

    Public Function WriteData(ByVal startAddress As String, ByVal dataToWrite As String) As String Implements IComComponent.WriteData
        Dim tmp() As String = {dataToWrite}
        If DLL(MyDLLInstance) Is Nothing Then CreateDLLInstance()
        Try
            Return DLL(MyDLLInstance).WriteAny(startAddress, tmp, 1)
        Catch ex As Exception
            MsgBox("WriteData1. " & ex.Message)
            Throw
        End Try
    End Function

    Public Function WriteData(ByVal startAddress As String, ByVal dataToWrite() As String, ByVal numberOfElements As UInt16) As Integer
        If DLL(MyDLLInstance) Is Nothing Then CreateDLLInstance()
        Try
            DLL(MyDLLInstance).WriteAny(startAddress, dataToWrite, numberOfElements)
        Catch ex As Exception
            MsgBox("WriteData2. " & ex.Message)
            Throw
        End Try
    End Function

    Private Dr As EventHandler = AddressOf DataLinkLayer_DataReceived
    Private Sub DataLinkLayer_DataReceived()
        '* Should we only raise an event if we are in AsyncMode?
        '        If m_AsyncMode Then
        '**************************************************************************
        '* If the parent form property (Synchronizing Object) is set, then sync the event with its thread
        '**************************************************************************
        '* Get the low byte from the Sequence Number
        '* The sequence number was added onto the end of the CIP packet by this code
        'Dim TNSReturned As Integer = DLL(MyDLLInstance).DataPacket(DLL(MyDLLInstance).DataPacket.Count - 2)
        'DataPackets(TNSReturned) = DLL(MyDLLInstance).DataPacket
        'Responded(TNSReturned) = True
    End Sub
End Class
