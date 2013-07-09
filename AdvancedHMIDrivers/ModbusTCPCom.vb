'******************************************************************************
'* Modbus TCP Protocol Implementation
'*
'* Archie Jacobs
'* Manufacturing Automation, LLC
'* ajacobs@mfgcontrol.com
'* 13-OCT-11
'*
'* Copyright 2011 Archie Jacobs
'*
'* Implements driver for communication to ModbusTCP devices
'* 5-MAR-12 Fixed a bug where ReadAny would call itself instead of the overload
'* 9-JAN-13 When TNS was over 255 it would go out of bounds in Transactions array
'* 27-JAN-13 Add the second byte that some require for writing to bits
'*******************************************************************************
Public Class ModbusTcpCom
    Inherits System.ComponentModel.Component
    Implements AdvancedHMIDrivers.IComComponent

    Public Event DataReceived As EventHandler(Of MfgControl.AdvancedHMI.Drivers.Common.PlcComEventArgs)
    Public Event UnsolicitedMessageReceived As EventHandler
    Public Event ComError As EventHandler

    '* Use a shared Data Link Layer so multiple instances will not create multiple connections
    Private Shared DLL(10) As MfgControl.AdvancedHMI.Drivers.ModbusTCP.ModbusTcpDataLinkLayer
    Private MyDLLInstance As Integer

    Private tmrPollList As New List(Of System.Timers.Timer)

    Private PolledAddressList As New List(Of PolledAddressInfo)

    Private Structure PolledAddressInfo
        Dim Address As MfgControl.AdvancedHMI.Drivers.ModbusTCP.ModbusAddress
        Dim dlgCallBack As IComComponent.ReturnValues
        Dim PollRate As Integer
        Dim ID As Integer
    End Structure


    Private TransactionNumber As New MfgControl.AdvancedHMI.Drivers.Common.TransactionNumber(0, 1024)

    '* Save the requests to reference when a response is received
    Private Transactions(255) As MfgControl.AdvancedHMI.Drivers.ModbusTCP.ModbusAddress

    '* Save the last 256 recieved packets
    Private SavedEventArgs(255) As MfgControl.AdvancedHMI.Drivers.Common.PlcComEventArgs
    Private SavedErrorEventArgs(255) As MfgControl.AdvancedHMI.Drivers.Common.PlcComErrorEventArgs


#Region "Properties"
    Private m_IPAddress As String = "0.0.0.0"   '* this is a default value
    <System.ComponentModel.Category("Communication Settings")> _
    Public Property IPAddress() As String
        Get
            Return m_IPAddress.ToString
        End Get
        Set(ByVal value As String)
            m_IPAddress = value

            '* If a new instance needs to be created, such as a different AMS Address
            CreateDLLInstance()


            If DLL(MyDLLInstance) IsNot Nothing Then
                DLL(MyDLLInstance).IPAddress = value
            End If
        End Set
    End Property

    Private m_TcpipPort As UInt16 = 502
    <System.ComponentModel.Category("Communication Settings")> _
    Public Property TcpipPort() As UInt16
        Get
            Return m_TcpipPort
        End Get
        Set(ByVal value As UInt16)
            m_TcpipPort = value

            '* If a new instance needs to be created, such as a different AMS Address
            CreateDLLInstance()


            If DLL(MyDLLInstance) IsNot Nothing Then
                DLL(MyDLLInstance).TcpipPort = value
            End If
        End Set
    End Property

    Private m_UnitId As Byte
    <System.ComponentModel.Category("Communication Settings")> _
    Public Property UnitId() As Byte
        Get
            Return m_UnitId
        End Get
        Set(ByVal value As Byte)
            m_UnitId = value
        End Set
    End Property

    '************************************************************
    '* If this is false, then wait for response before returning
    '* from read and writes
    '************************************************************
    Private m_AsyncMode As Boolean
    <System.ComponentModel.Category("Communication Settings")> _
    Public Property AsyncMode() As Boolean
        Get
            Return m_AsyncMode
        End Get
        Set(ByVal value As Boolean)
            m_AsyncMode = value
        End Set
    End Property

    Private _PollRateOverride As Integer
    <System.ComponentModel.Category("Communication Settings")> _
    Public Property PollRateOverride() As Integer
        Get
            Return _PollRateOverride
        End Get
        Set(ByVal value As Integer)
            _PollRateOverride = value
        End Set
    End Property

    '**************************************************
    '* Its purpose is to fetch
    '* the main form in order to synchronize the
    '* notification thread/event
    '**************************************************
    Private m_SynchronizingObject As System.ComponentModel.ISynchronizeInvoke
    '* do not let this property show up in the property window
    ' <System.ComponentModel.Browsable(False)> _
    Public Property SynchronizingObject() As System.ComponentModel.ISynchronizeInvoke
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

        Set(ByVal Value As System.ComponentModel.ISynchronizeInvoke)
            If Not Value Is Nothing Then
                m_SynchronizingObject = Value
            End If
        End Set
    End Property

    '*********************************************************************************
    '* Used to stop subscription updates when not needed to reduce communication load
    '*********************************************************************************
    Private m_DisableSubscriptions As Boolean
    Public Property DisableSubscriptions() As Boolean Implements IComComponent.DisableSubscriptions
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
    Public Sub New(ByVal container As System.ComponentModel.IContainer)
        MyClass.New()

        '* Default UnitID
        m_UnitId = 1

        'Required for Windows.Forms Class Composition Designer support
        container.Add(Me)
    End Sub

    Public Sub New()
        MyBase.New()

        '* Default UnitID
        m_UnitId = 1
    End Sub

    'Component overrides dispose to clean up the component list.
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        '* The handle linked to the DataLink Layer has to be removed, otherwise it causes a problem when a form is closed
        If DLL(MyDLLInstance) IsNot Nothing Then
            RemoveHandler DLL(MyDLLInstance).DataReceived, AddressOf DataLinkLayer_DataReceived
            RemoveHandler DLL(MyDLLInstance).ComError, AddressOf DataLinkLayer_ComError
        End If

        MyBase.Dispose(disposing)
    End Sub

    '***************************************************************
    '* Create the Data Link Layer Instances
    '* if the IP Address is the same, then resuse a common instance
    '***************************************************************
    Private Sub CreateDLLInstance()
        '* Still default, so ignore
        If m_IPAddress = "0.0.0.0" Then Exit Sub

        If DLL(0) IsNot Nothing Then
            '* At least one DLL instance already exists,
            '* so check to see if it has the same IP address
            '* if so, reuse the instance, otherwise create a new one
            Dim i As Integer
            While DLL(i) IsNot Nothing AndAlso DLL(i).IPAddress <> m_IPAddress AndAlso i < 11
                i += 1
            End While
            MyDLLInstance = i
        End If

        If DLL(MyDLLInstance) Is Nothing Then
            DLL(MyDLLInstance) = New MfgControl.AdvancedHMI.Drivers.ModbusTCP.ModbusTcpDataLinkLayer(m_IPAddress)
            AddHandler DLL(MyDLLInstance).DataReceived, AddressOf DataLinkLayer_DataReceived
            AddHandler DLL(MyDLLInstance).ComError, AddressOf DataLinkLayer_ComError
        End If
    End Sub
#End Region

#Region "Subscription"
    '*******************************************************************
    '*******************************************************************
    Private CurrentID As Integer
    Public Function Subscribe(ByVal plcAddress As String, ByVal numberOfElements As Int16, ByVal pollRate As Integer, ByVal callback As IComComponent.ReturnValues) As Integer Implements IComComponent.Subscribe
        '* If PollRateOverride is other than 0, use that poll rate for this subscription
        If _PollRateOverride > 0 Then
            pollRate = _PollRateOverride
        End If

        '* Avoid a 0 poll rate
        If pollRate <= 0 Then
            pollRate = 500
        End If

        '***********************************************
        '* Create an Address object address information
        '***********************************************
        Dim address As New MfgControl.AdvancedHMI.Drivers.ModbusTCP.ModbusAddress(plcAddress)

        '***********************************************************
        '* Check if there was already a subscription made for this
        '***********************************************************
        Dim index As Integer

        While index < PolledAddressList.Count AndAlso _
            (PolledAddressList(index).Address.Address <> plcAddress Or PolledAddressList(index).dlgCallBack <> callback)
            index += 1
        End While


        '* If a subscription was already found, then returns it's ID
        If (index < PolledAddressList.Count) Then
            '* Return the subscription that already exists
            Return PolledAddressList(index).ID
        Else
            '* The ID is used as a reference for removing polled addresses
            CurrentID += 1

            Dim tmpPA As PolledAddressInfo

            tmpPA.PollRate = pollRate
            tmpPA.dlgCallBack = callback
            tmpPA.ID = CurrentID
            tmpPA.Address = address

            '* Add this subscription to the collection and sort
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
                Dim tmrTemp As New System.Timers.Timer
                If pollRate > 0 Then
                    tmrTemp.Interval = pollRate
                Else
                    tmrTemp.Interval = 250
                End If

                tmrPollList.Add(tmrTemp)
                AddHandler tmrPollList(j).Elapsed, AddressOf PollUpdate

                tmrTemp.Enabled = True
            End If

            Return tmpPA.ID
        End If
    End Function

    '***************************************************************
    '* Used to sort polled addresses by File Number and element
    '* This helps in optimizing reading
    '**************************************************************
    Private Function SortPolledAddresses(ByVal A1 As PolledAddressInfo, ByVal A2 As PolledAddressInfo) As Integer
        If A1.Address.Address = A2.Address.Address Then
            If A1.Address.Address > A2.Address.Address Then
                Return 1
            ElseIf A1.Address.Address = A2.Address.Address Then
                Return 0
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
    Private Sub PollUpdate(ByVal sender As System.Object, ByVal e As System.Timers.ElapsedEventArgs)
        Dim intTimerIndex As Integer = tmrPollList.IndexOf(sender)
        '* Point a timer variable to the sender object for early binding
        Dim CurrentTimer As System.Timers.Timer = sender

        '* Stop the poll timer
        tmrPollList(intTimerIndex).Enabled = False
        Dim i As Integer
        While i < PolledAddressList.Count
            '* Is this firing timer at the requested poll rate
            If CurrentTimer.Interval = PolledAddressList(i).PollRate Then
                Dim value As String()
                Try
                    Dim SavedAsync As Boolean = m_AsyncMode
                    '*TODO : run subscriptions in Async Mode
                    m_AsyncMode = False
                    value = ReadAny(PolledAddressList(i).Address.Address, 1)
                    m_AsyncMode = SavedAsync
                    '* TODO : optimize in as few reads as possible - try group reading
                    '* and perform in Async Mode

                    m_SynchronizingObject.BeginInvoke(PolledAddressList(i).dlgCallBack, New Object() {value})
                Catch ex As Exception
                    RaiseEvent ComError(Me, New MfgControl.AdvancedHMI.Drivers.Common.PlcComErrorEventArgs(-1, ex.Message))
                    m_SynchronizingObject.BeginInvoke(PolledAddressList(i).dlgCallBack, New Object() {New String() {ex.Message}})
                End Try
            End If

            i += 1
        End While

        '* Start the poll timer
        tmrPollList(intTimerIndex).Enabled = True
    End Sub

    Public Function Unsubscribe(ByVal id As Integer) As Integer Implements IComComponent.Unsubscribe
        Dim i As Integer = 0
        While i < PolledAddressList.Count AndAlso PolledAddressList(i).ID <> id
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
#End Region

#Region "Public Methods"
    Public Function ReadSynchronous(ByVal startAddress As String, ByVal numberOfElements As Integer) As String() Implements IComComponent.ReadSynchronous
        Return ReadAny(startAddress, numberOfElements, False)
    End Function

    Public Function ReadAny(ByVal startAddress As String, ByVal numberOfElements As Integer) As String() Implements IComComponent.ReadAny
        Return ReadAny(startAddress, numberOfElements, m_AsyncMode)
    End Function

    Public Function ReadAny(ByVal startAddress As String) As String Implements IComComponent.ReadAny
        Return ReadAny(startAddress, 1)(0)
    End Function

    Public Function ReadAny(ByVal startAddress As String, ByVal numberOfElements As Integer, ByVal AsyncModeIn As Boolean) As String()
        Dim ma As New MfgControl.AdvancedHMI.Drivers.ModbusTCP.ModbusAddress(startAddress, numberOfElements)

        Dim TNS As UInt16 = TransactionNumber.GetNextNumber("m")

        '* Save the requested data information
        Transactions(TNS And 255) = ma


        Dim Header As New MfgControl.AdvancedHMI.Drivers.ModbusTCP.ModbusTCPHeaderPacket(TNS, m_UnitId)
        Dim Packet As New MfgControl.AdvancedHMI.Drivers.ModbusTCP.ModbusTCPPacket(Header, ma.ReadFunctionCode, ma.GetReadPacket)


        DLL(MyDLLInstance).SendData(Packet.GetByteStream)

        If AsyncModeIn Then
            Return New String() {"0"}
        Else
            '*************************
            '* Wait for response
            '*************************
            Dim TNSByte As Integer = TNS And 255
            Dim result As Integer = WaitForResponse(TNSByte)
            If result = 0 Then
                If Transactions(TNSByte).Responded Then
                    Dim tmp(SavedEventArgs(TNSByte).Values.Count - 1) As String
                    For i As Integer = 0 To tmp.Length - 1
                        tmp(i) = SavedEventArgs(TNSByte).Values(i)
                    Next
                    Return tmp
                Else
                    '* Reference : MODBUS Application Protocol Specification V1.1b, Page 49
                    Select Case SavedErrorEventArgs(TNSByte).ErrorId
                        Case 1
                            Throw New MfgControl.AdvancedHMI.Drivers.Common.PLCDriverException("Illegal Modbus Function")
                        Case 2
                            Throw New MfgControl.AdvancedHMI.Drivers.Common.PLCDriverException("Illegal Modbus Address")
                        Case 3
                            Throw New MfgControl.AdvancedHMI.Drivers.Common.PLCDriverException("Illegal Data Value")
                        Case 4
                            Throw New MfgControl.AdvancedHMI.Drivers.Common.PLCDriverException("Slave Device Failure")
                        Case 6
                            Throw New MfgControl.AdvancedHMI.Drivers.Common.PLCDriverException("Mobus Slave Device Busy")
                        Case Else
                            Throw New MfgControl.AdvancedHMI.Drivers.Common.PLCDriverException("Modbus Error Code " & SavedErrorEventArgs(TNS).ErrorId)
                    End Select
                End If
            Else
                Throw New MfgControl.AdvancedHMI.Drivers.Common.PLCDriverException("Com Fail - " & result)
            End If
        End If

    End Function

    Public Function WriteData(ByVal startAddress As String, ByVal dataToWrite As String) As String Implements IComComponent.WriteData
        Dim data() As String = {dataToWrite}
        Return WriteData(startAddress, data, 1)
    End Function

    Public Function WriteData(ByVal startAddress As String, ByVal dataToWrite() As String, ByVal numberOfElements As UInt16) As Integer
        Dim ma As New MfgControl.AdvancedHMI.Drivers.ModbusTCP.ModbusAddress(startAddress, numberOfElements)

        Dim TNS As UInt16 = TransactionNumber.GetNextNumber("o")
        '* Save the requested data information
        Transactions(TNS And 255) = ma

        Dim Header As New MfgControl.AdvancedHMI.Drivers.ModbusTCP.ModbusTCPHeaderPacket(TNS, m_UnitId)
        Dim WritePacket As New List(Of Byte)
        WritePacket.AddRange(ma.GetWritePacket)

        '* Attach the data
        If ma.BitsPerElement = 1 Then
            '* Encode bits in bytes
            Dim BitNumber As Integer
            Dim ByteNumber As Integer
            Dim element As Integer
            Dim data As Byte
            While element < numberOfElements
                If dataToWrite(element).ToUpper(System.Globalization.CultureInfo.CurrentCulture) = "TRUE" Or dataToWrite(element) > 0 Then
                    ByteNumber = Math.Floor(element / 8)
                    BitNumber = element - 8 * ByteNumber

                    data += 2 ^ BitNumber
                End If

                element += 1
                If BitNumber = 7 Or element >= numberOfElements Then
                    '* 28-SEP-12 Force a single coil use 255 instead of 1
                    If ma.WriteFunctionCode = 5 Then
                        If data > 0 Then
                            WritePacket.Add(255)
                        Else
                            WritePacket.Add(0)
                        End If
                        '* 27-JAN-13 Add the second byte that some require
                        WritePacket.Add(0)
                    Else
                        WritePacket.Add(data)
                    End If
                    data = 0
                End If
            End While
        Else
            '* put word values into big endian byte array
            For i As Integer = 0 To numberOfElements - 1
                Select Case ma.BitsPerElement
                    Case 1
                    Case 16
                        Dim x() As Byte = BitConverter.GetBytes(CShort(dataToWrite(i)))
                        WritePacket.Add(x(1))
                        WritePacket.Add(x(0))
                    Case 32
                        Dim x() As Byte = BitConverter.GetBytes(CInt(dataToWrite(i)))
                        WritePacket.Add(x(3))
                        WritePacket.Add(x(2))
                        WritePacket.Add(x(1))
                        WritePacket.Add(x(0))
                End Select
            Next
        End If

        Dim Packet As New MfgControl.AdvancedHMI.Drivers.ModbusTCP.ModbusTCPPacket(Header, ma.WriteFunctionCode, WritePacket.ToArray)

        '* Save the requested data information
        ma.IsWrite = True
        Transactions(TNS And 255) = ma

        Return DLL(MyDLLInstance).SendData(Packet.GetByteStream)
    End Function
#End Region

#Region "Private Methods"
    '****************************************************
    '* Wait for a response from PLC before returning
    '* Used for Synchronous communications
    '****************************************************
    Private MaxTicks As Integer = 350  '* 50 ticks per second
    Private Function WaitForResponse(ByVal ID As UInt16) As Integer
        Dim Loops As Integer = 0
        While Not Transactions(ID And 255).Responded And Not Transactions(ID And 255).ErrorReturned And Loops < MaxTicks
            System.Threading.Thread.Sleep(25)
            Loops += 1
        End While

        If Loops >= MaxTicks Then
            Return -20
        Else
            '* Only let the 1st time be a long delay
            MaxTicks = 75
            Return 0
        End If
    End Function


    '************************************************
    '* Process data recieved from controller
    '************************************************
    Private Sub DataLinkLayer_DataReceived(ByVal sender As Object, ByVal e As MfgControl.AdvancedHMI.Drivers.Common.PlcComEventArgs)
        Dim TNSByte As Integer = e.TransactionNumber And 255

        If Not Transactions(TNSByte).IsWrite Then
            If Transactions(TNSByte).BitsPerElement < 8 Then
                '*******************
                '* Extract the bits
                '*******************
                For i As Integer = 0 To Transactions(TNSByte).NumberOfElements - 1
                    Dim ByteNumber As Integer = Math.Floor(i / 8)
                    Dim BitNumber As Integer = i - 8 * ByteNumber
                    If (e.RawData(ByteNumber) And (2 ^ BitNumber)) > 0 Then
                        e.Values.Add("True")
                    Else
                        e.Values.Add("False")
                    End If
                Next
            Else
                '*************************
                '* Extract Integer values
                '*************************
                Dim BytesPerElement As Integer = Transactions(TNSByte).BitsPerElement / 8
                Dim NumberOfElements As Integer = Transactions(TNSByte).NumberOfElements
                For i As Integer = 0 To NumberOfElements - 1
                    Select Case BytesPerElement
                        Case 2
                            '* Reverse the Big Endian so the BitConverter works
                            Dim x() As Byte = {e.RawData(i * BytesPerElement + 1), e.RawData(i * BytesPerElement)}
                            e.Values.Add(BitConverter.ToInt16(x, 0))
                        Case 4
                            Dim x() As Byte = {e.RawData(i * BytesPerElement + 3), e.RawData(i * BytesPerElement + 2), e.RawData(i * BytesPerElement + 1), e.RawData(i * BytesPerElement)}
                            e.Values.Add(BitConverter.ToInt32(x, 0))
                    End Select
                Next
            End If
        End If

            '* Save this for other uses
        SavedEventArgs(TNSByte) = e
        Transactions(TNSByte).Responded = True

            '**************************************************
            '* Synchronize back to main calling thread if
            '* synchronizing object is defined
            '**************************************************
        If SynchronizingObject IsNot Nothing Then
            Dim Parameters() As Object = {Me, e}
            SynchronizingObject.BeginInvoke(drsd, Parameters)
        Else
            RaiseEvent DataReceived(Me, e)
        End If
    End Sub

    '************************************************
    '* Process error recieved from controller
    '************************************************
    'TODO: 
    Private Sub DataLinkLayer_ComError(ByVal sender As Object, ByVal e As MfgControl.AdvancedHMI.Drivers.Common.PlcComErrorEventArgs)
        '* Save this for other uses
        SavedErrorEventArgs(e.TransactionNumber And 255) = e

        If Transactions(e.TransactionNumber And 255) IsNot Nothing Then
            Transactions(e.TransactionNumber).ErrorReturned = True
        Else
            MsgBox(e.ErrorMessage)
        End If
    End Sub

    '***********************************************************
    '* Used to synchronize the event back to the calling thread
    '***********************************************************
    Dim drsd As EventHandler = AddressOf DataReceivedSync
    Private Sub DataReceivedSync(ByVal sender As Object, ByVal e As EventArgs)
        RaiseEvent DataReceived(sender, e)
    End Sub
#End Region
End Class