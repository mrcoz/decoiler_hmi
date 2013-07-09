'**********************************************************************************************
'* Ethernet/IP for ControlLogix
'*
'* Archie Jacobs
'* Manufacturing Automation, LLC
'* ajacobs@mfgcontrol.com
'* 14-DEC-10
'*
'*
'* Copyright 2010 Archie Jacobs
'*
'* This class implements the Ethernet/IP protocol.
'*
'* Distributed under the GNU General Public License (www.gnu.org)
'*
'* This program is free software; you can redistribute it and/or
'* as published by the Free Software Foundation; either version 2
'* of the License, or (at your option) any later version.
'*
'* This program is distributed in the hope that it will be useful,
'* but WITHOUT ANY WARRANTY; without even the implied warranty of
'* MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'* GNU General Public License for more details.

'* You should have received a copy of the GNU General Public License
'* along with this program; if not, write to the Free Software
'* Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301, USA.
'*
'* 23-MAY-12  Renamed PolledAddress* variables to Subscription* for clarity
'* 23-MAY-12  Created ComError event and check if subscription exists before sending error
'* 24-SEP-12  Added array reading optimization for subscriptions
'* 11-OCT-12  Do not optimize complex types, such as strings
'* 07-NOV-12  Array tags not sorting by element and caused number of elements to read to be wrong
'* 22-JAN-13  Added BasePLCAddress to Subscription info and used to sort arrays properly
'*******************************************************************************************************
Imports System.ComponentModel.Design

'<Assembly: system.Security.Permissions.SecurityPermissionAttribute(system.Security.Permissions.SecurityAction.RequestMinimum)> 
'<Assembly: CLSCompliant(True)> 
Public Class EthernetIPforCLXComm
    Inherits System.ComponentModel.Component
    Implements AdvancedHMIDrivers.IComComponent

    '* Create a common instance to share so multiple driver instances can be used in a project
    Private Shared DLL(10) As MfgControl.AdvancedHMI.Drivers.CIP
    Private MyDLLInstance As Integer


    Private DisableEvent As Boolean
    Private rnd As New Random

    Public Delegate Sub PLCCommEventHandler(ByVal sender As Object, ByVal e As MfgControl.AdvancedHMI.Drivers.Common.PlcComEventArgs)
    Public Event DataReceived As PLCCommEventHandler
    Public Event ComError As PLCComErrorEventHandler

    Public Event UnsolictedMessageRcvd As EventHandler

    Private Responded(255) As Boolean
    Private ActiveRequest As Boolean
    Private TNS1 As New MfgControl.AdvancedHMI.Drivers.Common.TransactionNumber(0, 32767)
    Private ReturnedInfo(255) As MfgControl.AdvancedHMI.Drivers.Common.PlcComEventArgs

    '*********************************************************************************
    '* This is used for linking a notification.
    '* An object can request a continuous poll and get a callback when value updated
    '*********************************************************************************
    Delegate Sub ReturnValues(ByVal Values As String)

    Friend Class SubscriptionInfo
        Public Sub New()
            PLCAddress = New MfgControl.AdvancedHMI.Drivers.CLXAddress
        End Sub

        Public PLCAddress As MfgControl.AdvancedHMI.Drivers.CLXAddress
        'Public FileType As Long
        'Public SubElement As Integer
        'Public BitNumber As Integer

        'Public BytesPerElement As Integer
        Public dlgCallBack As IComComponent.ReturnValues
        Public PollRate As Integer
        Public ID As Integer
        'Public LastValue As Object
        Public ElementsToRead As Integer
        Public SkipReads As Integer
        Public DataType As Byte
    End Class

    Private CurrentID As Integer = 1

    '* keep the original address by ref of low TNS byte so it can be returned to a linked polling address
    Private Shared PLCAddressByTNS(255) As CLXAddressRead
    Private SubscriptionList As New List(Of SubscriptionInfo)
    Private SubscriptionPollTimer As System.Timers.Timer

    Private Structure ParsedDataAddress
        Dim PLCAddress As String
        Dim SubElement As Integer
        Dim BitNumber As Integer
        Dim Element As Integer

        Dim SequenceNumber As Integer
        Dim InternallyRequested As Boolean
        Dim AsyncMode As Boolean

        'Dim BytesPerElements As Integer
        'Dim NumberOfElements As Integer
        'Dim FileType As Integer
    End Structure

#Region "Constructor"
    Private components As System.ComponentModel.IContainer
    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New(ByVal container As System.ComponentModel.IContainer)
        MyClass.New()

        'Required for Windows.Forms Class Composition Designer support
        container.Add(Me)
    End Sub

    Public Sub New()
        MyBase.New()

        'CreateDLLInstance()
        SubscriptionPollTimer = New System.Timers.Timer
        AddHandler SubscriptionPollTimer.Elapsed, AddressOf PollUpdate
        SubscriptionPollTimer.Interval = 500
    End Sub

    '***************************************************************
    '* Create the Data Link Layer Instances
    '* if the IP Address is the same, then resuse a common instance
    '***************************************************************
    Private Sub CreateDLLInstance()
        If DLL(0) IsNot Nothing Then
            '* At least one DLL instance already exists,
            '* so check to see if it has the same IP address
            '* if so, reuse the instance, otherwise create a new one
            Dim i As Integer
            While DLL(i) IsNot Nothing AndAlso DLL(i).EIPEncap.IPAddress <> m_IPAddress AndAlso i < 11
                i += 1
            End While
            MyDLLInstance = i
        End If

        If DLL(MyDLLInstance) Is Nothing Then
            DLL(MyDLLInstance) = New MfgControl.AdvancedHMI.Drivers.CIP
            DLL(MyDLLInstance).EIPEncap.IPAddress = m_IPAddress
            DLL(MyDLLInstance).ProcessorSlot = m_ProcessorSlot
        End If
        AddHandler DLL(MyDLLInstance).DataReceived, Dr
        AddHandler DLL(MyDLLInstance).CommError, ce
    End Sub

    'Private Dr As EventHandler = AddressOf DataLinkLayer_DataReceived
    'Private ce As PLCCommErrorEventHandler = AddressOf CommError
    Private ce As EventHandler = AddressOf CommError
    Public Delegate Sub PLCComErrorEventHandler(ByVal sender As Object, ByVal e As MfgControl.AdvancedHMI.Drivers.Common.PlcComErrorEventArgs)

    'Component overrides dispose to clean up the component list.
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        '* Stop the poll timers
        SubscriptionPollTimer.Enabled = False

        '* The handle linked to the DataLink Layer has to be removed, otherwise it causes a problem when a form is closed
        If DLL(MyDLLInstance) IsNot Nothing Then
            CloseConnection()
            RemoveHandler DLL(MyDLLInstance).DataReceived, Dr
            RemoveHandler DLL(MyDLLInstance).CommError, ce
        End If


        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub
#End Region

#Region "Properties"
    Private m_ProcessorSlot As Integer
    Public Property ProcessorSlot() As Integer
        Get
            Return m_ProcessorSlot
        End Get
        Set(ByVal value As Integer)
            m_ProcessorSlot = value
            If DLL(MyDLLInstance) IsNot Nothing Then
                DLL(MyDLLInstance).ProcessorSlot = value
            End If
        End Set
    End Property

    Private m_IPAddress As String = "192.168.0.10"
    Public Property IPAddress() As String
        Get
            'Return DLL(MyDLLInstance).EIPEncap.IPAddress
            Return m_IPAddress
        End Get
        Set(ByVal value As String)
            m_IPAddress = value

            '* If a new instance needs to be created, such as a different IP Address
            CreateDLLInstance()


            If DLL(MyDLLInstance) Is Nothing Then
            Else
                DLL(MyDLLInstance).EIPEncap.IPAddress = value
            End If
        End Set
    End Property

    '**************************************************************
    '* Determine whether to wait for a data read or raise an event
    '**************************************************************
    Private m_AsyncMode As Boolean
    Public Property AsyncMode() As Boolean
        Get
            Return m_AsyncMode
        End Get
        Set(ByVal value As Boolean)
            '* 22-MAY-13
            SyncLock (ReadLock)
                m_AsyncMode = value
            End SyncLock
        End Set
    End Property

    Private _PollRateOverride As Integer
    <System.ComponentModel.Category("Communication Settings")> _
    Public Property PollRateOverride() As Integer
        Get
            Return _PollRateOverride
        End Get
        Set(ByVal value As Integer)
            If value > 0 Then
                _PollRateOverride = value
                SubscriptionPollTimer.Interval = value
            End If
        End Set
    End Property

    '**************************************************************
    '* Stop the polling of subscribed data
    '**************************************************************
    Private m_DisableSubscriptions
    Public Property DisableSubscriptions() As Boolean Implements IComComponent.DisableSubscriptions
        Get
            Return m_DisableSubscriptions
        End Get
        Set(ByVal value As Boolean)
            m_DisableSubscriptions = value

            SubscriptionPollTimer.Enabled = Not value
        End Set
    End Property

    '**************************************************
    '* Its purpose is to fetch
    '* the main form in order to synchronize the
    '* notification thread/event
    '**************************************************
    Private m_SynchronizingObject As Object
    '* do not let this property show up in the property window
    ' <System.ComponentModel.Browsable(False)> _
    Public Property SynchronizingObject() As Object
        Get
            'If Me.Site.DesignMode Then

            Dim host1 As IDesignerHost
            Dim obj1 As Object
            If (m_SynchronizingObject Is Nothing) AndAlso MyBase.DesignMode Then
                host1 = CType(Me.GetService(GetType(IDesignerHost)), IDesignerHost)
                If host1 IsNot Nothing Then
                    obj1 = host1.RootComponent
                    m_SynchronizingObject = CType(obj1, System.ComponentModel.ISynchronizeInvoke)
                End If
            End If
            'End If
            Return m_SynchronizingObject



            'Dim dh As IDesignerHost = DirectCast(Me.GetService(GetType(IDesignerHost)), IDesignerHost)
            'If dh IsNot Nothing Then
            '    Dim obj As Object = dh.RootComponent
            '    If obj IsNot Nothing Then
            '        m_ParentForm = DirectCast(obj, Form)
            '    End If
            'End If

            'Dim instance As IDesignerHost = Me.GetService(GetType(IDesignerHost))
            'm_SynchronizingObject = instance.RootComponent
            ''End If
            'Return m_SynchronizingObject
        End Get

        Set(ByVal Value As Object)
            If Not Value Is Nothing Then
                m_SynchronizingObject = Value
            End If
        End Set
    End Property
#End Region

#Region "Public Methods"
    Public Function ReadSynchronous(ByVal startAddress As String, ByVal numberOfElements As Integer) As String() Implements IComComponent.ReadSynchronous
        Return ReadAny(startAddress, numberOfElements, False)
    End Function

    Public Function ReadAny(ByVal startAddress As String, ByVal numberOfElements As Integer) As String() Implements IComComponent.ReadAny
        Return ReadAny(startAddress, numberOfElements, m_AsyncMode)
    End Function


    Private ReadLock As New Object
    Public Function ReadAny(ByVal startAddress As String, ByVal numberOfElements As Integer, AsyncModeIn As Boolean) As String()
        '* Grab this value before it changes
        'Dim AsyncModeLocal As Boolean = m_AsyncMode

        SyncLock (ReadLock)
            ActiveRequest = True
            '* We must get the sequence number
            '* and save the read information because it can comeback before this
            '* information gets put in the PLCAddressByTNS array
            Dim SequenceNumber As Integer = TNS1.GetNextNumber("ReadAny1")

            PLCAddressByTNS(SequenceNumber And 255) = New CLXAddressRead
            PLCAddressByTNS(SequenceNumber And 255).AsyncMode = AsyncModeIn

            PLCAddressByTNS(SequenceNumber And 255).TagName = startAddress
            Responded(SequenceNumber And 255) = False
            PLCAddressByTNS(SequenceNumber And 255).TransactionNumber = SequenceNumber
            PLCAddressByTNS(SequenceNumber And 255).InternalRequest = InternalRequest

            'ActiveRequest = True
            DLL(MyDLLInstance).ReadTagValue(PLCAddressByTNS(SequenceNumber And 255), numberOfElements, SequenceNumber)
            'Threading.Thread.Sleep(10)

            If PLCAddressByTNS(SequenceNumber And 255).AsyncMode Then
                Dim x() As String = {SequenceNumber}
                Return x
            Else
                Dim result As Integer
                result = WaitForResponse(SequenceNumber And 255)
                If result = 0 Then
                    '* a Bit Array will return number of elements rounded up to 32, so return only the amount requested
                    If ReturnedInfo(SequenceNumber And 255).Values.Count > numberOfElements Then
                        Dim v(numberOfElements - 1) As String
                        For i As Integer = 0 To v.Length - 1
                            v(i) = ReturnedInfo(SequenceNumber And 255).Values(i)
                        Next i
                        Return v
                    Else
                        Return ReturnedInfo(SequenceNumber And 255).Values.ToArray
                    End If
                Else
                    Throw New MfgControl.AdvancedHMI.Drivers.Common.PLCDriverException("Read Failed - Status Code=" & result)
                End If
            End If
            'ActiveRequest = False
        End SyncLock
    End Function

    Public Function GetTagList() As MfgControl.AdvancedHMI.Drivers.CLXTag()

        '* We must get the sequence number from the DLL
        '* and save the read information because it can comeback before this
        '* information gets put in the PLCAddressByTNS array
        Dim SequenceNumber As Integer = TNS1.GetNextNumber("c")

        Responded(SequenceNumber And 255) = False

        PLCAddressByTNS(SequenceNumber And 255) = New AdvancedHMIDrivers.CLXAddressRead
        PLCAddressByTNS(SequenceNumber And 255).TransactionNumber = SequenceNumber
        PLCAddressByTNS(SequenceNumber And 255).InternalRequest = False 'InternalRequest
        'PLCAddressByTNS(SequenceNumber And 255).PLCAddress = startAddress
        PLCAddressByTNS(SequenceNumber And 255).AsyncMode = False 'm_AsyncMode


        Dim d() As MfgControl.AdvancedHMI.Drivers.CLXTag = DLL(MyDLLInstance).GetCLXTags()

        TNS1.ReleaseNumber(SequenceNumber)
        If PLCAddressByTNS(SequenceNumber And 255).AsyncMode Then
            Dim x() As String = {SequenceNumber}
            Return Nothing
        Else
            Return d
        End If
    End Function

    Public Sub CloseConnection()
        DLL(MyDLLInstance).ForwardClose()
    End Sub

    Public Function Subscribe(ByVal PLCAddress As String, ByVal numberOfElements As Int16, ByVal PollRate As Integer, ByVal CallBack As IComComponent.ReturnValues) As Integer Implements IComComponent.Subscribe
        If _PollRateOverride > 0 Then
            PollRate = _PollRateOverride
        End If

        '* Avoid a 0 poll rate
        If PollRate <= 0 Then
            PollRate = 500
        End If

        '* Valid address?
        'If ParsedResult.FileType <> 0 Then
        Dim tmpPA As New SubscriptionInfo
        tmpPA.PLCAddress.TagName = PLCAddress
        tmpPA.PollRate = PollRate
        tmpPA.dlgCallBack = CallBack
        tmpPA.ID = CurrentID
        tmpPA.ElementsToRead = 1


        SubscriptionList.Add(tmpPA)

        SubscriptionList.Sort(AddressOf SortPolledAddresses)

        '* The ID is used as a reference for removing polled addresses
        CurrentID += 1


        SubscriptionPollTimer.Enabled = True


        Return tmpPA.ID
    End Function

    '***************************************************************
    '* Used to sort polled addresses by File Number and element
    '* This helps in optimizing reading
    '**************************************************************
    Private Function SortPolledAddresses(ByVal A1 As SubscriptionInfo, ByVal A2 As SubscriptionInfo) As Integer
        '* Sort Tags

        '* 22-JAN-13
        '* Are they in the same array?
        If A1.PLCAddress.BaseArrayTag <> A2.PLCAddress.BaseArrayTag Then
            If A1.PLCAddress.BaseArrayTag > A2.PLCAddress.BaseArrayTag Then
                Return 1
            ElseIf A1.PLCAddress.BaseArrayTag < A2.PLCAddress.BaseArrayTag Then
                Return -1
            End If
        Else
            '* TODO : sort multidimensional array
            If A1.PLCAddress.ArrayIndex1 > A2.PLCAddress.ArrayIndex1 Then
                Return 1
            ElseIf A1.PLCAddress.ArrayIndex1 < A2.PLCAddress.ArrayIndex1 Then
                Return -1
            End If
        End If

        Return 0
    End Function

    Public Function UnSubscribe(ByVal ID As Integer) As Integer Implements IComComponent.Unsubscribe
        Dim i As Integer = 0
        While i < SubscriptionList.Count AndAlso SubscriptionList(i).ID <> ID
            i += 1
        End While

        If i < SubscriptionList.Count Then
            SubscriptionList.RemoveAt(i)
            '* if no more items to be polled, so remove all polling timers 28-NOV-10
            If SubscriptionList.Count <= 0 Then
                SubscriptionPollTimer.Enabled = False
            End If
        End If

        Return 0
    End Function

    '**************************************************************
    '* Perform the reads for the variables added for notification
    '* Attempt to optimize by grouping reads
    '**************************************************************
    Private InternalRequest As Boolean '* This is used to dinstinquish when to send data back to notification request
    Private Sub PollUpdate(ByVal sender As System.Timers.Timer, ByVal e As System.Timers.ElapsedEventArgs)
        If m_DisableSubscriptions Then Exit Sub
        '* 7-OCT-12 - Seemed like the timer would fire before the subscription was added
        If SubscriptionList Is Nothing OrElse SubscriptionList.Count <= 0 Then Exit Sub

        '* Stop the poll timer
        SubscriptionPollTimer.Enabled = False


        Dim i, NumberToRead, FirstElement As Integer
        While i < SubscriptionList.Count
            'Dim NumberToReadCalc As Integer
            NumberToRead = SubscriptionList(i).ElementsToRead
            FirstElement = i
            '            Dim PLCAddress As String = SubscriptionList(FirstElement).PLCAddress

            '* This eliminates the error of late binding when porting to Win CE
            'Dim tmr As System.Timers.Timer = sender

            Dim GroupedCount As Integer

            'If SubscriptionList(i).PollRate = sender.Interval Then
            GroupedCount = 1
            If SubscriptionList(i).SkipReads >= 0 Then
                '* Make sure it does not wait for return value befoe coming back
                Dim tmp As Boolean = Me.AsyncMode

                '* Optimize by reading array elements together - only single dimension array
                If SubscriptionList(FirstElement).PLCAddress.ArrayIndex1 >= 0 And SubscriptionList(FirstElement).PLCAddress.ArrayIndex2 < 0 And _
                    SubscriptionList(FirstElement).DataType > 0 Then
                    'SubscriptionList(FirstElement).DataType > 0 And SubscriptionList(FirstElement).DataType <> &HCE Then
                    Try
                        While FirstElement + GroupedCount < SubscriptionList.Count AndAlso _
                          SubscriptionList(FirstElement).PLCAddress.BaseArrayTag = _
                            SubscriptionList(FirstElement + GroupedCount).PLCAddress.BaseArrayTag

                            '* Add the number of span between the array
                            '* 07-NOV-12 - Sorting is Alphanumeric, so array do not sort properly
                            If NumberToRead < (SubscriptionList(FirstElement + GroupedCount).PLCAddress.ArrayIndex1 - SubscriptionList(FirstElement).PLCAddress.ArrayIndex1 + 1) Then
                                NumberToRead = SubscriptionList(FirstElement + GroupedCount).PLCAddress.ArrayIndex1 - SubscriptionList(FirstElement).PLCAddress.ArrayIndex1 + 1
                            End If
                            GroupedCount += 1
                        End While
                    Catch ex As Exception
                        Dim dbg = 0
                    End Try
                End If
                SyncLock (ReadLock) ' Do not let the asyncmode or Internal request change during another read
                    Me.AsyncMode = True
                    Try
                        InternalRequest = True
                        Dim SequenceNumber As UInt16 '= DLL(MyDLLInstance).EIPEncap.GetSequenceNumber

                        '23-OCT-11 Moved before the next 4 lines
                        SequenceNumber = Me.ReadAny(SubscriptionList(FirstElement).PLCAddress.TagName, NumberToRead)(0)

                        InternalRequest = False
                    Catch ex As Exception
                        '* Send this message back to the requesting control
                        Dim TempArray() As String = {ex.Message}

                        m_SynchronizingObject.BeginInvoke(SubscriptionList(i).dlgCallBack, New Object() {TempArray})
                        '* Slow down the poll rate to avoid app freezing
                        SubscriptionList(i).SkipReads = 10
                        'If SavedPollRate(0) = 0 Then SavedPollRate(0) = tmrPollList(intTimerIndex).Interval
                        'tmrPollList(intTimerIndex).Interval = 5000
                    Finally
                        '* Start the poll timer
                        'tmrPollList(intTimerIndex).Enabled = True

                        Me.AsyncMode = tmp
                    End Try
                End SyncLock
            Else
                If SubscriptionList(i).SkipReads > 0 Then SubscriptionList(i).SkipReads -= 1
            End If
            'End If
            i += GroupedCount
        End While

        '* Re-start the poll timer
        SubscriptionPollTimer.Enabled = True
    End Sub


    '********************************************************************
    '* Extract the data from the byte stream returned
    '* Use the abreviated type byte that is returned with data
    '********************************************************************
    Private Shared Function ExtractData(ByVal startAddress As String, ByVal AbreviatedType As Byte, ByVal ReturnedData() As Byte) As String()
        '* Get the element size in bytes
        Dim ElementSize As Integer
        Select Case AbreviatedType
            Case &HC1 '* BIT
                ElementSize = 1
            Case &HC2 '* SINT
                ElementSize = 1
            Case &HC3 ' * INT
                ElementSize = 2
            Case &HC4, &HCA '* DINT, REAL Value read (&H91)
                ElementSize = 4
            Case &HD3 '* Bit Array
                ElementSize = 4
            Case &H82, &H83 ' * Timer, Counter, Control
                ElementSize = 14
            Case &HCE '* String
                'ElementSize = ReturnedData(0) + ReturnedData(1) * 256
                ElementSize = 88
            Case Else
                ElementSize = 2
        End Select

        Dim BitsPerElement As Integer = ElementSize * 8
        '***************************************************
        '* Extract returned data into appropriate data type
        '***************************************************
        Dim ElementsToReturn As UInt16 = Math.Floor(ReturnedData.Length / ElementSize) - 1
        '* Bit Arrays are return as DINT, so it will have to be extracted
        Dim BitIndex As UInt16
        If AbreviatedType = &HD3 Then
            If startAddress.LastIndexOf("[") > 0 Then
                BitIndex = startAddress.Substring(startAddress.LastIndexOf("[") + 1, startAddress.LastIndexOf("]") - startAddress.LastIndexOf("[") - 1)
            End If
            BitIndex -= Math.Floor(BitIndex / 32) * 32
            '* Return all the bits that came back even if less were requested
            ElementsToReturn = ReturnedData.Length * 8 - BitIndex - 1
            BitsPerElement = 32
        End If

        '* 18-MAY-12
        '* Check if it is addressing a single bit in a larger data type
        Dim PeriodPos As Integer = startAddress.IndexOf(".")
        If PeriodPos > 0 Then
            Dim SubElement As String = startAddress.Substring(PeriodPos + 1)

            Try
                If Integer.TryParse(SubElement, BitIndex) Then
                    'BitIndex = CInt(SubElement)

                    Select Case AbreviatedType
                        Case &HC3 '* INT Value read 
                            BitsPerElement = 16
                        Case &HC4 '* DINT Value read (&H91)
                            BitsPerElement = 32
                        Case &HC2 '* SINT
                            BitsPerElement = 8
                    End Select

                    If BitIndex > 0 And BitIndex < BitsPerElement Then
                        '* Force it to appear like a bit array
                        AbreviatedType = &HD3
                        BitIndex -= Math.Floor(BitIndex / BitsPerElement) * BitsPerElement
                        '* Return all the bits that came back even if less were requested
                        ElementsToReturn = ReturnedData.Length * 8 - BitIndex - 1
                    End If
                End If
            Catch ex As Exception
                '* If the value can not be converted, then it is not a valid integer
            End Try
        End If


        Dim result(ElementsToReturn) As String

        '* If requesting 0 elements, then default to 1
        'Dim ArrayElements As Int16 = Math.Max(result.Length - 1 - 1, 0)


        Select Case AbreviatedType
            Case &HC1 '* BIT
                For i As Integer = 0 To result.Length - 1
                    If ReturnedData(i) Then
                        result(i) = True
                    Else
                        result(i) = False
                    End If
                Next
            Case &HCA '* REAL read (&H8A)
                For i As Integer = 0 To result.Length - 1
                    result(i) = BitConverter.ToSingle(ReturnedData, (i * 4))
                Next
            Case &HC3 '* INT Value read 
                For i As Integer = 0 To result.Length - 1
                    result(i) = BitConverter.ToInt16(ReturnedData, (i * 2))
                Next
            Case &HC4 '* DINT Value read (&H91)
                For i As Integer = 0 To result.Length - 1
                    result(i) = BitConverter.ToInt32(ReturnedData, (i * 4))
                Next
            Case &HC2 '* SINT
                For i As Integer = 0 To result.Length - 1
                    result(i) = ReturnedData(i)
                Next
            Case &HD3 '* BOOL Array
                Dim x, i, l As UInt32
                Dim CurrentBitPos As UInt16 = BitIndex
                For j As Integer = 0 To (ReturnedData.Length / 4) - 1
                    x = BitConverter.ToUInt32(ReturnedData, (j * 4))
                    While CurrentBitPos < BitsPerElement
                        l = (2 ^ (CurrentBitPos))
                        result(i) = (x And l) > 0
                        i += 1
                        CurrentBitPos += 1
                    End While
                    CurrentBitPos = 0
                Next
            Case &H82, &H83 '* TODO: Timer, Counter, Control 
                Dim StartByte As Int16 = 2
                '                Dim x = startAddress
                '* Look for which sub element is specificed
                If startAddress.IndexOf(".") >= 0 Then
                    If startAddress.ToUpper.IndexOf("PRE") > 0 Then
                        StartByte = 6
                    ElseIf startAddress.ToUpper.IndexOf("ACC") > 0 Then
                        StartByte = 10
                    End If
                Else
                    '* If no subelement, then use ACC
                End If

                For i As Integer = 0 To result.Length - 1
                    result(i) = BitConverter.ToInt32(ReturnedData, (i + StartByte))
                Next
            Case &HCE ' * String
                For i As Integer = 0 To result.Length - 1
                    result(i) = System.Text.Encoding.ASCII.GetString(ReturnedData, 88 * i + 4, ReturnedData(88 * i) + ReturnedData(88 * i + 1) * 256)
                Next
            Case Else
                For i As Integer = 0 To result.Length - 1
                    result(i) = BitConverter.ToInt16(ReturnedData, (i * 2))
                Next
        End Select


        Return result
    End Function
    '*************************************************************
    '* Overloaded method of ReadAny - that reads only one element
    '*************************************************************
    ''' <summary>
    ''' Synchronous read of any data type
    ''' this function returns results as a string
    ''' </summary>
    ''' <param name="startAddress"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ReadAny(ByVal startAddress As String) As String Implements IComComponent.ReadAny
        Return ReadAny(startAddress, 1)(0)
    End Function



    '*****************************************************************
    '* Write Section
    '*
    '* Address is in the form of <file type><file Number>:<offset>
    '* examples  N7:0, B3:0,
    '******************************************************************

    '* Handle one value of Integer type
    ''' <summary>
    ''' Write a single integer value to a PLC data table
    ''' The startAddress is in the common form of AB addressing (e.g. N7:0)
    ''' </summary>
    ''' <param name="startAddress"></param>
    ''' <param name="dataToWrite"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function WriteData(ByVal startAddress As String, ByVal dataToWrite As Integer) As Integer
        Dim temp(1) As String
        temp(0) = CStr(dataToWrite)
        Return WriteData(startAddress, 1, temp)
    End Function


    '* Write an array of integers
    ''' <summary>
    ''' Write multiple consectutive integer values to a PLC data table
    ''' The startAddress is in the common form of AB addressing (e.g. N7:0)
    ''' </summary>
    ''' <param name="startAddress"></param>
    ''' <param name="numberOfElements"></param>
    ''' <param name="dataToWrite"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function WriteData(ByVal startAddress As String, ByVal numberOfElements As Integer, ByVal dataToWrite() As String) As Integer
        Dim Delays As Integer
        While ActiveRequest And Delays < 50
            System.Threading.Thread.Sleep(10)
            Delays += 1
        End While
        'If ActiveRequest Then
        '    Dim x = 0
        'End If

        Dim StringVals(numberOfElements) As String
        For i As Integer = 0 To numberOfElements - 1
            StringVals(i) = CStr(dataToWrite(i))
        Next
        Dim SequenceNumber As Integer = TNS1.GetNextNumber("WD")
        Dim tag As New CLXAddressRead
        tag.TagName = startAddress
        SyncLock (ReadLock)
            DLL(MyDLLInstance).WriteTagValue(tag, StringVals, numberOfElements, SequenceNumber)
        End SyncLock
    End Function

    '* Handle one value of Single type
    ''' <summary>
    ''' Write a single floating point value to a data table
    ''' The startAddress is in the common form of AB addressing (e.g. F8:0)
    ''' </summary>
    ''' <param name="startAddress"></param>
    ''' <param name="dataToWrite"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function WriteData(ByVal startAddress As String, ByVal dataToWrite As Single) As Integer
        Dim temp(1) As Single
        temp(0) = dataToWrite
        Return WriteData(startAddress, 1, temp)
    End Function

    '* Write an array of Singles
    ''' <summary>
    ''' Write multiple consectutive floating point values to a PLC data table
    ''' The startAddress is in the common form of AB addressing (e.g. F8:0)
    ''' </summary>
    ''' <param name="startAddress"></param>
    ''' <param name="numberOfElements"></param>
    ''' <param name="dataToWrite"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function WriteData(ByVal startAddress As String, ByVal numberOfElements As Integer, ByVal dataToWrite() As Single) As Integer
        Dim StringVals(numberOfElements) As String
        For i As Integer = 0 To numberOfElements - 1
            StringVals(i) = CStr(dataToWrite(i))
        Next

        WriteData(startAddress, numberOfElements, StringVals)
    End Function

    Public Function WriteData(ByVal startAddress As String, ByVal dataToWrite() As Single) As Integer
        Dim StringVals(dataToWrite.Length - 1) As String
        For i As Integer = 0 To StringVals.Length - 1
            StringVals(i) = CStr(dataToWrite(i))
        Next

        WriteData(startAddress, dataToWrite.Length, StringVals)
    End Function


    Public Function WriteData(ByVal startAddress As String, ByVal numberOfElements As Integer, ByVal dataToWrite() As Integer) As Integer
        Dim StringVals(numberOfElements) As String
        For i As Integer = 0 To numberOfElements - 1
            StringVals(i) = CStr(dataToWrite(i))
        Next

        WriteData(startAddress, numberOfElements, StringVals)
    End Function

    Public Function WriteData(ByVal startAddress As String, ByVal dataToWrite() As Integer) As Integer
        Dim StringVals(dataToWrite.Length - 1) As String
        For i As Integer = 0 To StringVals.Length - 1
            StringVals(i) = CStr(dataToWrite(i))
        Next

        WriteData(startAddress, dataToWrite.Length, StringVals)
    End Function


    '* Write a String
    ''' <summary>
    ''' Write a string value to a string data table
    ''' The startAddress is in the common form of AB addressing (e.g. ST9:0)
    ''' </summary>
    ''' <param name="startAddress"></param>
    ''' <param name="dataToWrite"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function WriteData(ByVal startAddress As String, ByVal dataToWrite As String) As String Implements IComComponent.WriteData
        If dataToWrite Is Nothing Then
            Return 0
        End If

        'WriteData(startAddress, 1, dataToWrite)

        Dim Delays As Integer
        While ActiveRequest And Delays < 10
            System.Threading.Thread.Sleep(10)
            Delays += 1
        End While

        Dim StringVals() As String = {dataToWrite}
        Dim TNS As Integer = TNS1.GetNextNumber("w")
        Dim tag As New CLXAddressRead
        tag.TagName = startAddress
        SyncLock (ReadLock)
            DLL(MyDLLInstance).WriteTagValue(tag, StringVals, 1, TNS)
        End SyncLock
        Return 0
    End Function

    'End of Public Methods
#End Region

#Region "Shared Methods"
    '****************************************************************
    '* Convert an array of words into a string as AB PLC's represent
    '* Can be used when reading a string from an Integer file
    '****************************************************************
    ''' <summary>
    ''' Convert an array of integers to a string
    ''' This is used when storing strings in an integer data table
    ''' </summary>
    ''' <param name="words"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function WordsToString(ByVal words() As Int32) As String
        Dim WordCount As Integer = words.Length
        Return WordsToString(words, 0, WordCount)
    End Function

    ''' <summary>
    ''' Convert an array of integers to a string
    ''' This is used when storing strings in an integer data table
    ''' </summary>
    ''' <param name="words"></param>
    ''' <param name="index"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function WordsToString(ByVal words() As Int32, ByVal index As Integer) As String
        Dim WordCount As Integer = (words.Length - index)
        Return WordsToString(words, index, WordCount)
    End Function

    ''' <summary>
    ''' Convert an array of integers to a string
    ''' This is used when storing strings in an integer data table
    ''' </summary>
    ''' <param name="words"></param>
    ''' <param name="index"></param>
    ''' <param name="wordCount"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function WordsToString(ByVal words() As Int32, ByVal index As Integer, ByVal wordCount As Integer) As String
        Dim j As Integer = index
        Dim result2 As New System.Text.StringBuilder
        While j < wordCount
            result2.Append(Chr(words(j) / 256))
            '* Prevent an odd length string from getting a Null added on
            If CInt(words(j) And &HFF) > 0 Then
                result2.Append(Chr(words(j) And &HFF))
            End If
            j += 1
        End While

        Return result2.ToString
    End Function


    '**********************************************************
    '* Convert a string to an array of words
    '*  Can be used when writing a string to an Integer file
    '**********************************************************
    ''' <summary>
    ''' Convert a string to an array of words
    ''' Can be used when writing a string into an integer data table
    ''' </summary>
    ''' <param name="source"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function StringToWords(ByVal source As String) As Int32()
        If source Is Nothing Then
            Return Nothing
            ' Throw New ArgumentNullException("input")
        End If

        Dim ArraySize As Integer = CInt(Math.Ceiling(source.Length / 2)) - 1

        Dim ConvertedData(ArraySize) As Int32

        Dim i As Integer
        While i <= ArraySize
            ConvertedData(i) = Asc(source.Substring(i * 2, 1)) * 256
            '* Check if past last character of odd length string
            If (i * 2) + 1 < source.Length Then ConvertedData(i) += Asc(source.Substring((i * 2) + 1, 1))
            i += 1
        End While

        Return ConvertedData
    End Function

#End Region

#Region "Helper"
    '****************************************************
    '* Wait for a response from PLC before returning
    '****************************************************
    Dim MaxTicks As Integer = 100  '* 100 ticks per second
    Private Function WaitForResponse(ByVal rTNS As Integer) As Integer
        Dim Loops As Integer = 0
        While Not Responded(rTNS) And Loops < MaxTicks
            System.Threading.Thread.Sleep(10)
            Loops += 1
        End While

        If Loops >= MaxTicks Then
            Return -20
        Else
            Return 0
        End If
    End Function

    '**************************************************************
    '* This method implements the common application routine
    '* as discussed in the Software Layer section of the AB manual
    '**************************************************************
    '**************************************************************
    '* This method Sends a response from an unsolicited msg
    '**************************************************************
    Private Function SendResponse(ByVal Command As Byte, ByVal rTNS As Integer) As Integer
    End Function

    ' TODO : Put this in a New event
    'Private EventHandleAdded As Boolean
    '* This is needed so the handler can be removed
    Private Dr As EventHandler = AddressOf DataLinkLayer_DataReceived

    '************************************************
    '* Convert the message code number into a string
    '* Ref Page 8-3
    '************************************************
    Public Shared Function DecodeMessage(ByVal msgNumber As Integer) As String
        Select Case msgNumber
            Case 0
                DecodeMessage = ""
            Case -4
                Return "Unknown Message from DataLink Layer"
            Case -5
                Return "Invalid Address"
            Case -7
                Return "No data specified to data link layer"
            Case -8
                Return "No data returned from PLC"
            Case -20
                Return "No Data Returned"

                '*** Errors coming from PLC
            Case 4
                Return "Invalid Tag Address."
            Case 5
                Return "The particular item referenced (usually instance) could not be found"
            Case &HA
                Return "An error has occurred trying to process one of the attributes"
            Case &H13
                Return "Not enough command data / parameters were supplied in the command to execute the service requested"
            Case &H1C
                Return "An insufficient number of attributes were provided compared to the attribute count"
            Case &H26
                Return "The IOI word length did not match the amount of IOI which was processed"
            Case 32
                Return "PLC Has a Problem and Will Not Communicate"

                '* EXT STS Section - 256 is added to code to distinguish EXT codes
            Case 257
                Return "A field has an illegal value"
            Case 258
                Return "Less levels specified in address than minimum for any address"
            Case 270
                Return "Command cannot be executed"
                '* Extended CIP codes - Page 13 Logix5000 Data Access
            Case &H2105
                Return "You have tried to access beyond the end of the data object"
            Case &H2107
                Return "The abbreviated type does not match the data type of the data object."
            Case &H2104
                Return "The beginning offset was beyond the end of the template"
            Case Else
                Return "Unknown Message - " & msgNumber
        End Select
    End Function

    '***************************************************************************************
    '* If an error comes back from the driver, return the description back to the control
    '***************************************************************************************
    Private Sub CommError(ByVal sender As Object, ByVal e As MfgControl.AdvancedHMI.Drivers.Common.PlcComErrorEventArgs)
        Dim d() As String = {DecodeMessage(e.ErrorId * -1)}
        'Dim x As New MfgControl.AdvancedHMI.Drivers.PLCCommEventArgs(d, Nothing, Nothing)
        If m_SynchronizingObject IsNot Nothing Then
            '* If a subscription exists, then return error to subscription
            If SubscriptionList.Count > 0 Then
                If SubscriptionList(0).dlgCallBack Is Nothing Then
                    Dim Parameters() As Object = {Me, e}
                    m_SynchronizingObject.BeginInvoke(drsd, Parameters)
                Else
                    m_SynchronizingObject.BeginInvoke(SubscriptionList(0).dlgCallBack, New Object() {d})
                End If
            End If
        End If

        RaiseEvent ComError(Me, e)
    End Sub

    Private Sub DataLinkLayer_DataReceived(ByVal sender As Object, ByVal e As MfgControl.AdvancedHMI.Drivers.Common.PlcComEventArgs)
        '* Was this TNS requested by this instance
        If TNS1.IsMyTNS(e.TransactionNumber) Then
            TNS1.ReleaseNumber(e.TransactionNumber)
        Else
            Dim x = 0
            Exit Sub
        End If

        ActiveRequest = False


        '**************************************************************************
        '* If the parent form property (Synchronizing Object) is set, then sync the event with its thread
        '**************************************************************************
        '* Get the low byte from the Sequence Number
        '* The sequence number was added onto the end of the CIP packet by the EthernetIPLayer object
        'Dim SequenceNumber As Integer = DLL(MyDLLInstance).DataPacket(SequenceNumber)(DLL(MyDLLInstance).DataPacket(SequenceNumber).Count - 2)
        'DataPackets(e.SequenceNumber) = DLL(MyDLLInstance).DataPacket(e.SequenceNumber)

        ' Responded(e.TransactionNumber And 255) = True

        ' If PLCAddressByTNS(e.TransactionNumber And 255).AsyncMode Then
        '* Check the status byte
        Dim StatusByte As Int16 = e.RawData(2)
        '* Extended status code, Page 13 of Logix5000 Data Access
        If StatusByte = &HFF And e.RawData.Count >= 5 Then
            StatusByte = e.RawData(5) * 255 + e.RawData(4)
        End If

        If StatusByte = 0 Then
            '**************************************************************
            '* Only extract and send back if this response contained data
            '**************************************************************
            If e.RawData.Count > 5 Then
                '***************************************************
                '* Extract returned data into appropriate data type
                '* Transfer block of data read to the data table array
                '***************************************************
                '* TODO: Check array bounds
                Try
                    Dim DataType As Byte = e.RawData(4)
                    Dim DataStartIndex As UInt16 = 6
                    '* Is it a complex data type
                    If DataType = &HA0 Then
                        DataType = e.RawData(6)
                        DataStartIndex = 8
                    End If
                    Dim ReturnedData(e.RawData.Count - DataStartIndex - 1) As Byte
                    For i As Integer = 0 To ReturnedData.Length - 1
                        ReturnedData(i) = e.RawData(i + DataStartIndex)
                    Next

                    '* Pass the abreviated data type (page 11 of 1756-RM005A)
                    Dim d() As String = ExtractData(PLCAddressByTNS(e.TransactionNumber And 255).TagName, DataType, ReturnedData)
                    ReturnedInfo(e.TransactionNumber And 255) = New MfgControl.AdvancedHMI.Drivers.Common.PlcComEventArgs(d, PLCAddressByTNS(e.TransactionNumber And 255).TagName, e.TransactionNumber And 255)
                    Responded(e.TransactionNumber And 255) = True


                    If Not PLCAddressByTNS(e.TransactionNumber And 255).InternalRequest Then
                        If Not DisableEvent Then
                            If m_SynchronizingObject IsNot Nothing Then
                                Dim Parameters() As Object = {Me, ReturnedInfo(e.TransactionNumber And 255)}
                                m_SynchronizingObject.BeginInvoke(drsd, Parameters)
                            Else
                                RaiseEvent DataReceived(Me, ReturnedInfo(e.TransactionNumber And 255))
                            End If
                        End If
                    Else
                        '*********************************************************
                        '* Check to see if this is from the Polled variable list
                        '*********************************************************
                        For i As Integer = 0 To SubscriptionList.Count - 1
                            '****************************************************************
                            '* Get the address without array element 07-SEP-12
                            '* this is for reads that group array elements into single read
                            '*****************************************************************
                            'Dim BaseArrayAddress As String = PLCAddressByTNS(e.TransactionNumber And 255).BaseArrayTag
                            'If PLCAddressByTNS(e.TransactionNumber And 255).TagName.IndexOf("[") > 0 Then
                            '    BaseArrayAddress = PLCAddressByTNS(e.TransactionNumber And 255).PLCAddress.Substring(0, PLCAddressByTNS(e.TransactionNumber And 255).PLCAddress.IndexOf("["))
                            'End If


                            Dim MaxElementNeeded As Integer = SubscriptionList(i).PLCAddress.ArrayIndex1 + SubscriptionList(i).ElementsToRead
                            Dim MaxElementRead As Integer = PLCAddressByTNS(e.TransactionNumber And 255).ArrayIndex1 + d.Length
                            Dim StartElementRead As Integer = PLCAddressByTNS(e.TransactionNumber And 255).ArrayIndex1
                            Dim StartElementNeeded As Integer = SubscriptionList(i).PLCAddress.ArrayIndex1
                            If (SubscriptionList(i).PLCAddress.BaseArrayTag = PLCAddressByTNS(e.TransactionNumber And 255).BaseArrayTag And ((StartElementRead <= StartElementNeeded And MaxElementNeeded <= MaxElementRead) Or StartElementNeeded < 0)) Then
                                Dim BitResult(SubscriptionList(i).ElementsToRead - 1) As String

                                '* All other data types
                                For k As Integer = 0 To SubscriptionList(i).ElementsToRead - 1
                                    '* a -1 in ArrayElement number means it is not an array
                                    If SubscriptionList(i).PLCAddress.ArrayIndex1 >= 0 Then
                                        BitResult(k) = d((SubscriptionList(i).PLCAddress.ArrayIndex1 - PLCAddressByTNS(e.TransactionNumber And 255).ArrayIndex1 + k))
                                    Else
                                        BitResult(k) = d(k)
                                    End If
                                Next

                                '* 23-APR-13 Did we read an array of integers, but the subscribed element was a bit?
                                If PLCAddressByTNS(e.TransactionNumber And 255).BitNumber < 0 And SubscriptionList(i).PLCAddress.BitNumber >= 0 Then
                                    BitResult(0) = (((2 ^ SubscriptionList(i).PLCAddress.BitNumber) And BitResult(0)) > 0)
                                End If

                                SubscriptionList(i).DataType = DataType

                                m_SynchronizingObject.BeginInvoke(SubscriptionList(i).dlgCallBack, New Object() {BitResult})
                                'Dim x As New MfgControl.AdvancedHMI.Drivers.PLCCommEventArgs(BitResult, Nothing, Nothing, PolledAddressList(i).dlgCallBack)
                                'm_SynchronizingObject.BeginInvoke(drsd, x)
                            End If
                        Next
                    End If

                Catch ex As Exception
                    Dim dbg = 0
                End Try
            End If
        Else
            '*****************************
            '* Failed Status was returned
            '*****************************
            Dim d() As String = {DecodeMessage(StatusByte) & " CIP Status " & StatusByte}
            If Not PLCAddressByTNS(e.TransactionNumber And 255).InternalRequest Then
                If Not DisableEvent Then
                    Dim x As New MfgControl.AdvancedHMI.Drivers.Common.PlcComEventArgs(d, PLCAddressByTNS(e.TransactionNumber And 255).TagName, e.TransactionNumber And 255)
                    If m_SynchronizingObject IsNot Nothing Then
                        Dim Parameters() As Object = {Me, x}
                        m_SynchronizingObject.BeginInvoke(drsd, Parameters)
                    Else
                        'TODO : This should raise the Com Error event
                        RaiseEvent DataReceived(Me, x)
                    End If
                End If
            Else
                For i As Integer = 0 To SubscriptionList.Count - 1
                    If SubscriptionList(i).PLCAddress.TagName = PLCAddressByTNS(e.TransactionNumber And 255).TagName Then
                        m_SynchronizingObject.BeginInvoke(SubscriptionList(i).dlgCallBack, New Object() {d})
                    End If
                Next
            End If
        End If
        'End If
        'mut.ReleaseMutex()
    End Sub

    '******************************************************************
    '* This is called when a message instruction was sent from the PLC
    '******************************************************************
    Private Sub DF1DataLink1_UnsolictedMessageRcvd()
        If m_SynchronizingObject IsNot Nothing Then
            Dim Parameters() As Object = {Me, EventArgs.Empty}
            m_SynchronizingObject.BeginInvoke(drsd, Parameters)
        Else
            RaiseEvent UnsolictedMessageRcvd(Me, System.EventArgs.Empty)
        End If
    End Sub


    '****************************************************************************
    '* This is required to sync the event back to the parent form's main thread
    '****************************************************************************
    Dim drsd As EventHandler = AddressOf DataReceivedSync
    'Delegate Sub DataReceivedSyncDel(ByVal sender As Object, ByVal e As EventArgs)
    Private Sub DataReceivedSync(ByVal sender As Object, ByVal e As MfgControl.AdvancedHMI.Drivers.Common.PlcComEventArgs)
        RaiseEvent DataReceived(Me, e)
    End Sub

    Private Sub UnsolictedMessageRcvdSync(ByVal sender As Object, ByVal e As EventArgs)
        RaiseEvent UnsolictedMessageRcvd(sender, e)
    End Sub
#End Region

End Class
