'SNP-X serial port driver.  Based on GEFCOMMSerial written by Doug MacLeod at GE Fanuc Tech
'Support.  A total rewrite using User's Manual GFK-0582C "Series 90 PLC Serial 'Communications" 
'while using GEFCOMM as a baseline reference.  End target is driver for Advanced HMI written 
'by Archie Jacobs and available on SourceForge. 

'********************************************************************************************************************
'From Designer enter addresses in standard GE format e.g.,Q1, I1, M100 ... for various Control's PLCAdress properties
'********************************************************************************************************************

'All discrete types ie., I, Q, G, M, T, and S are by default handled as Booleans.  Adding " %B" to the address string
'results in accessing the PLC memory address as a byte instead of by the bit. e.g., "M501 %B".  Keep in mind that the PLC
'address passed will be adjusted for the byte where it is located so in the above example byte address 62 is referenced with
'"M501" being the 4th bit in that byte.  Subscriptions are always read at bit level.  When passing byte data as a string 
'pass in hex format "00" to "FF". When calling WriteData(...) and passing data as an integer or array of integers, as with 
'multiple element writes, for bit level writes data should only be equal to 0 or 1 and at byte level equal from 0 to 255.
'Out of range data values will result in a -30 return value

'SNP-X works at the 16 bit word level so types R, AI and AQ are by default 16 bit signed integers.
'These types can be overridden by adding " %L" or " %R" to the address string causing the read/write subs to treat that
'memory locations as a 32 bit signed integer or 32 bit floating point.  The memory location immediately following the address
'passed will also be used. eg., "R500 %L" read/write 16 ls bits in R500 and the 16 ms bits in R501.

'There is no addressing range checking so an invalid PLC subscribe read request and most other types of
'errors will result with an exception being thrown. An invalid non-subscribed read will result in a zero value
'being returned and Global CurrentSNPXRequestInError being set. Change ReadAny(ByVal startAddress As String)
'commenting out the "if case" and uncomment the "throw" to always catch invalid memory addresses.  Code developed 
'with a multi-drop system in mind but I don't have such a system available for testing.  The SNPid and Attach 
'properties will require further development for such a system.

'There are three globals that can be used for monitoring status of PLC - CurrentSNPXRequestInError, 
'LastSNPXErrorCode & CurrentSNPXStatusCode. Both error and status codes come with accompanying overloaded decode fuctions.

'01/04/2012 Corrected bug where SNPX message with Byte Check Code equal to 0 was being ignored.

'12/29/2011 Added " %B" suffix to discrete memory type to allow for byte level writes and reads.  Subscriptions always
'           read at the Bit level regardless if modifier is present.

'12/19/2011 Added " %L" and " %R" suffix to memory types R, AI and AQ to allow for 32 bit signed integers and floating points
'           resulting in a rewrite of SNPX_DataLinkLayer_DataReceived() and ExtractData()

'11/2/2011  Was not reading or writing negative numbers correctly. Currently the only data types are Boolean and signed
'           16 bit Integers

'11/1/2011  Added Try/Catch to SendSerialBreak because an exception was being triggered when computer waking up from sleep
'           also changed property Attached to readonly since it should only be updated internally - user can monitor this
'           from Form using it as a warning trigger for when communication is lost with PLC



Imports System.ComponentModel.Design
Imports System.Text.RegularExpressions
Imports System.ComponentModel
Imports System
Imports System.Drawing
Imports System.Drawing.Design
Imports System.Reflection
Imports System.Windows.Forms
Imports System.Windows.Forms.Design

'<Assembly: system.Security.Permissions.SecurityPermissionAttribute(system.Security.Permissions.SecurityAction.RequestMinimum)> 
'<Assembly: CLSCompliant(True)> 
Public Class SNPXComm

    Inherits System.ComponentModel.Component
    Implements AdvancedHMIDrivers.IComComponent

    '* Create a common instance to share so multiple SNPX can be used in a project
    Private Shared DLL As SNPX_DataLinkLayer
    'Private Shared DLL As Object

    Private DisableEvent As Boolean

    Public DataPackets() As System.Collections.ObjectModel.Collection(Of Byte)

    Public LastReadWasWord As Boolean

    'Need to double buffer these structures here so that they can be available to the form
    Public RequestMessage As XRequest
    Public ResponseMessage As XResponse
    Public BufferMessage As XBuffer

    Public Event DataReceived As SNPXCommEventHandler
    Public Event UnsolictedMessageRcvd As EventHandler
    'Public Event AutoDetectTry As EventHandler
    'Public Event DownloadProgress As EventHandler
    'Public Event UploadProgress As EventHandler

    'SNP-X doesn't have Transaction numbers nor a Target Node so index by SNP id 
    Private Shared PLCAddressByTNS() As ParsedDataAddress
    Public Responded() As Boolean

    '* keep the original address by ref of low TNS byte so it can be returned to a linked polling address
    Private CurrentAddressToRead As String
    Private PolledAddressList As New List(Of PolledAddressInfo)
    Private tmrPollList As New List(Of Windows.Forms.Timer)

    '*********************************************************************************
    '* This is used for linking a notification.
    '* An object can request a continuous poll and get a callback when value updated
    '*********************************************************************************
    Delegate Sub ReturnValues(ByVal Values As String)

    Private Class PolledAddressInfo
        Public PLCAddress As String
        Public Address As Integer
        Public Segment As String
        Public BitNumber As Integer
        Public ByteAddress As Integer
        Public DataIsWord As Boolean
        Public DataTypeModifier As Integer
        Public dlgCallBack As IComComponent.ReturnValues
        Public PollRate As Integer
        Public ID As Integer
        Public LastValue As String
        Public ElementsToRead As Integer
    End Class

    Private CurrentID As Integer = 1

    Private components As System.ComponentModel.IContainer

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New(ByVal container As System.ComponentModel.IContainer)
        MyClass.New()
        'Required for Windows.Forms Class Composition Designer support
        container.Add(Me)
    End Sub

    Public Sub New()
        MyBase.New()
        If DLL Is Nothing Then
            DLL = New SNPX_DataLinkLayer
        End If
        AddHandler DLL.DataReceived, Dr
        '****************************************************************
        '* Create WorkSpace at element (0) and default (1) as NULL SNPid
        '****************************************************************
        ReDim SNPidList(1)
        ReDim DataPackets(1)
        ReDim Responded(1)
        ReDim PLCAddressByTNS(1)
        SNPidList(0) = "       " 'Workspace
        SNPidList(1) = ""        'No SNPid defined
    End Sub

    'Component overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        '* The handle linked to the DataLink Layer has to be removed, otherwise it causes a problem when a form is closed
        RemoveHandler DLL.DataReceived, Dr

        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

#Region "Properties"
    Private m_BaudRate As String = "19200"
    <EditorAttribute(GetType(SNPXBaudRateEditor), GetType(System.Drawing.Design.UITypeEditor))> _
    Public Property BaudRate() As String
        Get
            Return m_BaudRate
            'Return DLL.BaudRate
        End Get
        Set(ByVal value As String)
            'If DLL IsNot Nothing AndAlso value <> m_BaudRate Then DLL.CloseComms()
            'm_BaudRate = value
            'If value <> DLL.BaudRate Then DLL.CloseComms()
            If value <> m_BaudRate Then
                If DLL IsNot Nothing Then
                    DLL.CloseComms()
                    Try
                        DLL.BaudRate = value
                    Catch ex As Exception
                        '* 0 means AUTO to the data link layer
                        DLL.BaudRate = 0
                    End Try
                End If
                m_BaudRate = value

            End If
            DLL.BaudRate = value
        End Set
    End Property

    'Private m_ComPort As String = "COM1"
    Public Property ComPort() As String
        Get
            Return DLL.ComPort
        End Get
        Set(ByVal value As String)
            If value <> DLL.ComPort Then DLL.CloseComms()
            DLL.ComPort = value
        End Set
    End Property

    Private m_Parity As System.IO.Ports.Parity = IO.Ports.Parity.Odd
    Public Property Parity() As System.IO.Ports.Parity
        Get
            Return DLL.Parity
        End Get
        Set(ByVal value As System.IO.Ports.Parity)
            If value <> DLL.Parity Then DLL.CloseComms()
            DLL.Parity = value
        End Set
    End Property

    Private m_StopBits As System.IO.Ports.StopBits = IO.Ports.StopBits.One
    Public Property StopBits() As System.IO.Ports.StopBits
        Get
            Return DLL.StopBits
        End Get
        Set(ByVal value As System.IO.Ports.StopBits)
            If value <> DLL.StopBits Then DLL.CloseComms()
            DLL.StopBits = value
        End Set
    End Property

    Private m_SNPid As String = ""
    <Description("Change ONLY if CPU has a SNP Ident")> _
    Public Property SNPid() As String
        Get
            Return DLL.SNPid
        End Get
        Set(ByVal value As String)
            DLL.SNPid = value
        End Set
    End Property

    '**************************************************************
    '* Determine whether to wait for a data read or raise an event
    '**************************************************************
    Private m_AsyncMode As Boolean
    <Description("Asynchronous data reads")> _
    Public Property AsyncMode() As Boolean
        Get
            Return m_AsyncMode
        End Get
        Set(ByVal value As Boolean)
            m_AsyncMode = value
        End Set
    End Property

    '**************************************************************
    '* Stop the polling of subscribed data
    '**************************************************************
    Private m_DisableSubscriptions
    <Description("Disable polling updates")> _
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

        Set(ByVal Value As Windows.Forms.Form)
            If Not Value Is Nothing Then
                m_SynchronizingObject = Value
            End If
        End Set
    End Property

    Private m_Attached As Boolean = False
    <Description("Driver Attached to PLC with SNPid")> _
    <System.ComponentModel.ReadOnly(True)> _
    Public Property Attached() As Boolean
        Get
            Return m_Attached
        End Get
        Set(ByVal value As Boolean)
            If (value) Then
                If (Not m_Attached) Then
                    If (0 = SNPX_Attach()) Then
                        m_Attached = value
                    Else
                        m_Attached = False
                    End If
                End If
            Else
                m_Attached = False
            End If
        End Set
    End Property
#End Region

#Region "Public Methods"
    'Private InternalRequest As Boolean '* This is used to dinstinquish when to send data back to notification request

Public Function Subscribe(ByVal PLCAddress As String, ByVal numberOfElements As Short, ByVal updateRate As Integer, ByVal CallBack As IComComponent.ReturnValues) As Integer Implements IComComponent.Subscribe
        '*******************************************************************
        '*******************************************************************
        Dim LocalPollRate As Integer = 75
        If (LocalPollRate < updateRate) Then updateRate = LocalPollRate

        If (Not Attached) Then
            Attached = True
        End If

        Dim ParsedResult As ParsedDataAddress = ParseAddress(PLCAddress)
        Dim StartAddressInt As Integer
        Try
            StartAddressInt = ParsedResult.Address
        Catch ex As Exception
            Return -1
        End Try

        '* Valid address?
        Dim tmpPA As New PolledAddressInfo
        tmpPA.PLCAddress = PLCAddress
        tmpPA.Address = ParsedResult.Address
        tmpPA.Segment = ParsedResult.Segment
        tmpPA.BitNumber = ParsedResult.BitNumber
        tmpPA.ByteAddress = ParsedResult.ByteAddress
        tmpPA.DataIsWord = ParsedResult.DataIsWord
        tmpPA.DataTypeModifier = ParsedResult.DataTypeModifier
        If (tmpPA.DataTypeModifier = DataTypeModifier.TypeByte) Then
            'No Byte acess for subscribed addresses
            tmpPA.DataTypeModifier = DataTypeModifier.TypeDefault
        End If
        If (tmpPA.DataTypeModifier > DataTypeModifier.TypeDefault) Then
            'Two 16bit words
            tmpPA.ElementsToRead = 2
        Else
            tmpPA.ElementsToRead = 1
        End If
        'If (tmpPA.Segment = "R") Then 'Put Registers on slower update rate
        '    updateRate = 500
        'End If
        tmpPA.PollRate = updateRate
        tmpPA.dlgCallBack = CallBack
        tmpPA.ID = CurrentID

        PolledAddressList.Add(tmpPA)
        PolledAddressList.Sort(AddressOf SortPolledAddresses)

        '* The ID is used as a reference for removing polled addresses
        CurrentID += 1

        '********************************************************************
        '* Check to see if there already exists a timer for this poll rate
        '********************************************************************
        Dim j As Integer = 0
        While j < tmrPollList.Count AndAlso tmrPollList(j) IsNot Nothing AndAlso tmrPollList(j).Interval <> updateRate
            j += 1
        End While

        If j >= tmrPollList.Count Then
            '* Add new timer
            Dim tmrTemp As New Windows.Forms.Timer
            If updateRate > 0 Then
                tmrTemp.Interval = updateRate
            Else
                tmrTemp.Interval = 500 'default
            End If
            tmrTemp.Interval = updateRate
            tmrPollList.Add(tmrTemp)
            AddHandler tmrPollList(j).Tick, AddressOf PollUpdate
            tmrTemp.Enabled = True
        End If

        Return tmpPA.ID
    End Function

    '***************************************************************
    '* Used to sort polled addresses by Segment and Address
    '* This helps in optimizing reading
    '**************************************************************
    Private Function SortPolledAddresses(ByVal A1 As PolledAddressInfo, ByVal A2 As PolledAddressInfo) As Integer
        If A1.Segment = A2.Segment Then
            If A1.Address > A2.Address Then
                Return 1
            ElseIf A1.Address = A2.Address Then
                Return 0
            Else
                Return -1
            End If
        ElseIf (A1.Segment > A2.Segment) Then
            Return 1
        Else
            Return -1
        End If
    End Function

    Public Function UnSubscribe(ByVal ID As Integer) As Integer Implements IComComponent.Unsubscribe
        Dim i As Integer = 0
        While i < PolledAddressList.Count AndAlso PolledAddressList(i).ID <> ID
            i += 1
        End While

        If i < PolledAddressList.Count Then
            PolledAddressList.RemoveAt(i)
            If PolledAddressList.Count = 0 Then
                For j As Integer = 0 To tmrPollList.Count - 1
                    tmrPollList(j).Enabled = False
                    tmrPollList.Remove(tmrPollList(0))
                Next
            End If
        End If
    End Function

    '**************************************************************
    '* Perform the reads for the variables added for notification
    '* Attempt to optimize by grouping reads
    '**************************************************************
    Private InternalRequest As Boolean '* This is used to dinstinquish when to send data back to notification request
    Private SavedPollRate As Integer
    Private Sub PollUpdate(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If m_DisableSubscriptions Then Exit Sub

        Dim intTimerIndex As Integer = tmrPollList.IndexOf(sender)

        '* Stop the poll timer
        tmrPollList(intTimerIndex).Enabled = False

        Dim i, NumberToRead, FirstElement As Integer

        Dim Segment As String = PolledAddressList(0).Segment
        While i < PolledAddressList.Count
            Dim NumberToReadCalc As Integer
            Dim Difference As Integer
            NumberToRead = PolledAddressList(i).ElementsToRead
            FirstElement = i
            Dim MaxElementsPerRead As Integer
            Dim SearchComplete As Boolean
            NumberToReadCalc = NumberToRead

            If (PolledAddressList(i).DataIsWord) Then
                MaxElementsPerRead = 64
            Else
                MaxElementsPerRead = 128
            End If
            SearchComplete = False
            While Not SearchComplete
                'Setup group read by segment and MaxElementsPerRead. Also seperate if on different pollrates
                Difference = PolledAddressList(i).Address - PolledAddressList(FirstElement).Address
                If (PolledAddressList(i).Segment <> Segment) Then
                    SearchComplete = True
                    Segment = PolledAddressList(i).Segment
                    i -= 1 ' Backup
                ElseIf (PolledAddressList(i).PollRate <> PolledAddressList(FirstElement).PollRate) Then
                    SearchComplete = True
                    i -= 1 ' Backup
                ElseIf (Difference < MaxElementsPerRead) Then
                    i += 1
                    NumberToReadCalc = Difference + 1
                Else
                    SearchComplete = True
                    i -= 1 ' Backup
                End If
                If (i > (PolledAddressList.Count - 1)) Then
                    SearchComplete = True
                End If
            End While

            Dim PLCAddress As String = PolledAddressList(FirstElement).PLCAddress
            PolledAddressList(FirstElement).ElementsToRead = NumberToReadCalc

            'Check to see if last address in read sequence has modifier and if so add 1 to read count
            Dim j As Integer = i
            If (j > (PolledAddressList.Count - 1)) Then
                j = PolledAddressList.Count - 1
            End If
            If (PolledAddressList(j).DataTypeModifier > DataTypeModifier.TypeDefault) Then
                NumberToReadCalc += 1
            End If

            If PolledAddressList(FirstElement).PollRate = sender.Interval Or SavedPollRate > 0 Then
                '* Make sure it does not wait for return value before coming back
                Dim tmp As Boolean = Me.AsyncMode
                Me.AsyncMode = True
                Try
                    InternalRequest = True
                    Me.ReadAny(PLCAddress, NumberToReadCalc)
                    InternalRequest = False
                    If SavedPollRate <> 0 Then
                        tmrPollList(0).Interval = SavedPollRate
                        SavedPollRate = 0
                    End If
                Catch ex As Exception
                    '* Send this message back to the requesting control
                    Dim TempArray() As String = {ex.Message}

                    m_SynchronizingObject.BeginInvoke(PolledAddressList(i).dlgCallBack, CObj(TempArray))
                    '* Slow down the poll rate to avoid app freezing
                    If SavedPollRate = 0 Then SavedPollRate = tmrPollList(intTimerIndex).Interval
                    tmrPollList(intTimerIndex).Interval = 5000
                End Try
                Me.AsyncMode = tmp
            End If
            i += 1
        End While
        '* Start the poll timer
        tmrPollList(intTimerIndex).Enabled = True
    End Sub


    Public Function OpenComm() As Integer
        Return DLL.OpenComms()
    End Function

    'Closes the comm port
    Public Sub CloseComms()
        DLL.CloseComms()
    End Sub

    Public Function ReadSynchronous(ByVal startAddress As String, ByVal numberOfElements As Integer) As String() Implements IComComponent.ReadSynchronous
        Return ReadAny(startAddress, numberOfElements, False)
    End Function

    Public Function ReadAny(ByVal startAddress As String, ByVal numberOfElements As Integer) As String() Implements IComComponent.ReadAny
        Return ReadAny(startAddress, numberOfElements, m_AsyncMode)
    End Function



    '******************************************
    '* Synchronous read of any data type
    '*  this function does not declare its return type because it dependent on the data type read
    '******************************************
    ''' <summary>
    ''' Synchronous read of any data type
    ''' this function returns results as an array of strings
    ''' </summary>
    ''' <param name="startAddress"></param>
    ''' <param name="numberOfElements"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ReadAny(ByVal startAddress As String, ByVal numberOfElements As Integer, ByVal asyncModeIn As Boolean) As String()

        If (Not Attached) Then
            Attached = True
        End If
        If (Not Attached) Then
            'Still not Attached get out
            Dim x() As String = {"0"}
            Return x
        End If

        Dim ParsedResult As ParsedDataAddress = ParseAddress(startAddress)

        If ((ParsedResult.DataTypeModifier > DataTypeModifier.TypeDefault) And (ParsedResult.DataTypeModifier < DataTypeModifier.TypeByte)) And (Not InternalRequest) Then
            'Directly called read not a subscription. If address is modified and not a discrete byte read then data is 4 bytes instead of 2.
            ParsedResult.NumberOfElements = numberOfElements * 2
        Else
            ParsedResult.NumberOfElements = numberOfElements
        End If

        '* Invalid address?
        Dim StartAddressInt As Integer
        Try
            StartAddressInt = ParsedResult.Address
        Catch ex As Exception
            Throw New SNPXException("Invalid Address")
        End Try

        'Check if SNPid Property has changed and if so update current to reflect
        Dim i As Integer = SNPidIndex(Me.SNPid)
        If (i = -1) Then 'Not in list Throw Exception
            Throw New SNPXException("SNPid " & Me.SNPid & " not Attached.")
        End If
        ParsedResult.SNPid = Me.SNPid

        '**********************************************************************
        '* Link the TNS to the original address for use by the linked polling
        '**********************************************************************
        'PLCAddressByTNS(SNPidCurrentIndex).InternallyRequested = ParsedResult.IntInternalRequest
        PLCAddressByTNS(SNPidCurrentIndex).Address = ParsedResult.Address
        PLCAddressByTNS(SNPidCurrentIndex).Segment = ParsedResult.Segment
        PLCAddressByTNS(SNPidCurrentIndex).DataIsWord = ParsedResult.DataIsWord
        PLCAddressByTNS(SNPidCurrentIndex).DataTypeModifier = ParsedResult.DataTypeModifier
        PLCAddressByTNS(SNPidCurrentIndex).BitNumber = ParsedResult.BitNumber
        PLCAddressByTNS(SNPidCurrentIndex).ByteAddress = ParsedResult.ByteAddress
        PLCAddressByTNS(SNPidCurrentIndex).NumberOfElements = ParsedResult.NumberOfElements
        PLCAddressByTNS(SNPidCurrentIndex).SNPid = ParsedResult.SNPid
        PLCAddressByTNS(SNPidCurrentIndex).PLCAddress = ParsedResult.PLCAddress

        Dim rTNS As Integer = SNPidCurrentIndex
        Dim result As Integer
        '    InternalRequest = InternalRequest

        Dim Data(1) As Byte 'Send zero'ed data

        Dim Wait As Boolean
        If Not asyncModeIn Then
            Wait = True
        Else
            Wait = False
        End If
        Dim BitLevel As Boolean = Not (ParsedResult.DataTypeModifier = DataTypeModifier.TypeByte)
        If (ParsedResult.InternallyRequested) Then
            'Subscribed reads only allowed bitlevel access
            BitLevel = True
        End If
        If (Not BitLevel) Then
            'Read the byte not bit
            ParsedResult.Address = ParsedResult.ByteAddress
        End If

        result = PrefixAndSend(MessageType.SNP_X, RequestCode.XRead, Data, Wait, rTNS, ParsedResult, BitLevel)

        If m_AsyncMode Then
            Dim x() As String = {result}
            Return x
        Else
            If (PLCAddressByTNS(rTNS).DataIsWord) Then
                LastReadWasWord = True
            Else
                LastReadWasWord = False
            End If
            Return ExtractData(PLCAddressByTNS(rTNS), ResponseMessage)
        End If

    End Function

    '*************************************************************
    '* Overloaded method of ReadAny - that reads only one element
    '*************************************************************
    ''' Synchronous read of any data type
    ''' this function returns results as a string

    Public Function ReadAny(ByVal startAddress As String) As String Implements IComComponent.ReadAny
        Try
            Return ReadAny(startAddress, 1)(0)
        Catch ex As Exception
            'Throw New SNPXException("No Read Data to return - Possible Invalid Address")
            If (Not CurrentSNPXRequestInError) Then
                Throw New SNPXException("No Read Data to return - Possible Invalid Address")
            Else
                Return "0"
            End If
        End Try
    End Function


    '*****************************************************************
    '* Write Section
    '*
    '* Address is in the form of <segment>:<offset>
    '* examples  Q0001, R0010, S32, AI1
    '******************************************************************
    '* Handle one value of Integer type
    ''' <summary>
    ''' Write a single integer value to a PLC data table
    ''' </summary>
    ''' <param name="startAddress"></param>
    ''' <param name="dataToWrite"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function WriteData(ByVal startAddress As String, ByVal dataToWrite As Integer) As Integer
        Dim temp(1) As Integer
        temp(0) = dataToWrite
        Return WriteData(startAddress, 1, temp)
    End Function


    '* Write an array of integers
    ''' <summary>
    ''' Write multiple consectutive integer values to a PLC data table
    ''' The startAddress is in the common form of GE FANUC addressing (e.g. R100)
    ''' </summary>
    ''' <param name="startAddress"></param>
    ''' <param name="numberOfElements"></param>
    ''' <param name="dataToWrite"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function WriteData(ByVal startAddress As String, ByVal numberOfElements As Integer, ByVal dataToWrite() As Integer) As Integer
        Dim ParsedResult As ParsedDataAddress = ParseAddress(startAddress)
        Dim Data(1) As Byte 'Default 2 byte write
        Dim DoXBuffer As Boolean
        Dim BitLevel As Boolean

        If (Not Attached) Then
            Attached = True
        End If
        If (Not Attached) Then
            'Still not Attached get out
            Return -3
        End If

        ParsedResult.NumberOfElements = numberOfElements

        '* Invalid address?
        Dim StartAddressInt As Integer
        Try
            StartAddressInt = ParsedResult.Address
        Catch ex As Exception
            Throw New SNPXException("Invalid Address")
        End Try

        'Check if SNPid Property has changed and if so update current to reflect
        Dim i As Integer = SNPidIndex(Me.SNPid)
        If (i = -1) Then 'Not in list Throw Exception
            Throw New SNPXException("SNPid " & Me.SNPid & " not Attached.")
        End If
        ParsedResult.SNPid = Me.SNPid

        '**********************************************************************
        '* Link the TNS to the original address for use by the linked polling
        '**********************************************************************
        'PLCAddressByTNS(SNPidCurrentIndex).InternallyRequested = ParsedResult.IntInternalRequest
        PLCAddressByTNS(SNPidCurrentIndex).Address = ParsedResult.Address
        PLCAddressByTNS(SNPidCurrentIndex).Segment = ParsedResult.Segment
        PLCAddressByTNS(SNPidCurrentIndex).DataIsWord = ParsedResult.DataIsWord
        PLCAddressByTNS(SNPidCurrentIndex).DataTypeModifier = ParsedResult.DataTypeModifier
        PLCAddressByTNS(SNPidCurrentIndex).BitNumber = ParsedResult.BitNumber
        PLCAddressByTNS(SNPidCurrentIndex).ByteAddress = ParsedResult.ByteAddress
        PLCAddressByTNS(SNPidCurrentIndex).NumberOfElements = ParsedResult.NumberOfElements
        PLCAddressByTNS(SNPidCurrentIndex).SNPid = ParsedResult.SNPid
        PLCAddressByTNS(SNPidCurrentIndex).PLCAddress = ParsedResult.PLCAddress
        PLCAddressByTNS(SNPidCurrentIndex).NumberOfElements = numberOfElements
        ParsedResult.NumberOfElements = numberOfElements

        Dim rTNS As Integer = SNPidCurrentIndex
        Dim result As Integer
        '    InternalRequest = InternalRequest

        If (ParsedResult.DataIsWord) Then
            BitLevel = False
            If (numberOfElements <= 1) Then
                Dim DataWord As Int16
                '2 Byte Write
                Try
                    DataWord = CShort(dataToWrite(0))
                Catch ex As Exception
                    'Overflow return no response
                    Return -8
                End Try
                Data(0) = (DataWord And &HFF)
                Data(1) = ((DataWord And &HFF00) >> 8)
                DoXBuffer = False
            Else
                Dim DataWord As UShort
                'More than 2 bytes so increase data size and set XBuffer flag 
                ReDim Data((numberOfElements * 2) - 1)
                For j As Integer = 0 To numberOfElements - 1
                    Try
                        DataWord = (dataToWrite(j) And &HFFFF) 'CShort(dataToWrite(j))
                    Catch ex As Exception
                        'Overflow return no response
                        Return -8
                    End Try
                    Data(j * 2) = (DataWord And &HFF)
                    Data(j * 2 + 1) = ((DataWord And &HFF00) >> 8)
                Next j
                DoXBuffer = True
            End If
        Else
            BitLevel = Not (ParsedResult.DataTypeModifier = DataTypeModifier.TypeByte) 'True 'For now do all discretes at bit level
            If (numberOfElements <= 1) Then
                Data(1) = 0
                If (BitLevel) Then
                    If (dataToWrite(0) < 0) Or (dataToWrite(0) > 1) Then
                        Return (-30) ' Data out of range
                    End If
                    Data(0) = ((dataToWrite(0) * 255) And (2 ^ ParsedResult.BitNumber))
                Else
                    If (dataToWrite(0) < 0) Or (dataToWrite(0) > 255) Then
                        Return (-30) ' Data out of range
                    End If
                    Data(0) = dataToWrite(0)
                End If
            Else
                Dim Address As Integer = ParsedResult.Address
                Dim ByteAddress As Integer = ParsedResult.ByteAddress
                Dim BitNumber As Integer = ParsedResult.BitNumber
                Dim LastByteAddress As Integer = ByteAddress
                Dim DataIndex As Integer = 0
                'More than one discrete so pack the bytes
                For j As Integer = 0 To numberOfElements - 1
                    If (ByteAddress > LastByteAddress) Then
                        DataIndex += 1
                        ReDim Preserve Data(DataIndex)
                    End If
                    If (BitLevel) Then
                        If (dataToWrite(j) < 0) Or (dataToWrite(j) > 1) Then
                            Return (-30) ' Data out of range
                        End If
                        Data(DataIndex) = Data(DataIndex) Or ((dataToWrite(j) * 255) And (2 ^ BitNumber))
                    Else
                        If (dataToWrite(j) < 0) Or (dataToWrite(j) > 255) Then
                            Return (-30) ' Data out of range
                        End If
                        ReDim Preserve Data(j)
                        Data(j) = dataToWrite(j)
                    End If
                    LastByteAddress = ByteAddress
                    Address += 1
                    ByteAddress = Math.Floor(Address / 8)
                    BitNumber = ((Address / 8) - ByteAddress) * 8
                Next j
                If (Data.Length > 2) Then
                    'More than 2 bytes so set XBuffer flag
                    DoXBuffer = True
                End If
            End If
            If Not (BitLevel) Then
                ParsedResult.Address = ParsedResult.ByteAddress
            End If
        End If

        Dim Wait As Boolean
        If Not m_AsyncMode Then
            Wait = True
        Else
            Wait = False
        End If

        If (DoXBuffer) Then
            result = PrefixAndSend(MessageType.SNP_X, RequestCode.XWrite, Data, True, rTNS, ParsedResult, BitLevel)
            result = PrefixAndSend(MessageType.DataBuffer, RequestCode.XWrite, Data, Wait, rTNS, ParsedResult, BitLevel)
        Else
            result = PrefixAndSend(MessageType.SNP_X, RequestCode.XWrite, Data, Wait, rTNS, ParsedResult, BitLevel)
        End If
        Return result
    End Function

    ''* Handle one value of Single type
    '''' <summary>
    '''' Write a single floating point value to a data table
    '''' The startAddress is in the common form of AB addressing (e.g. F8:0)
    '''' </summary>
    '''' <param name="startAddress"></param>
    '''' <param name="dataToWrite"></param>
    '''' <returns></returns>
    '''' <remarks></remarks>
    'Public Function WriteData(ByVal startAddress As String, ByVal dataToWrite As Single) As Integer
    '    Dim temp(1) As Single
    '    temp(0) = dataToWrite
    '    Return WriteData(startAddress, 1, temp)
    'End Function

    ''* Write an array of Singles
    '''' <summary>
    '''' Write multiple consectutive floating point values to a PLC data table
    '''' The startAddress is in the common form of AB addressing (e.g. F8:0)
    '''' </summary>
    '''' <param name="startAddress"></param>
    '''' <param name="numberOfElements"></param>
    '''' <param name="dataToWrite"></param>
    '''' <returns></returns>
    '''' <remarks></remarks>
    'Public Function WriteData(ByVal startAddress As String, ByVal numberOfElements As Integer, ByVal dataToWrite() As Single) As Integer


    'End Function


    '* Write a String
    ''' <summary>
    ''' Write a string value to a string data table
    ''' The startAddress is in the common form of GE Fanuc addressing (e.g. M100)
    ''' </summary>
    ''' <param name="startAddress"></param>
    ''' <param name="dataToWrite"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function WriteData(ByVal startAddress As String, ByVal dataToWrite As String) As String Implements IComComponent.WriteData

        Dim ParsedResult As ParsedDataAddress = ParseAddress(startAddress)

        Select Case ParsedResult.DataTypeModifier
            Case DataTypeModifier.TypeInt32
                ' Convert string to 32 bit signed integer then send to WriteData as two signed integers where they will be
                ' convert to 16 bit signed integers then sent to the DLL driver that will futher break it down sending as bytes.
                ' At the PLC the 16 bit word address following the address passed to this sub will also be written to.
                ' No boundary check takes place.       
                Dim result As Integer
                If (Integer.TryParse(dataToWrite, result)) Then
                    Dim results(1) As Integer
                    Dim IntegerBytes() As Byte = BitConverter.GetBytes(result)
                    results(0) = IntegerBytes(1)
                    results(0) = (results(0) << 8) Or IntegerBytes(0)
                    results(1) = IntegerBytes(3)
                    results(1) = (results(1) << 8) Or IntegerBytes(2)
                    Return (WriteData(startAddress, 2, results))
                End If
                Return (-30) ' Data out of range
            Case DataTypeModifier.TypeReal
                ' Convert string to 32 bit floating point then send to WriteData as two signed integers where they will be
                ' convert to 16 bit signed integers then sent to the DLL driver that will futher break it down sending as bytes.
                ' At the PLC the 16 bit word address following the address passed to this sub will also be written to.
                ' No boundary check takes place.
                Dim result As Single
                If (Single.TryParse(dataToWrite, result)) Then
                    Dim results(1) As Integer
                    Dim IntegerBytes() As Byte = BitConverter.GetBytes(result)
                    results(0) = IntegerBytes(1)
                    results(0) = (results(0) << 8) Or IntegerBytes(0)
                    results(1) = IntegerBytes(3)
                    results(1) = (results(1) << 8) Or IntegerBytes(2)
                    Return (WriteData(startAddress, 2, results))
                End If
                Return (-30) ' Data out of range
            Case DataTypeModifier.TypeByte
                ' Convert string to a byte then send to WriteData as a signed integer.
                Dim result As Byte
                If (Byte.TryParse(dataToWrite, System.Globalization.NumberStyles.AllowHexSpecifier, Nothing, result)) Then
                    Dim results As Integer = result
                    Return (WriteData(startAddress, results))
                End If
                Return (-30) ' Data out of range
            Case Else
                'No modifier so send treating as an integer
                Return WriteData(startAddress, CInt(dataToWrite))
        End Select

    End Function

    ''**************************************************************
    ''* Write to a PLC data file
    ''*
    ''**************************************************************
    'Private Function WriteRawData(ByVal ParsedResult As ParsedDataAddress, ByVal numberOfBytes As Integer, ByVal dataToWrite() As Byte) As Integer
    '    'NA GE Fanuc
    'End Function

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
    Private Structure ParsedDataAddress
        Dim PLCAddress As String
        Dim SNPid As String
        Dim ResponseCode As Byte
        Dim Segment As String
        Dim Address As String
        Dim ByteAddress As Integer
        Dim BitNumber As Integer
        Dim InternallyRequested As Boolean
        Dim NumberOfElements As Integer
        Dim RequestActive As Boolean
        Dim DataIsWord As Boolean
        Dim DataTypeModifier As Integer
    End Structure

    Private Enum DataTypeModifier As Integer
        TypeDefault = 0
        TypeInt32 = 1
        TypeReal = 2
        TypeByte = 3
    End Enum


    Private Dr As EventHandler = AddressOf SNPX_DataLinkLayer_DataReceived
    Private Const LongBreakDuration As Integer = 100
    Private Const T4Time As Integer = 50

    Public SNPidList() As String
    Public SNPidCount As Integer
    Public SNPidCurrentIndex As Integer


    Private InputPatterns As Regex() = { _
 New Regex("(?'segment'[IQGMTSRA])"), _
 New Regex("(?'segment'[S][ABC])"), _
 New Regex("(?'segment'[A][IQ])"), _
 New Regex("(?'segment'[IQGMTSR])(?'address'\d{1,5})"), _
 New Regex("(?'segment'[S][ABC])(?'address'\d{1,5})"), _
 New Regex("(?'segment'[A][IQ])(?'address'\d{1,5})"), _
 New Regex("(?'segment'[IQGMTSR])(?'address'\d{1,5})(?'flag'[%])(?'type'[RLB])"), _
 New Regex("(?'segment'[S][ABC])(?'address'\d{1,5})(?'flag'[%])(?'type'[RLB])"), _
 New Regex("(?'segment'[A][IQ])(?'address'\d{1,5})(?'flag'[%])(?'type'[RLB])") _
 }
    Private BitTypes As New Regex("[IQGMTS]")
    Private WordTypes As New Regex("[RA]")

    Private Function ParseAddress(ByVal DataAddress As String) As ParsedDataAddress

        Dim NumOfPatterns As Integer
        Dim AddressToParse As String = DataAddress.ToUpper()
        Dim LastGood_mc As MatchCollection
        Dim mc As MatchCollection
        Dim segment As String = ""
        Dim address As String = ""
        Dim typeModifier As String = ""
        Dim gotmatch As Boolean
        Dim result As New ParsedDataAddress


        AddressToParse = AddressToParse.Replace(" ", "")
        NumOfPatterns = InputPatterns.Length - 1
        LastGood_mc = InputPatterns(0).Matches("")
        For i As Integer = 0 To NumOfPatterns
            mc = InputPatterns(i).Matches(AddressToParse)
            If (mc.Count > 0) Then
                gotmatch = True
            Else
                gotmatch = False
            End If
            If (gotmatch) Then
                If mc.Item(0).Groups("segment").Length > 0 Then
                    segment = mc.Item(0).Groups("segment").Value
                End If
                If mc(0).Groups("address").Length > 0 Then
                    address = mc.Item(0).Groups("address").Value
                Else
                    LastGood_mc = mc
                End If
                If (mc.Item(0).Groups("flag").Length > 0) And (mc.Item(0).Groups("type").Length > 0) Then
                    typeModifier = mc.Item(0).Groups("type").Value
                End If
            End If
        Next i
        If (segment <> "") And (address <> "") Then
            result.Address = address
            result.Segment = segment
            'Address looks good so futher process it
            result.Address -= 1 'Zero Based - But no negative numbers
            If (result.Address < 0) Then
                result.Address = 0
            End If
            'Dim WordNumber As Integer = Math.Floor(result.Address / 16)
            Dim ByteAddress As Integer = Math.Floor(result.Address / 8)
            Dim BitNumber As Integer = ((result.Address / 8) - ByteAddress) * 8
            mc = WordTypes.Matches(segment)
            If ((mc.Count > 0) And (result.Segment <> "SA")) Then
                Select Case typeModifier
                    Case "R"
                        result.DataTypeModifier = DataTypeModifier.TypeReal
                    Case "L"
                        result.DataTypeModifier = DataTypeModifier.TypeInt32
                    Case Else
                        result.DataTypeModifier = DataTypeModifier.TypeDefault
                End Select
                result.DataIsWord = True
            Else
                'Otherwise a discrete type
                If (typeModifier = "B") Then
                    result.DataTypeModifier = DataTypeModifier.TypeByte
                Else
                    result.DataTypeModifier = DataTypeModifier.TypeDefault
                End If
                result.DataIsWord = False
                result.ByteAddress = ByteAddress
                result.BitNumber = BitNumber
            End If
            result.PLCAddress = DataAddress
        End If
        Return result
    End Function


    '****************************************************
    '* Wait for a response from PLC before returning
    '****************************************************
    Dim MaxTicks As Integer = 50  'ticks per second
    Private Function WaitForResponse(ByVal rTNS As Integer) As Integer
        'Responded = False

        Dim Loops As Integer = 0
        While Not Responded(rTNS) And Loops < MaxTicks
            'Application.DoEvents()
            System.Threading.Thread.Sleep(20)
            Loops += 1
        End While

        If Loops >= MaxTicks Then
            Return -20
        ElseIf DLL.LastResponseWasNAK Then
            Return -21
        Else
            Return 0
        End If
    End Function
    '**************************************************************
    ' Build the X-Message packet
    '**************************************************************
    Private Function PrefixAndSend(ByVal Type As Byte, ByVal Request As Byte, ByVal Data() As Byte, ByVal Wait As Boolean, ByRef rTNS As Integer, ByVal ParsedData As ParsedDataAddress, ByVal BitLevel As Boolean) As Integer

        Dim xm As XRequest
        'Dim the structure arrary to proper sizes
        ReDim xm.SNPid(0 To 7)
        ReDim xm.Data(0 To 6)
        Dim xb As XBuffer
        Dim PacketSize As Integer = 22 'default Read/Write (X-Message - BCC byte that DLL will add before sending)
        If (Type = MessageType.DataBuffer) Then
            PacketSize = (Data.Length + 6) 'For X-Buffer # data Bytes + 7 (zero based array so -1)
        End If
        Dim CommandPacket(0 To PacketSize) As Byte

        If (Type = MessageType.DataBuffer) Then
            'X-Buffer is the buffer message for a write of more than 2 bytes
            Dim dl As Integer = Data.Length - 1
            ReDim xb.Data(dl)
            xb.Start = MessageType.Start
            xb.MessageType = MessageType.DataBuffer
            Data.CopyTo(xb.Data, 0)
            xb.EndOfBlock = MessageType.EndOfBlock
            xb.NextMessageType = 0
            xb.NextMessageLength = 0
            xb.NotUsed = 0

            DLL.BufferMessage = xb 'copy structure to Dll

            'Build the CommandPacket Buffer
            CommandPacket(0) = xb.Start
            CommandPacket(1) = xb.MessageType
            xb.Data.CopyTo(CommandPacket, 2)
            CommandPacket(PacketSize - 4) = xm.EndOfBlock
            CommandPacket(PacketSize - 3) = xm.NextMessageType
            CommandPacket(PacketSize - 2) = (xm.NextMessageLength And &HFF)
            CommandPacket(PacketSize - 1) = (xm.NextMessageLength >> 8)
            CommandPacket(PacketSize) = xm.NotUsed

        Else 'Normal X-SNP  X-Read or X-Write or X-Intermediate
            'Fill in structure although it will go mostly unused.
            xm.Start = MessageType.Start
            xm.MessageType = MessageType.SNP_X
            Dim Max As Integer = Len(ParsedData.SNPid)
            If Max > 7 Then
                Max = 7
            End If
            If (Max > 0) Then
                For i As Integer = 0 To (Max - 1)
                    xm.SNPid(i) = CByte(Asc(ParsedData.SNPid(i)))
                Next
            End If
            xm.RequestCode = Request
            xm.Data(0) = DecodeSegment(ParsedData.Segment, BitLevel)
            xm.Data(1) = (ParsedData.Address And &HFF)
            xm.Data(2) = (ParsedData.Address >> 8)
            xm.Data(3) = (ParsedData.NumberOfElements And &HFF)
            xm.Data(4) = (ParsedData.NumberOfElements >> 8)
            xm.Data(5) = Data(0)
            xm.Data(6) = Data(1)
            xm.EndOfBlock = MessageType.EndOfBlock
            If (Data.Length > 2) Then
                'if the data buffer array is more than 2 elements this is a X-Write request that will be followed by a X-Buffer 
                xm.NextMessageType = MessageType.DataBuffer
                xm.NextMessageLength = Data.Length + 8
                xm.Data(5) = 0
                xm.Data(6) = 0
            Else
                xm.NextMessageType = 0
                xm.NextMessageLength = 0
            End If
            xm.NotUsed = 0

            DLL.RequestMessage = xm 'copy structure to Dll

            'Build the CommandPacket Buffer
            CommandPacket(0) = xm.Start
            CommandPacket(1) = xm.MessageType
            Buffer.BlockCopy(xm.SNPid, 0, CommandPacket, 2, 8)
            CommandPacket(10) = xm.RequestCode
            Buffer.BlockCopy(xm.Data, 0, CommandPacket, 11, 7)
            CommandPacket(18) = xm.EndOfBlock
            CommandPacket(19) = xm.NextMessageType
            CommandPacket(20) = (xm.NextMessageLength And &HFF)
            CommandPacket(21) = (xm.NextMessageLength >> 8)
            CommandPacket(22) = xm.NotUsed
        End If

        '*Mark whether this was requested by a subscription or not
        '* FIX
        PLCAddressByTNS(rTNS).InternallyRequested = InternalRequest
        PLCAddressByTNS(rTNS).RequestActive = True

        Responded(rTNS) = False
        Dim result As Integer

        result = SendData(CommandPacket)
        If result = 0 And Wait Then
            result = WaitForResponse(rTNS)
            '* Return status byte that came from controller
            If result = 0 Then
                If DataPackets(rTNS) IsNot Nothing Then
                    result = ResponseMessage.StatusWord
                Else
                    result = -8 '* no response came back from PLC
                End If
            Else
                Dim DebugCheck As Integer = 0
            End If
        ElseIf ((result = -3) And Wait) Then
            'No longer attached
            Attached = False
        Else
        End If
        'IncrementTNS()
        Return result
    End Function


    Private Function ExtractData(ByVal ParsedResult As ParsedDataAddress, ByVal xm As XResponse) As String()

        Dim DataIsWord As Boolean = ParsedResult.DataIsWord
        Dim DataAsString(0) As String
        Dim DataStringList As New List(Of String)
        Dim DataByteLength As Integer = xm.OptionalData.Length
        Dim CurrentIndex As Integer = -1
        Dim NumOfMatches As Integer = 0

        If (DataIsWord) Then
            Dim DataWordLength As Integer = DataByteLength / 2
            If DataWordLength <> ParsedResult.NumberOfElements Then
                DataWordLength = 0
            End If
            If (DataWordLength > 0) Then
                Dim AddressInRange As Boolean = True
                Dim FirstAddress As Integer = ParsedResult.Address
                Dim LastAddress As Integer = FirstAddress + DataWordLength
                Dim Modifier As Integer = 0
                Dim ByteIndex As Integer
                Dim ConvertBuffer(3) As Byte
                Dim ResultAsReal As Single
                Dim ResultAsInt32 As Integer
                Dim ResultAsShort As Short

                If (ParsedResult.InternallyRequested) Then
                    'This is an internal subscribe so get the index of the Polled List for this startng address
                    For CurrentIndex = 0 To PolledAddressList.Count - 1
                        If (ParsedResult.PLCAddress = PolledAddressList(CurrentIndex).PLCAddress) Then
                            Exit For
                        End If
                    Next
                Else
                    'Not an internal request so apply any modifier to every element
                    Modifier = ParsedResult.DataTypeModifier
                End If
                While (AddressInRange)
                    If (CurrentIndex > -1) Then
                        'if working off an internal request then get modifier from Poll List
                        Modifier = PolledAddressList(CurrentIndex).DataTypeModifier
                    End If
                    Select Case Modifier
                        Case 1
                            'Int32
                            Buffer.BlockCopy(xm.OptionalData, ByteIndex, ConvertBuffer, 0, 4)
                            ByteIndex += 4
                            ResultAsInt32 = BitConverter.ToInt32(ConvertBuffer, 0)
                            DataStringList.Add(ResultAsInt32.ToString)
                        Case 2
                            'Real
                            Buffer.BlockCopy(xm.OptionalData, ByteIndex, ConvertBuffer, 0, 4)
                            ByteIndex += 4
                            ResultAsReal = BitConverter.ToSingle(ConvertBuffer, 0)
                            DataStringList.Add(ResultAsReal.ToString)
                        Case Else
                            'default signed 16bit integer
                            Buffer.BlockCopy(xm.OptionalData, ByteIndex, ConvertBuffer, 0, 2)
                            ByteIndex += 2
                            ResultAsShort = BitConverter.ToInt16(ConvertBuffer, 0)
                            DataStringList.Add(ResultAsShort.ToString)
                    End Select
                    If (CurrentIndex > -1) Then
                        CurrentIndex += 1
                        If (CurrentIndex < PolledAddressList.Count) Then
                            If (PolledAddressList(CurrentIndex).Address < LastAddress) Then
                                'Calc index for internal requests
                                ByteIndex = (PolledAddressList(CurrentIndex).Address - FirstAddress) * 2
                            End If
                        End If
                    End If
                    If (ByteIndex < DataByteLength) Then
                        NumOfMatches += 1
                    Else
                        AddressInRange = False
                    End If
                End While
            Else
                DataAsString(0) = vbNullString
            End If
        Else
            If (DataByteLength > 0) Then
                ReDim DataAsString(0)
                Dim BitMask As Integer = ParsedResult.BitNumber
                Dim DataIndex As Integer = 0
                Dim Result As Integer

                If (ParsedResult.InternallyRequested) Then
                    'This is an internal subscribe so get the index of the Polled List 
                    For CurrentIndex = 0 To PolledAddressList.Count - 1
                        If (ParsedResult.PLCAddress = PolledAddressList(CurrentIndex).PLCAddress) Then
                            Exit For
                        End If
                    Next
                End If
                For i As Integer = 0 To (ParsedResult.NumberOfElements - 1)
                    If (DataIndex >= xm.OptionalData.Length) Then
                        'Error Condition
                        For z As Integer = 0 To (ParsedResult.NumberOfElements - 1)
                            DataStringList.Add(vbNullString)
                        Next z
                    Else
                        If (ParsedResult.DataTypeModifier = DataTypeModifier.TypeByte) Then
                            Result = xm.OptionalData(i)
                        Else
                            Result = (xm.OptionalData(DataIndex) And (2 ^ BitMask))
                            Result = (Result >> BitMask)
                        End If
                        If (CurrentIndex > -1) Then
                            'Part of subscription - Find it in the Polled List
                            If (CurrentIndex < PolledAddressList.Count) Then
                                If ((ParsedResult.Address + i) = PolledAddressList(CurrentIndex).Address) Then
                                    DataStringList.Add(Result.ToString)
                                    CurrentIndex += 1
                                    For z As Integer = CurrentIndex To PolledAddressList.Count - 1
                                        'Look forward in list for multiply subscriptions to same address
                                        If ((ParsedResult.Address + i) = PolledAddressList(z).Address) Then
                                            DataStringList.Add(Result.ToString)
                                            CurrentIndex += 1
                                        Else
                                            z = PolledAddressList.Count
                                        End If
                                    Next
                                End If
                            End If
                        Else
                            DataStringList.Add(Result.ToString)
                        End If
                    End If
                    If (BitMask < 7) Then
                        BitMask += 1
                    Else
                        BitMask = 0
                        DataIndex += 1
                    End If
                Next i
            Else
                DataAsString(0) = vbNullString
            End If
        End If

        If (xm.DataLength > 0) Then
            'Only if PLC says there is data do we return array of strings
            If (DataStringList.Count > 0) Then
                'Copy the results to string array for return value
                ReDim DataAsString(DataStringList.Count - 1)
                DataStringList.CopyTo(DataAsString)
            End If
            ExtractData = DataAsString
        Else
            ExtractData = Nothing
        End If

    End Function


    Private Function SendData(ByVal data() As Byte) As Integer
        If DLL Is Nothing Then
            DLL = New SNPX_DataLinkLayer
        End If
        'If Not EventHandleAdded Then
        '    AddHandler DLL.DataReceived, AddressOf DataLinkLayer_DataReceived
        '    EventHandleAdded = True
        'End If
        Return DLL.SendData(data)
    End Function

    '************************************************
    '* Convert the message code number into a string
    '* Ref Page 8-3
    '************************************************
    Public Shared Function DecodeMessage(ByVal msgNumber As Integer) As String
        Select Case msgNumber
            Case 0
                DecodeMessage = ""
            Case -2
                Return "Not Acknowledged (NAK)"
            Case -3
                Return "No Response, Check COM Settings"
            Case -4
                Return "Unknown Message from DataLink Layer"
            Case -5
                Return "Invalid Address"
            Case -6
                Return "Could Not Open Com Port"
            Case -7
                Return "No data specified to data link layer"
            Case -8
                Return "No data returned from PLC"
            Case -9
                Return "Failed To Open " & DLL.ComPort
            Case -20
                Return "No Data Returned"
            Case -21
                Return "Received Message invalid BCC"
            Case -30
                Return "Data Value out of Range"

                '*** Errors coming from PLC - GE Fanuc SNPX Major 0Fh 
            Case 3841
                Return "Service request code in an X-Request message is unsupported or invalid at this time.  May occur is a SNP-X session has not been successfully established at the slave device."
            Case 3842
                Return "Insufficient privilege level in the slave PLC CPU for the requested SNP-X service. Password protection at PLC CPU may be preventing the request service."
            Case 3843
                Return "Invalid slave memory type in X-Request message."
            Case 3844
                Return "Invalid slave memory address or range in X-Request message."
            Case 3845
                Return "Invalid data length in X-Request message."
            Case 3846
                Return "X-Buffer data length does not match the service request in X-Request message."
            Case 3847
                Return "Queue full indication from Service Request Processor in slave PLC CPU.  The slave is temporarily unable to complete the service request; the master should try again later."
            Case 3848
                Return "Service Request Processor response exceeds 1000 bytes; the SNP-X slave device cannot return the data in a X-response message."
            Case 3856
                Return "Unexpected Service Request Processor error.  The unexpected SRP error code is saved in the Diagnostic Status Words in the CMM module."
            Case 3861
                Return "Requested service is not permitted in a Broadcast request."
            Case 3872
                Return "Invalid Message Type field in a received X-Request message."
            Case 3873
                Return "Invalid Next Message Type or Next Message Length."
            Case 3874
                Return "Invalid Message Type field in a received X-Buffer message."
            Case 3875
                Return "Invalid Next Message Type field in a received X-Buffer message."
            Case Else
                Return "Unknown Message - " & Hex(msgNumber)
        End Select
    End Function

    Private Sub SNPX_DataLinkLayer_DataReceived()
        'Got a DataPacket with valid BCC so process it

        'GE doesn't use equivalent of TNS of DF1Comm so use the SNPid as the TNS.  In a multi drop situation
        'keep a SNPidList to reference against.  

        '***>>> Do not have a Multi drop enviroment so this has not been tested. <<<****

        'A Response Message should only come in as a result of a Request message.  SNPid is not included in
        'normal Response only in a Response to a X-Attach request.  So index off of SNPidCurrentIndex that is 
        'set at the time of the last Request message.

        Dim lowByte As UShort
        Dim highByte As UShort
        Dim TNSReturned As Integer = SNPidCurrentIndex
        DataPackets(TNSReturned) = DLL.DataPacket
        Dim xResponse As XResponse = DLL.ResponseMessage 'Get Copy From DLL

        '**************************************************************
        '* Only extract and send back if this repsonse contained data
        '**************************************************************
        Dim ReturnedBytes As Integer = DataPackets(TNSReturned).Count

        'EndOfBlock should be 6 bytes from end.  Make sure there are more than 14 bytes - Smallest size XResponse
        If (ReturnedBytes >= 14) Then
            If (DataPackets(TNSReturned)(ReturnedBytes - 6) = MessageType.EndOfBlock) Then
                xResponse = DLL.ResponseMessage 'Get Copy From DLL
                ReDim xResponse.OptionalData(0)
                ReDim Preserve xResponse.SNPid(0 To 7)
                If ((ReturnedBytes = 24) And (DataPackets(TNSReturned)(10))) Then
                    For iByte As Integer = 0 To (ReturnedBytes - 1)
                        '24 Bytes = X-Attach Response
                        Select Case iByte
                            Case 0
                                xResponse.Start = DataPackets(TNSReturned)(iByte)
                            Case 1
                                xResponse.MessageType = DataPackets(TNSReturned)(iByte)
                            Case 2 To 9 'SNP id
                                xResponse.SNPid(iByte - 2) = DataPackets(TNSReturned)(iByte)
                            Case 10
                                xResponse.ResponseCode = DataPackets(TNSReturned)(iByte)
                            Case (ReturnedBytes - 6)
                                xResponse.EndOfBlock = DataPackets(TNSReturned)(iByte)
                            Case (ReturnedBytes - 5)
                                xResponse.NextMessageType = DataPackets(TNSReturned)(iByte)
                            Case (ReturnedBytes - 4)
                                xResponse.NextMessageLength = DataPackets(TNSReturned)(iByte)
                            Case (ReturnedBytes - 3)
                                xResponse.NextMessageLength = DataPackets(TNSReturned)(iByte)
                            Case (ReturnedBytes - 2)
                                xResponse.NotUsed = DataPackets(TNSReturned)(iByte)
                            Case (ReturnedBytes - 1)
                                xResponse.BCC = DataPackets(TNSReturned)(iByte)
                            Case Else 'It's Data
                                ReDim Preserve xResponse.OptionalData(0 To (iByte - 11))
                                xResponse.OptionalData(iByte - 11) = DataPackets(TNSReturned)(iByte)
                        End Select
                    Next
                Else
                    'It's either a X-Read/X-write or Intermediate
                    For iByte As Integer = 0 To (ReturnedBytes - 1)
                        Select Case iByte
                            Case 0
                                xResponse.Start = DataPackets(TNSReturned)(iByte)
                            Case 1
                                xResponse.MessageType = DataPackets(TNSReturned)(iByte)
                            Case 2
                                xResponse.ResponseCode = DataPackets(TNSReturned)(iByte)
                            Case 3
                                xResponse.StatusWord = DataPackets(TNSReturned)(iByte)
                            Case 4
                                lowByte = xResponse.StatusWord
                                highByte = DataPackets(TNSReturned)(iByte)
                                highByte = (highByte << 8)
                                xResponse.StatusWord = (lowByte Or highByte)
                            Case 5
                                xResponse.MajorError = DataPackets(TNSReturned)(iByte)
                            Case 6
                                xResponse.MinorError = DataPackets(TNSReturned)(iByte)
                                lowByte = xResponse.MinorError
                                highByte = xResponse.MajorError
                                highByte = (highByte << 8)
                                xResponse.ErrorCode = (lowByte Or highByte)
                            Case 7
                                xResponse.DataLength = DataPackets(TNSReturned)(iByte)
                            Case 8
                                lowByte = xResponse.DataLength
                                highByte = DataPackets(TNSReturned)(iByte)
                                highByte = (highByte << 8)
                                xResponse.DataLength = (lowByte Or highByte)
                            Case (ReturnedBytes - 6)
                                xResponse.EndOfBlock = DataPackets(TNSReturned)(iByte)
                            Case (ReturnedBytes - 5)
                                xResponse.NextMessageType = DataPackets(TNSReturned)(iByte)
                            Case (ReturnedBytes - 4)
                                xResponse.NextMessageLength = DataPackets(TNSReturned)(iByte)
                            Case (ReturnedBytes - 3)
                                lowByte = xResponse.NextMessageLength
                                highByte = DataPackets(TNSReturned)(iByte)
                                highByte = (highByte << 8)
                                xResponse.NextMessageLength = (lowByte Or highByte)
                            Case (ReturnedBytes - 2)
                                xResponse.NotUsed = DataPackets(TNSReturned)(iByte)
                            Case (ReturnedBytes - 1)
                                xResponse.BCC = DataPackets(TNSReturned)(iByte)
                            Case Else 'It's Data
                                ReDim Preserve xResponse.OptionalData(0 To (iByte - 9))
                                xResponse.OptionalData(iByte - 9) = DataPackets(TNSReturned)(iByte)
                        End Select
                    Next
                    'Get the last Request SNPid it wasn't included in Response
                    xResponse.SNPid = ResponseMessage.SNPid
                End If
                Dim RunMode = xResponse.StatusWord >> 12
                If ((xResponse.ResponseCode = ResponseCode.XRead) And (RunMode > 1)) Then
                    'Zero out data if PLC is not in a run mode.  PLC will return incorrect data not reflecting actual state of addresses when stopped.
                    For z As Integer = 0 To xResponse.OptionalData.Length - 1
                        xResponse.OptionalData(z) = 0
                    Next
                End If
                'Update local and DLL global copies
                ResponseMessage = xResponse
                DLL.ResponseMessage = xResponse
                PLCAddressByTNS(TNSReturned).ResponseCode = ResponseMessage.ResponseCode
                Responded(TNSReturned) = True
            Else
                xResponse.ResponseCode = ResponseCode.Invalid
            End If

            PLCAddressByTNS(TNSReturned).RequestActive = False

            If (xResponse.ResponseCode = ResponseCode.XRead) Then
                If (xResponse.ErrorCode <> 0) Then
                    LastSNPXErrorCode = xResponse.ErrorCode
                    CurrentSNPXRequestInError = True
                    LastSNPXErrorAddress = PLCAddressByTNS(TNSReturned).PLCAddress
                Else
                    CurrentSNPXRequestInError = False
                End If
                CurrentSNPXStatusCode = xResponse.StatusWord
                Dim CurrentPA As ParsedDataAddress = PLCAddressByTNS(TNSReturned)
                Dim d() As String = ExtractData(CurrentPA, xResponse)
                Dim pa As ParsedDataAddress = PLCAddressByTNS(TNSReturned)
                LastReadWasWord = CurrentPA.DataIsWord
                If Not CurrentPA.InternallyRequested Then
                    If Not DisableEvent Then
                        Dim x As New SNPXCommEventArgs(d, CurrentPA.PLCAddress)
                        If m_SynchronizingObject IsNot Nothing Then
                            Dim Parameters() As Object = {Me, x}
                            m_SynchronizingObject.BeginInvoke(drsd, Parameters)
                        Else
                            'RaiseEvent DataReceived(Me, System.EventArgs.Empty)
                            RaiseEvent DataReceived(Me, x)
                        End If
                    End If
                Else

                    'Must be a read via subscribe
                    If (d Is Nothing) Then
                        Throw New SNPXException("No Read Data to return - Most likely an Invalid Address")
                    End If
                    For i As Integer = 0 To (PolledAddressList.Count - 1)
                        If (PolledAddressList(i).PLCAddress = CurrentPA.PLCAddress) Then
                            If (d.Length > 1) Then
                                For pi As Integer = i To PolledAddressList.Count - 1
                                    'Index from this starting PLC Address while within the same segment types
                                    If (PolledAddressList(pi).Segment = CurrentPA.Segment) Then
                                        'Cycle through all the subscriptions within the extracted data
                                        For z As Integer = 0 To d.Length - 1
                                            If (d(z) <> PolledAddressList(pi + z).LastValue) Then
                                                Dim Ed() As String = {d(z)}
                                                m_SynchronizingObject.BeginInvoke(PolledAddressList(pi + z).dlgCallBack, CObj(Ed))
                                                PolledAddressList(pi + z).LastValue = Ed(0)
                                            End If
                                        Next
                                        pi = PolledAddressList.Count
                                    Else
                                        pi = PolledAddressList.Count
                                    End If
                                Next pi
                            Else
                                If (d(0) <> PolledAddressList(i).LastValue) Then
                                    PolledAddressList(i).LastValue = d(0)
                                    m_SynchronizingObject.BeginInvoke(PolledAddressList(i).dlgCallBack, CObj(d))
                                End If
                            End If
                        End If
                    Next i
                    Responded(TNSReturned) = False
                End If
            End If
        End If
    End Sub

    '******************************************************************
    '* This is called when a message instruction was sent from the PLC
    '******************************************************************
    Private Sub SNPX_DataLink_UnsolictedMessageRcvd()

    End Sub

    '****************************************************************************
    '* This is required to sync the event back to the parent form's main thread
    '****************************************************************************
    Dim drsd As EventHandler = AddressOf DataReceivedSync
    'Delegate Sub DataReceivedSyncDel(ByVal sender As Object, ByVal e As EventArgs)
    Private Sub DataReceivedSync(ByVal sender As Object, ByVal e As SNPXCommEventArgs)
        RaiseEvent DataReceived(sender, e)
    End Sub

    Private Sub UnsolictedMessageRcvdSync(ByVal sender As Object, ByVal e As EventArgs)
        RaiseEvent UnsolictedMessageRcvd(sender, e)
    End Sub
#End Region

#Region "SNPX Message"

    Public CurrentSNPXRequestInError As Boolean
    Public LastSNPXErrorCode As Integer
    Public LastSNPXErrorAddress As String
    Public CurrentSNPXStatusCode As Integer

    Public Enum MessageType As Byte
        SNP_X = CByte(&H58)
        Intermediate = CByte(&H78)
        DataBuffer = CByte(&H54)
        Start = CByte(&H1B)
        EndOfBlock = CByte(&H17)
    End Enum

    Public Enum RequestCode As Byte
        XAttach = CByte(&H0)
        XRead = CByte(&H1)
        XWrite = CByte(&H2)
    End Enum

    Public Enum ResponseCode As Byte
        XAttach = CByte(&H80)
        XRead = CByte(&H81)
        XWrite = CByte(&H82)
        Invalid = CByte(&HFF)
    End Enum

    ' SNP-X Message Structure Types
    Public Structure XRequest
        'Header
        Dim Start As Byte
        Dim MessageType As Byte
        'Command Data
        Dim SNPid() As Byte
        Dim RequestCode As Byte
        Dim Data() As Byte
        'Trailer
        Dim EndOfBlock As Byte
        Dim NextMessageType As Byte
        Dim NextMessageLength As UShort
        Dim NotUsed As Byte
        Dim BCC As Byte
    End Structure

    Public Structure XResponse
        'Header
        Dim Start As Byte
        Dim MessageType As Byte
        'Command Data
        Dim ResponseCode As Byte
        Dim StatusWord As UShort
        Dim MajorError As Byte
        Dim MinorError As Byte
        Dim DataLength As UShort
        Dim OptionalData() As Byte 'Array size = DataLength
        'Trailer
        Dim EndOfBlock As Byte
        Dim NextMessageType As Byte
        Dim NextMessageLength As UShort
        Dim NotUsed As Byte
        Dim BCC As Byte
        'Additional not part of GE defined structure
        Dim SNPid() As Byte
        Dim ErrorCode As UShort
    End Structure

    Public Structure XBuffer
        'Header
        Dim Start As Byte
        Dim MessageType As Byte
        'Command Data
        Dim Data() As Byte 'Array size = Previous XResponse.NextMessageLength - 8 (Sizeof of other structure members) 
        'Trailer
        Dim EndOfBlock As Byte
        Dim NextMessageType As Byte
        Dim NextMessageLength As UShort
        Dim NotUsed As Byte
        Dim BCC As Byte
    End Structure

    Public Function SendXAttach() As Integer
        'This function started out as a general purpose SendXMessage but to integrate with AdvancedHMI PrefixAndSend was created and that is used for all other messages

        Dim xr As XRequest
        Dim Max As Integer
        Dim Index As Integer
        Dim OutBuff(0 To 22) As Byte 'Command Size - CheckSum Byte

        'Dim the structure array to proper sizes
        ReDim xr.SNPid(0 To 7)
        ReDim xr.Data(0 To 6)

        'Fill in common items
        xr.Start = MessageType.Start
        xr.MessageType = MessageType.SNP_X
        xr.EndOfBlock = MessageType.EndOfBlock

        Max = Len(Me.SNPid)
        If Max > 7 Then
            Max = 7
        End If
        If (Max > 0) Then
            For Index = 0 To (Max - 1)
                xr.SNPid(Index) = CByte(Asc(Me.SNPid(Index)))
            Next
        End If
        '***************************************************
        '* Check for SNPid and insert into list when needed.  
        '* RequestMessage.SNPid has the current id
        '***************************************************
        Dim SNPid(0 To 7) As Char
        xr.SNPid.CopyTo(SNPid, 0)
        Dim i As Integer = SNPidIndex(SNPid)
        If (i = -1) Then 'Not in list so Add SNPid
            SNPidIndex(SNPid) = 0 'This value doesn't matter it forces an update based on SNPid
            SNPidCount += 1
        End If

        Dim Request As Integer = RequestCode.XAttach
        'Type Specific structure build
        Select Case Request
            Case RequestCode.XAttach
                'Set Data fields to 0
                'For Index = 0 To 6
                '    xr.Data(Index) = 0
                'Next
                Array.Clear(xr.Data, 0, 7)
                xr.NextMessageType = 0
                xr.NextMessageLength = 0
                xr.NotUsed = 0
                'X.BCC = 0
                'Case RequestCode.XRead
                '    xr.Data(0) = SegmentType
                '    xr.Data(1) = Address And &HFF
                '    xr.Data(2) = Address / 256
                '    xr.Data(3) = DataLength And &HFF
                '    xr.Data(4) = DataLength / 256
                '    xr.Data(5) = xr.Data(6) = 0
                '    xr.NextMessageType = 0
                '    xr.NextMessageLength = 0
                '    xr.NotUsed = 0
                '    'X.BCC = 0
                'Case RequestCode.XWrite
                '    xr.Data(0) = SegmentType
                '    xr.Data(1) = Address And &HFF
                '    xr.Data(2) = Address / 256
                '    xr.Data(3) = DataLength And &HFF
                '    xr.Data(4) = DataLength / 256
                '    ' 2 Bytes or Less
                '    xr.Data(5) = Data(0)
                '    xr.Data(6) = Data(1)
                '    xr.NextMessageType = 0
                '    xr.NextMessageLength = 0
                '    xr.NotUsed = 0
                '    'X.BCC = 0
            Case Else
                'Invalid Request Type
                Return (-1)
        End Select

        'Build Buffer
        OutBuff(0) = xr.Start
        OutBuff(1) = xr.MessageType
        Buffer.BlockCopy(xr.SNPid, 0, OutBuff, 2, 8)
        OutBuff(10) = xr.RequestCode
        Buffer.BlockCopy(xr.Data, 0, OutBuff, 11, 7)
        OutBuff(18) = xr.EndOfBlock
        OutBuff(19) = xr.NextMessageType
        OutBuff(20) = xr.NextMessageLength And &HFF
        OutBuff(21) = xr.NextMessageLength / 256
        OutBuff(22) = xr.NotUsed
        RequestMessage = xr 'Save copy of Request minus the BCC
        DLL.RequestMessage = xr 'Also save it to the DLL copy
        'SendData will add the BlockCheckCode
        Return SendData(OutBuff)

    End Function


    Public Function SNPX_Attach() As Integer
        Dim DataBuff(1) As Byte
        Dim ReturnVal As Integer
        If (0 = OpenComm()) Then
            ReturnVal = DLL.SendSerialBreak(LongBreakDuration)
            If (ReturnVal = 0) Then
                ReturnVal = SendXAttach()
                System.Threading.Thread.Sleep(T4Time)
            End If
        Else
            m_Attached = False
            Return Err.Number
        End If
        m_Attached = True
        Return ReturnVal
    End Function

    Public Function DecodeSegment(ByVal segment As String, ByVal bit As Boolean) As Byte
        Dim value As Byte
        Select Case segment
            Case "R"
                value = &H8
            Case "AI"
                value = &HA
            Case "AQ"
                value = &HC
            Case "I"
                If bit Then
                    value = &H46
                Else
                    value = &H10
                End If
            Case "Q"
                If bit Then
                    value = &H48
                Else
                    value = &H12
                End If
            Case "T"
                If bit Then
                    value = &H4A
                Else
                    value = &H14
                End If
            Case "M"
                If bit Then
                    value = &H4C
                Else
                    value = &H16
                End If
            Case "S"
                If bit Then
                    value = &H54
                Else
                    value = &H1E
                End If
            Case "SA"
                If bit Then
                    value = &H4E
                Else
                    value = &H18
                End If
            Case "SB"
                If bit Then
                    value = &H50
                Else
                    value = &H1A
                End If
            Case "SC"
                If bit Then
                    value = &H52
                Else
                    value = &H1C
                End If
            Case "G"
                If bit Then
                    value = &H56
                Else
                    value = &H38
                End If
            Case Else
                value = 0
        End Select
        DecodeSegment = value
    End Function

    Public Function DecodeErrorCode(ByRef ErrorCode As Integer) As String
        'Decode passed ErrorCode and return string
        Return DecodeMessage(ErrorCode)
    End Function


    Public Function DecodeErrorCode() As String
        'Overload default to Global Error Code
        Return DecodeErrorCode(LastSNPXErrorCode)
    End Function


    Public Function DecodeStatusCode(ByRef StatusCode As Integer) As String
        'Decode passed StatusCode returning multi-line string with status details
        Dim PLCState As Integer
        Dim ReturnString As String

        PLCState = StatusCode And &HF000
        PLCState = PLCState >> 12
        Select Case PLCState
            Case 0
                ReturnString = "Run I/O enabled."
            Case 1
                ReturnString = "Run I/O disabled."
            Case 2
                ReturnString = "Stop I/O disabled."
            Case 3
                ReturnString = "CPU stop faulted."
            Case 4
                ReturnString = "CPU halted."
            Case 5
                ReturnString = "CPU Suspended."
            Case 6
                ReturnString = "Stop I/O enabled."
            Case Else
                ReturnString = "Invalid PLC State."
        End Select

        If (StatusCode And 1) Then
            ReturnString = ReturnString + vbCrLf + "Constant Sweep Value exceeded."
        Else
            ReturnString = ReturnString + vbCrLf + "No oversweep condition exists."
        End If
        If (StatusCode And 2) Then
            ReturnString = ReturnString + vbCrLf + "Constant Sweep Mode active."
        Else
            ReturnString = ReturnString + vbCrLf + "Constant Sweep Mode is not active."
        End If
        If (StatusCode And 4) Then
            ReturnString = ReturnString + vbCrLf + "PLC fault table has changed since last read."
        Else
            ReturnString = ReturnString + vbCrLf + "PLC fault table is unchanged since last read."
        End If
        If (StatusCode And 8) Then
            ReturnString = ReturnString + vbCrLf + "I/O fault table has changed since last read."
        Else
            ReturnString = ReturnString + vbCrLf + "I/O fault table is unchanged since last read."
        End If
        If (StatusCode And 16) Then
            ReturnString = ReturnString + vbCrLf + "One or more fault entries in PLC fault table."
        Else
            ReturnString = ReturnString + vbCrLf + "PLC fault table is empty."
        End If
        If (StatusCode And 32) Then
            ReturnString = ReturnString + vbCrLf + "One or more fault entries in I/O fault table."
        Else
            ReturnString = ReturnString + vbCrLf + "I/O fault table is empty."
        End If
        If (StatusCode And 64) Then
            ReturnString = ReturnString + vbCrLf + "Programmer attachment found."
        Else
            ReturnString = ReturnString + vbCrLf + "No Programmer attachment found."
        End If
        If (StatusCode And 128) Then
            ReturnString = ReturnString + vbCrLf + "Front panel Enable/Disable - Outputs disabled."
        Else
            ReturnString = ReturnString + vbCrLf + "Front panel Enable/Disable - Outputs enabled."
        End If
        If (StatusCode And 256) Then
            ReturnString = ReturnString + vbCrLf + "Front panel Run/Stop - In Run."
        Else
            ReturnString = ReturnString + vbCrLf + "Front panel Run/Stop - In Stop."
        End If
        If (StatusCode And 512) Then
            ReturnString = ReturnString + vbCrLf + "OEM protection in effect."
        Else
            ReturnString = ReturnString + vbCrLf + "No OEM protection."
        End If
        Return ReturnString
    End Function

    Public Function DecodeStatusCode() As String
        'Overload default to Global StatusCode
        Return DecodeStatusCode(CurrentSNPXStatusCode)
    End Function


    Private m_SNPidIndex As Integer
    Public Property SNPidIndex(ByVal SNPid As String) As Integer
        Get
            For i As Integer = 0 To (SNPidList.Length - 1)
                If (0 = SNPidList(i).CompareTo(SNPid)) Then
                    SNPidCurrentIndex = i
                    Me.SNPid = SNPid
                    Return i
                End If
            Next i
            Return -1
        End Get
        Set(ByVal value As Integer)
            Dim i As Integer
            Dim FoundIt As Boolean = False
            For i = 0 To (SNPidList.Length - 1)
                If (SNPidList(i) = SNPid) Then
                    FoundIt = True
                End If
            Next i
            If (Not FoundIt) Then
                'Add to the list
                ReDim Preserve SNPidList(0 To i)
                SNPidList(i) = SNPid
                SNPidCurrentIndex = i
                'Also add to other arrays that are base on TNS (SNPid in this case)
                ReDim Preserve DataPackets(0 To i)
                ReDim Preserve Responded(0 To i)
                ReDim Preserve PLCAddressByTNS(0 To i)
            End If
            Me.SNPid = SNPid
        End Set
    End Property

#End Region

End Class

Public Delegate Sub SNPXCommEventHandler(ByVal sender As Object, ByVal e As SNPXCommEventArgs)
Public Class SNPXCommEventArgs
    Inherits EventArgs
    'Constructor.
    Public Sub New(ByVal Values() As String, ByVal PLCAddress As String)
        m_Values = Values
        m_PLCAddress = PLCAddress
    End Sub

    Private m_Values() As String
    Public ReadOnly Property Values() As String()
        Get
            Return m_Values
        End Get
        'Set(ByVal value As String())
        '    m_Values = value
        'End Set
    End Property

    Private m_PLCAddress As String
    Public ReadOnly Property PLCAddress() As String
        Get
            Return m_PLCAddress
        End Get
        'Set(ByVal value As String)
        '    m_PLCAddress = value
        'End Set
    End Property
End Class


Public Class SNPX_DataLinkLayer
#Region "Data Link Layer"
    '**************************************************************************************************************
    '**************************************************************************************************************
    '**************************************************************************************************************
    '*****                                  DATA LINK LAYER SECTION
    '**************************************************************************************************************
    '**************************************************************************************************************
    Friend LastResponseWasNAK As Boolean
    '* This is used to help problems that come from transmissions errors when using a USB converter
    Private SleepDelay As Integer = 0
    Public DataPacket As System.Collections.ObjectModel.Collection(Of Byte)

    Private WithEvents SerialPort As New System.IO.Ports.SerialPort

    Public RequestMessage As SNPXComm.XRequest
    Public ResponseMessage As SNPXComm.XResponse
    Public BufferMessage As SNPXComm.XBuffer

#Region "Properties"
    Private m_BaudRate As Integer = 19200
    Public Property BaudRate() As Integer
        Get
            Return m_BaudRate
        End Get
        Set(ByVal value As Integer)
            If value <> m_BaudRate Then CloseComms()
            m_BaudRate = value
        End Set
    End Property

    Private m_ComPort As String = "COM1"
    Public Property ComPort() As String
        Get
            Return m_ComPort
        End Get
        Set(ByVal value As String)
            If value <> m_ComPort Then CloseComms()
            m_ComPort = value
        End Set
    End Property

    Private m_Parity As System.IO.Ports.Parity = IO.Ports.Parity.Odd
    Public Property Parity() As System.IO.Ports.Parity
        Get
            Return m_Parity
        End Get
        Set(ByVal value As System.IO.Ports.Parity)
            m_Parity = value
        End Set
    End Property

    Private m_StopBits As System.IO.Ports.StopBits = IO.Ports.StopBits.One
    Public Property StopBits() As System.IO.Ports.StopBits
        Get
            Return m_StopBits
        End Get
        Set(ByVal value As System.IO.Ports.StopBits)
            m_StopBits = value
        End Set
    End Property

    Private m_SNPid As String = ""
    Public Property SNPid() As String
        Get
            Return m_SNPid
        End Get
        Set(ByVal value As String)
            m_SNPid = value
        End Set
    End Property
#End Region

    '***************************
    '* Calculate a BCC
    '***************************
    Private Shared Function CalculateBCC(ByVal DataInput() As Byte) As Byte
        Dim BlockCheck As Integer = 0
        'For i As Integer = 0 To DataInput.Length - 1
        '    BlockCheck = BlockCheck Xor DataInput(i)
        '    BlockCheck = (BlockCheck << 1)
        '    BlockCheck = BlockCheck + ((BlockCheck >> 8) And &H1)
        '    BlockCheck = BlockCheck And &HFF
        '    If ((DataInput(10) = 2) And (i = 16)) Then
        '        DataInput(10) = DataInput(10)
        '    End If
        'Next
        For i As Integer = 0 To DataInput.Length - 1
            BlockCheck = BlockCheck Xor DataInput(i)
            BlockCheck = 2 * BlockCheck
            If (BlockCheck > &HFF) Then
                BlockCheck = BlockCheck Or 1
                BlockCheck = BlockCheck And &HFF
            End If
            If ((DataInput(10) = 2) And (i = 16)) Then
                DataInput(10) = DataInput(10)
            End If
        Next
        CalculateBCC = BlockCheck
    End Function
    '* Overload
    Private Shared Function CalculateBCC(ByVal DataInput As System.Collections.ObjectModel.Collection(Of Byte)) As Byte
        Dim BlockCheck As Integer = 0
        For i As Integer = 0 To DataInput.Count - 1
            BlockCheck = BlockCheck Xor DataInput(i)
            BlockCheck = 2 * BlockCheck
            If (BlockCheck > &HFF) Then
                BlockCheck = BlockCheck Or 1
                BlockCheck = BlockCheck And &HFF
            End If
        Next
        CalculateBCC = BlockCheck
    End Function


    '******************************************
    '* Handle Data Received On The Serial Port
    '******************************************
    Private LastByte As Byte
    Private PacketStarted As Boolean
    Private PacketEnded As Boolean
    'Private NodeChecked As Boolean
    Dim b As Byte
    Private ETXPosition As Integer
    Private ReceivedDataPacketBuilder As New System.Collections.ObjectModel.Collection(Of Byte)
    Private ReceivedDataPacket As New System.Collections.ObjectModel.Collection(Of Byte)
    Private WithEvents FrameGapTimer As New System.Timers.Timer
    Private ReceivedByteCount As Integer
    Private Sub SerialPort_DataReceived(ByVal sender As Object, ByVal e As System.IO.Ports.SerialDataReceivedEventArgs) Handles SerialPort.DataReceived
        '**************************************************************
        '* Handle reading in a complete SNP-X message structure then call on Process...

        Dim BytesToRead As Integer = SerialPort.BytesToRead

        Dim BytesRead(BytesToRead - 1) As Byte
        BytesToRead = SerialPort.Read(BytesRead, 0, BytesToRead)

        '* Reset the frame/packet timer
        FrameGapTimer.Enabled = False

        Dim i As Integer = 0
        Static Dim j As Integer
        Static Dim FoundStart As Boolean = False
        Static Dim FoundTrailer As Boolean = False

        While i < BytesToRead
            b = BytesRead(i)
            'Look for a StartByte and valid MessageType
            If ((LastByte = SNPXComm.MessageType.Start) And ((b = SNPXComm.MessageType.SNP_X) Or (b = SNPXComm.MessageType.Intermediate) Or (b = SNPXComm.MessageType.DataBuffer))) Then
                FoundStart = True
                j = 0
                ReceivedDataPacketBuilder.Clear()
                ReceivedDataPacketBuilder.Add(LastByte)
            End If
            '**************************************************************
            '* Do not start capturing chars until start of packet received
            '**************************************************************
            If (FoundStart) Then
                ReceivedDataPacketBuilder.Add(b)
                'Look for Trailer that will be an EndOfBlock followed by 4 bytes of 00
                If (b = SNPXComm.MessageType.EndOfBlock) Then
                    j = ReceivedDataPacketBuilder.Count 'Save the Packet Count
                End If
                'Now monitor the 4 bytes after the EndOfBlock
                If ((j > 0) And (ReceivedDataPacketBuilder.Count > j)) Then
                    If ((j + 5) = ReceivedDataPacketBuilder.Count) Then
                        FoundTrailer = True 'Last 4 Bytes were 0
                    End If
                    If (Not b = 0) Then
                        j = 0
                    End If
                End If
                If (FoundStart And FoundTrailer) Then
                    'Last Byte added to Packet was the BCC
                    FoundStart = False
                    FoundTrailer = False
                    ReceivedDataPacket.Clear()
                    For x As Integer = 0 To ReceivedDataPacketBuilder.Count - 1
                        ReceivedDataPacket.Add(ReceivedDataPacketBuilder(x))
                    Next
                    '* Make sure last packet was processed before clearing it
                    ProcessReceivedData()
                    ReceivedDataPacketBuilder.Clear()
                    Acknowledged = True
                End If
            End If
            LastByte = b
            i += 1
        End While
        FrameGapTimer.Enabled = True
    End Sub

    Public Event DataReceived As EventHandler
    Private WithEvents AsyncProcessData As New Timer
    Private WithEvents BackgroundWorker1 As New System.ComponentModel.BackgroundWorker

    '*******************************************************************
    '* When a complete packet is received from the PLC, this is called
    '*******************************************************************
    'Private Sub ProcessData(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AsyncProcessData.Tick
    Private Sub ProcessReceivedData() Handles BackgroundWorker1.DoWork

        '*********************************************************************
        '* Validate checksum and raise DataReceived Event to process and parse
        '*********************************************************************
        'AsyncProcessData.Enabled = False

        '* Get the Checksum that came back from the PLC
        Dim CheckSumResult As Byte
        CheckSumResult = ReceivedDataPacket(ReceivedDataPacket.Count - 1)
        '****************************
        '* validate CRC received
        '****************************
        If DataPacket IsNot Nothing Then
            DataPacket.Clear()
        Else
            DataPacket = New System.Collections.ObjectModel.Collection(Of Byte)
        End If

        For i As Integer = 0 To ReceivedDataPacket.Count - 2 'Skip BCC
            DataPacket.Add(ReceivedDataPacket(i))
        Next
        ReceivedDataPacket.Clear()


        '* Calculate the checksum for the received data
        Dim CheckSumCalc As Byte
        CheckSumCalc = CalculateBCC(DataPacket)

        '***************************************************************************
        '* Send back an response to indicate whether data was received successfully
        '***************************************************************************
        If CheckSumResult = CheckSumCalc Then
            'Put Back in the BCC
            DataPacket.Add(CheckSumCalc)
            'Try
            'Responded = True

            RaiseEvent DataReceived(Me, System.EventArgs.Empty)
            'Catch ex As Exception
            '    Dim x As Int16 = 0
            '    Throw ex
            'End Try

            '* Keep this in case the PLC requests with ENQ
            LastResponseWasNAK = False

            'Responded(xTNS) = True
            '* Keep this in case the PLC requests with ENQ
        Else
            '* Slow down comms if multiple Checksum failures - helps with USB converter
            If SleepDelay < 50 Then
                If LastResponseWasNAK Then
                    SleepDelay += 5
                End If
            End If
            LastResponseWasNAK = True
        End If

    End Sub

    ' Opens the comm port to start communications
    Public Function OpenComms() As Integer
        '*****************************************
        '* Open serial port if not already opened
        '*****************************************
        If Not SerialPort.IsOpen Then
            If m_BaudRate > 0 Then
                Try
                    SerialPort.BaudRate = m_BaudRate
                    SerialPort.PortName = m_ComPort
                    SerialPort.Parity = m_Parity
                    SerialPort.StopBits = m_StopBits
                    SerialPort.ReceivedBytesThreshold = 1
                    SerialPort.Open()
                Catch ex As Exception
                    If SerialPort.IsOpen Then SerialPort.Close()
                    Return -9
                    'Throw New MfgControl.AdvancedHMI.Drivers.PLCDriverException("Failed To Open " & SerialPort.PortName & ". " & ex.Message)
                End Try
            End If
            Try
                SerialPort.DiscardInBuffer()
            Catch ex As Exception
            End Try
        End If
        Return 0
    End Function

    '*******************************************************************
    '* Send Data - this is the key entry used by the application layer
    '* A command stream in the form a a list of bytes are passed
    '* to this method. Protocol commands are then attached and
    '* then sent to the serial port
    '*******************************************************************
    Private Acknowledged As Boolean
    Private NotAcknowledged As Boolean
    Private AckWaitTicks As Integer
    Private MaxSendRetries As Integer = 1
    Dim MaxTicks As Integer = 50 'ticks per second
    Public Function SendData(ByVal data() As Byte) As Integer
        '* A USB converer may need this
        'System.Threading.Thread.Sleep(50)

        '* Make sure there is data to send
        If data.Length < 1 Then Return -7

        If Not SerialPort.IsOpen Then
            Dim OpenResult As Integer = OpenComms()
            If OpenResult <> 0 Then
                CloseComms()
                Return OpenResult
            End If
        End If
        '***************************************
        '* Calculate CheckSum of raw data
        '***************************************
        Dim CheckSumCalc As UInt16
        CheckSumCalc = CalculateBCC(data)

        Dim ByteCount As Integer = data.Length
        '*********************************
        '* Attach Checksum
        '*********************************
        Dim BytesToSend(0 To ByteCount) As Byte

        data.CopyTo(BytesToSend, 0)
        BytesToSend(ByteCount) = CheckSumCalc
        '*********************************************
        '* Send the data and retry 3 times if failed
        '*********************************************
        '* Prepare for response and retries
        Dim Retries As Integer = 0

        'GE Fanuc doesn't send back ACK 
        NotAcknowledged = True
        Acknowledged = False

        While Not Acknowledged And Retries < MaxSendRetries
            '* Reset the response for retries
            Acknowledged = False
            NotAcknowledged = False
            '*******************************************************
            '* The stream of data is complete, send it now
            '* For those who want examples of data streams, put a
            '*  break point here nd watch the BytesToSend variable
            '*******************************************************
            SerialPort.Write(BytesToSend, 0, BytesToSend.Length)

            '* Wait for response of a 1 second (50*20) timeout
            '* We only wait need to wait for an ACK
            '*  The PrefixAndSend Method will continue to wait for the data
            '* NOTE : THERE IS A SleepDelay that auto increments that may make this wait longer.
            AckWaitTicks = 0
            While (Not Acknowledged And Not NotAcknowledged) And AckWaitTicks < MaxTicks
                ' Speed Up response making lockups less frequent
                System.Threading.Thread.Sleep(2 + SleepDelay)
                AckWaitTicks += 1
                'If Response = ResponseTypes.Enquire Then Response = ResponseTypes.NoResponse
            End While
            Retries += 1
        End While
        '**************************************
        '* Return a code indicating the status
        '**************************************
        If Acknowledged Then
            Return 0
        ElseIf NotAcknowledged Then
            Return -2  '* Not Acknowledged
        Else
            Return -3  '* No Response
        End If
    End Function

    'Closes the comm port
    Public Sub CloseComms()
        If SerialPort IsNot Nothing AndAlso SerialPort.IsOpen Then
            'Try
            SerialPort.DiscardInBuffer()
            SerialPort.Close()
            'Catch ex As Exception
            'End Try
        End If
    End Sub

    Public Function SendSerialBreak(ByVal BreakDurationMS As Integer) As Integer
        '   Flush comm buffers and set break state, then wait specified milliseconds
        If (SerialPort.IsOpen) Then
            Try
                SerialPort.DiscardInBuffer()
                SerialPort.DiscardOutBuffer()
                SerialPort.BreakState = True
                System.Threading.Thread.Sleep(BreakDurationMS)
                SerialPort.BreakState = False
                SendSerialBreak = 0
            Catch ex As Exception
                SendSerialBreak = -1
            End Try
        Else
            SendSerialBreak = 1
        End If
    End Function

    '***********************************************
    '* Clear the buffer on a framing error
    '* This is an indication of incorrect baud rate
    '* If PLC is in DH485 mode, serial port throws
    '* an exception without this
    '***********************************************
    Private Sub SerialPort_ErrorReceived(ByVal sender As Object, ByVal e As System.IO.Ports.SerialErrorReceivedEventArgs) Handles SerialPort.ErrorReceived
        If e.EventType = IO.Ports.SerialError.Frame Then
            SerialPort.DiscardInBuffer()
        End If
    End Sub

#End Region

End Class


' This UITypeEditor can be associated with Int32, Double and Single
' properties to provide a design-mode angle selection interface.
<System.Security.Permissions.PermissionSetAttribute(System.Security.Permissions.SecurityAction.Demand, Name:="FullTrust")> _
Public Class SNPXBaudRateEditor
    Inherits System.Drawing.Design.UITypeEditor

    Public Sub New()
    End Sub

    ' Indicates whether the UITypeEditor provides a form-based (modal) dialog, 
    ' drop down dialog, or no UI outside of the properties window.
    Public Overloads Overrides Function GetEditStyle(ByVal context As System.ComponentModel.ITypeDescriptorContext) As System.Drawing.Design.UITypeEditorEditStyle
        Return UITypeEditorEditStyle.DropDown
    End Function

    ' Displays the UI for value selection.
    Dim edSvc As IWindowsFormsEditorService
    Private WithEvents lb As ListBox
    Public Overloads Overrides Function EditValue(ByVal context As System.ComponentModel.ITypeDescriptorContext, ByVal provider As System.IServiceProvider, ByVal value As Object) As Object
        ' Uses the IWindowsFormsEditorService to display a 
        ' drop-down UI in the Properties window.
        edSvc = CType(provider.GetService(GetType(IWindowsFormsEditorService)), IWindowsFormsEditorService)
        lb = New ListBox
        If (edSvc IsNot Nothing) Then
            'lb.Items.Add("AUTO")
            lb.Items.Add("19200")
            lb.Items.Add("9600")
            lb.Items.Add("4800")
            lb.Items.Add("2400")
            lb.Items.Add("1200")
            AddHandler lb.SelectedIndexChanged, AddressOf ListItemSelected

            edSvc.DropDownControl(lb)
        End If
        Return lb.SelectedItem
    End Function

    Private Sub ListItemSelected(ByVal sender As Object, ByVal e As System.EventArgs)
        edSvc.CloseDropDown()
    End Sub



    ' Indicates whether the UITypeEditor supports painting a 
    ' representation of a property's value.
    Public Overloads Overrides Function GetPaintValueSupported(ByVal context As System.ComponentModel.ITypeDescriptorContext) As Boolean
        Return False
    End Function
End Class

'*************************************************
'* Create an exception class for the DF1 class
'*************************************************
<SerializableAttribute()> _
Public Class SNPXException
    Inherits Exception

    Public Sub New()
        '* Use the resource manager to satisfy code analysis CA1303
        Me.New(New System.Resources.ResourceManager("en-US", System.Reflection.Assembly.GetExecutingAssembly()).GetString("Modbus Exception"))
    End Sub

    Public Sub New(ByVal message As String)
        Me.New(message, Nothing)
    End Sub

    Public Sub New(ByVal innerException As Exception)
        Me.New(New System.Resources.ResourceManager("en-US", System.Reflection.Assembly.GetExecutingAssembly()).GetString("Modbus Exception"), innerException)
    End Sub

    Public Sub New(ByVal message As String, ByVal innerException As Exception)
        MyBase.New(message, innerException)
    End Sub

    Protected Sub New(ByVal info As System.Runtime.Serialization.SerializationInfo, ByVal context As System.Runtime.Serialization.StreamingContext)
        MyBase.New(info, context)
    End Sub
End Class


