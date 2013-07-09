'**********************************************************************************************
'* Siemens PPI S7-200 Protocol
'*
'* Archie Jacobs
'* Manufacturing Automation, LLC
'* ajacobs@mfgcontrol.com
'* 22-OCT-10
'*
'* This class implements the two layers of the Siemens PPI protocol for s7-200.
'*
'* Reference : Siemens
'*
'* Copyright 2010 Archie Jacobs
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
'*
'**********************************************************************************************
Imports System.ComponentModel.Design
Imports System.Text.RegularExpressions
Imports System.ComponentModel

Public Class SiemensPPIComm
    Inherits System.ComponentModel.Component
    Implements AdvancedHMIDrivers.IComComponent

    '<Assembly: system.Security.Permissions.SecurityPermissionAttribute(system.Security.Permissions.SecurityAction.RequestMinimum)> 
    Private CurrentAddressToRead As String
    Private rnd As New Random
    '* create a random number as a starting point
    Private PDUReference As Int16 = (rnd.Next And &H7F) + 1
    Private DLL As MfgControl.AdvancedHMI.Drivers.PPIDataLinkLayer
    Public DataPackets(255) As System.Collections.ObjectModel.Collection(Of Byte)
    Private Responded(255) As Boolean
    '* keep the original address by ref of low PDU Ref byte so it can be returned to a linked polling address
    Private PLCAddressByPDURef(255) As ParsedDataAddress
    Private PolledAddressList As New List(Of PolledAddressInfo)
    Private tmrPollList As New List(Of System.Timers.Timer)
    Private DisableEvent As Boolean

    Public Delegate Sub PLCCommEventHandler(ByVal sender As Object, ByVal e As MfgControl.AdvancedHMI.Drivers.Common.PlcComEventArgs)

    Public Event DataReceived As PLCCommEventHandler

#Region "ComponentHandlers"

    Private components As System.ComponentModel.IContainer
    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New(ByVal container As System.ComponentModel.IContainer)
        MyClass.New()

        'Required for Windows.Forms Class Composition Designer support
        container.Add(Me)
    End Sub

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New()
        MyBase.New()


        If DLL Is Nothing Then
            DLL = New MfgControl.AdvancedHMI.Drivers.PPIDataLinkLayer
            DLL.BaudRate = m_BaudRate
        End If
        AddHandler DLL.DataReceived, Dr
    End Sub

    'Component overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        '* The handle linked to the DataLink Layer has to be removed, otherwise it causes a problem when a form is closed
        If DLL IsNot Nothing Then RemoveHandler DLL.DataReceived, Dr

        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub
#End Region


#Region "Properties"
    Private m_MyNode As Integer
    Public Property MyNode() As Integer
        Get
            Return m_MyNode
        End Get
        Set(ByVal value As Integer)
            m_MyNode = value
        End Set
    End Property

    Private m_TargetNode As Integer = 2
    Public Property TargetNode() As Integer
        Get
            Return m_TargetNode
        End Get
        Set(ByVal value As Integer)
            m_TargetNode = value
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
            m_AsyncMode = value
        End Set
    End Property

    Private m_BaudRate As String = "9600"
    <EditorAttribute(GetType(BaudRateEditor), GetType(System.Drawing.Design.UITypeEditor))> _
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
            'DLL.BaudRate = value
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
    Private m_SynchronizingObject As System.ComponentModel.ISynchronizeInvoke
    '* do not let this property show up in the property window
    ' <System.ComponentModel.Browsable(False)> _
    Public Property SynchronizingObject() As System.ComponentModel.ISynchronizeInvoke
        Get
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
        End Get

        Set(ByVal Value As System.ComponentModel.ISynchronizeInvoke)
            If Not Value Is Nothing Then
                m_SynchronizingObject = Value
            End If
        End Set
    End Property

#End Region
#Region "Subscription Routines"
    '*********************************************************************************
    '* This is used for linking a notification.
    '* An object can request a continuous poll and get a callback when value updated
    '*********************************************************************************
    Delegate Sub ReturnValues(ByVal Values As String)
    Private Structure PolledAddressInfo
        Dim PLCAddress As String
        Dim DataArea As Integer
        Dim DataType As Integer
        Dim Offset As Integer
        Dim SubElement As Integer
        Dim BitNumber As Integer
        Dim BytesPerElement As Integer
        Dim dlgCallBack As IComComponent.ReturnValues
        Dim PollRate As Integer
        Dim ID As Integer
        Dim LastValue As Object
        Dim ElementsToRead As Integer
        Dim TargetNode As Integer '27-APR-10
    End Structure
    Private CurrentID As Integer = 1


    Public Function Subscribe(ByVal PLCAddress As String, ByVal numberOfElements As Int16, ByVal PollRate As Integer, ByVal CallBack As IComComponent.ReturnValues) As Integer Implements IComComponent.Subscribe
        '*******************************************************************
        '*******************************************************************
        Dim ParsedResult As ParsedDataAddress = ParseAddress(PLCAddress)

        '* Valid address?
        If ParsedResult.DataType <> 0 Then
            Dim tmpPA As New PolledAddressInfo
            tmpPA.PLCAddress = PLCAddress
            'tmpPA.PollRate = CycleTime
            tmpPA.PollRate = PollRate
            tmpPA.dlgCallBack = CallBack
            tmpPA.ID = CurrentID
            tmpPA.DataArea = ParsedResult.DataArea
            tmpPA.Offset = ParsedResult.Offset
            tmpPA.SubElement = ParsedResult.SubElement
            tmpPA.BitNumber = ParsedResult.BitNumber
            tmpPA.DataType = ParsedResult.DataType
            tmpPA.BytesPerElement = ParsedResult.BytesPerElements
            tmpPA.ElementsToRead = 1
            tmpPA.TargetNode = m_TargetNode


            PolledAddressList.Add(tmpPA)

            PolledAddressList.Sort(AddressOf SortPolledAddresses)

            '* The ID is used as a reference for removing polled addresses
            CurrentID += 1


            '********************************************************************
            '* Check to see if there already exists a timer for this poll rate
            '********************************************************************
            If tmrPollList.Count < 1 Then
                Dim j As Integer = 0
                'While j < tmrPollList.Count AndAlso tmrPollList(j) IsNot Nothing AndAlso tmrPollList(j).Interval <> CycleTime
                '    j += 1
                'End While

                'If j >= tmrPollList.Count Then
                '* Add new timer
                Dim tmrTemp As New System.Timers.Timer
                'If CycleTime > 0 Then
                '    tmrTemp.Interval = CycleTime
                'Else
                '    tmrTemp.Interval = 10
                'End If
                If PollRate <= 0 Then PollRate = 250
                tmrTemp.Interval = PollRate

                tmrPollList.Add(tmrTemp)
                AddHandler tmrPollList(j).Elapsed, AddressOf PollUpdate

                tmrTemp.Enabled = True
                'End If
            End If

            Return tmpPA.ID
        End If
        Return -1
    End Function

    '***************************************************************
    '* Used to sort polled addresses by File Number and element
    '* This helps in optimizing reading
    '**************************************************************
    Private Function SortPolledAddresses(ByVal A1 As PolledAddressInfo, ByVal A2 As PolledAddressInfo) As Integer
        If A1.DataArea = A2.DataArea Then
            If A1.Offset > A2.Offset Then
                Return 1
            ElseIf A1.Offset = A2.Offset Then
                Return 0
            Else
                Return -1
            End If
        End If

        If A1.DataArea > A2.DataArea Then
            Return 1
        Else
            Return -1
        End If
    End Function

    Public Function UnSubscribe(ByVal ID As Integer) As Integer Implements IComComponent.UnSubscribe
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
    Private Sub PollUpdate(ByVal sender As System.Object, ByVal e As System.Timers.ElapsedEventArgs)
        If m_DisableSubscriptions Then Exit Sub

        Dim intTimerIndex As Integer = tmrPollList.IndexOf(sender)

        '* Stop the poll timer
        tmrPollList(intTimerIndex).Enabled = False

        Dim i, NumberToRead, FirstElement As Integer
        Dim HighestBit As Integer = PolledAddressList(i).BitNumber
        While i < PolledAddressList.Count
            'Dim NumberToReadCalc As Integer
            NumberToRead = PolledAddressList(i).ElementsToRead
            FirstElement = i
            Dim PLCAddress As String = PolledAddressList(FirstElement).PLCAddress

            '* Group into the same read if there is less than a 20 element gap
            '* Do not group IO addresses because they can exceed 16 bits which causes problems
            'Dim ElementSpan As Integer = 20
            'While i < PolledAddressList.Count - 1 AndAlso (PolledAddressList(i).FileNumber = PolledAddressList(i + 1).FileNumber And _
            '                ((PolledAddressList(i + 1).ElementNumber - PolledAddressList(i).ElementNumber < 20 And PolledAddressList(i).FileType <> &H8B And PolledAddressList(i).FileType <> &H8C) Or _
            '                    PolledAddressList(i + 1).ElementNumber = PolledAddressList(i).ElementNumber))
            '    NumberToReadCalc = PolledAddressList(i + 1).ElementNumber - PolledAddressList(FirstElement).ElementNumber + PolledAddressList(i + 1).ElementsToRead
            '    If NumberToReadCalc > NumberToRead Then NumberToRead = NumberToReadCalc

            '    '* This is used for IO addresses wher the bit can be above 15
            '    If PolledAddressList(i).BitNumber < 99 And PolledAddressList(i).BitNumber > HighestBit Then HighestBit = PolledAddressList(i).BitNumber

            '    i += 1
            'End While


            '* Get file type designation.
            '* Is it more than one character (e.g. "ST")
            'If PolledAddressList(FirstElement).PLCAddress.Substring(1, 1) >= "A" And PolledAddressList(FirstElement).PLCAddress.Substring(1, 1) <= "Z" Then
            '    PLCAddress = PolledAddressList(FirstElement).PLCAddress.Substring(0, 2) & PolledAddressList(FirstElement).DataArea & ":" & PolledAddressList(FirstElement).Offset
            'Else
            '    PLCAddress = PolledAddressList(FirstElement).PLCAddress.Substring(0, 1) & PolledAddressList(FirstElement).DataArea & ":" & PolledAddressList(FirstElement).Offset
            'End If

            PLCAddress = PolledAddressList(FirstElement).PLCAddress


            If PolledAddressList(i).PollRate = sender.Interval Or SavedPollRate > 0 Then
                '* Make sure it does not wait for return value befoe coming back
                Dim tmp As Boolean = Me.AsyncMode
                Me.AsyncMode = True
                Try
                    InternalRequest = True
                    Me.ReadAny(PLCAddress, NumberToRead)
                    'Me.ReadAny(PolledAddressList(FirstElement).PLCAddress, 1)
                    If SavedPollRate <> 0 Then
                        tmrPollList(0).Interval = SavedPollRate
                        SavedPollRate = 0
                    End If
                Catch ex As Exception
                    '* Send this message back to the requesting control
                    Dim TempArray() As String = {ex.Message}

                    'm_SynchronizingObject.BeginInvoke(PolledAddressList(i).dlgCallBack, CObj(TempArray))
                    m_SynchronizingObject.BeginInvoke(PolledAddressList(i).dlgCallBack, New Object() {TempArray})
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
#End Region


    '*************************************************************
    '* Overloaded method of ReadAny - that reads only one element
    '*************************************************************
    Public Function ReadAny(ByVal startAddress As String) As String Implements IComComponent.ReadAny
        Return ReadAny(startAddress, 1)(0)
    End Function

    Public Function ReadSynchronous(ByVal startAddress As String, ByVal numberOfElements As Integer) As String() Implements IComComponent.ReadSynchronous
        Return ReadAny(startAddress, numberOfElements, False)
    End Function

    Public Function ReadAny(ByVal startAddress As String, ByVal numberOfElements As Integer) As String() Implements IComComponent.ReadAny
        Return ReadAny(startAddress, numberOfElements, m_AsyncMode)
    End Function


    '*********************************************************************************************
    Public Function ReadAny(ByVal startAddress As String, ByVal numberOfElements As Integer, ByVal asyncModeIn As Boolean) As String()
        Dim ParsedResult As ParsedDataAddress = ParseAddress(startAddress)

        '* Save the address information, needed for extracting returned data
        PLCAddressByPDURef(PDUReference) = ParsedResult
        PLCAddressByPDURef(PDUReference).TargetNode = m_TargetNode
        PLCAddressByPDURef(PDUReference).InternallyRequested = InternalRequest

        '* Store this here for use by extract data
        ParsedResult.NumberOfElements = numberOfElements

        '* Invalid address?
        If ParsedResult.DataArea = 0 Then
            Throw New MfgControl.AdvancedHMI.Drivers.Common.PLCDriverException("Invalid Address")
        End If

        '* Save this for being used by the polling link
        CurrentAddressToRead = startAddress

        '* If requesting 0 elements, then default to 1
        Dim ArrayElements As Int16 = numberOfElements - 1
        If ArrayElements < 0 Then
            ArrayElements = 0
        End If


        '*********************************
        '* Start building the byte stream
        '*********************************
        '********************************************************
        Dim VariableStream() As Byte = BuildANYPointer(ParsedResult)
        Dim ParameterBlock(VariableStream.Length + 2 - 1) As Byte
        '* Service ID, 4=Read
        ParameterBlock(0) = &H4
        '* Number of variables in Parameter block
        ParameterBlock(1) = 1
        VariableStream.CopyTo(ParameterBlock, 2)

        '********************************************************
        Dim PDUHeader(9) As Byte
        'ProtoID, 32=S7-200
        PDUHeader(0) = &H32
        'Remote operating Services Control (ROSCTR)
        PDUHeader(1) = 1
        '* RED_ID - Unused so Always 0
        PDUHeader(2) = 0
        PDUHeader(3) = 0
        '* PDU Reference
        PDUHeader(4) = PDUReference >> 8
        PDUHeader(5) = PDUReference And 255
        '* Save this for a sync read
        Dim SavedPDURef As Int16 = PDUHeader(5)

        If PDUReference < 255 Then
            PDUReference += 1
        Else
            PDUReference = 0
        End If
        '* Parameter Length
        PDUHeader(6) = ParameterBlock.Length >> 8
        PDUHeader(7) = ParameterBlock.Length And 255
        '* Data Length  =0 for a read because it is not sending actual data values
        PDUHeader(8) = 0
        PDUHeader(9) = 0

        Dim PDU(PDUHeader.Length + ParameterBlock.Length - 1) As Byte
        PDUHeader.CopyTo(PDU, 0)
        ParameterBlock.CopyTo(PDU, PDUHeader.Length)



        Dim result As Integer
        result = DLL.RequestData(PDU, m_MyNode, m_TargetNode)

        '*************************
        '* Request the data back
        '*************************
        If result = 0 Then
            Responded(PDUReference And 255) = False
            DLL.RequestResponse(m_MyNode, m_TargetNode)

            If Not asyncModeIn Then
                Dim Ticks As Integer
                While Ticks < 20 And Not Responded(SavedPDURef And 255)
                    Threading.Thread.Sleep(20)
                    Ticks += 1
                End While

                Dim ReturnedData(DataPackets(SavedPDURef).Count - 22) As Byte
                For i As Integer = 0 To DLL.DataPacket.Count - 22
                    ReturnedData(i) = DataPackets(SavedPDURef)(i + 21)
                Next

                Return ExtractData(ParsedResult, ReturnedData)
                'Dim y = Responded(PDUReference And 255)
            End If
        End If

        Dim reply As Integer
        'Dim ReturnedData(NumberOfBytes - 1) As Byte
        'Dim ReturnedDataIndex As Integer


        'Dim Retries As Integer
        'While reply = 0 AndAlso ReturnedDataIndex < NumberOfBytes
        '    BytesToRead = NumberOfBytes

        '    Dim ReturnedData2(BytesToRead) As Byte
        '    ReturnedData2 = ReadRawData(ParsedResult, BytesToRead, reply)

        '    '* Point to next set of bytes to read in block
        '    If reply = 0 Then
        '        '* Point to next element to begin reading
        '        ReturnedData2.CopyTo(ReturnedData, ReturnedDataIndex)
        '        ReturnedDataIndex += BytesToRead
        '    ElseIf Retries < 2 Then
        '        Retries += 1
        '        reply = 0
        '    Else
        '        '* An error was returned from the read operation
        '        Throw New DF1Exception(DecodeMessage(reply))
        '    End If
        'End While


        If m_AsyncMode Then
            Dim x() As String = {reply}
            Return x
        Else
            'Dim AckWaitTicks As Integer
            'While (Not Acknowledged And Not NotAcknowledged) And AckWaitTicks < 50
            '    System.Threading.Thread.Sleep(20)
            '    AckWaitTicks += 1
            'End While

            'If AckWaitTicks >= 50 Then
            '    Dim x() As String = {-3}
            '    Return x
            'End If
            'Return -3

            Dim x = 0
            'Return ExtractData(ParsedResult, ReturnedData)
            'TODO:
            Dim d() As String = {0}
            Return d
        End If

    End Function

    Public Function WriteData(ByVal startAddress As String, ByVal dataToWrite As String) As String Implements IComComponent.WriteData
        If dataToWrite Is Nothing Then
            Return 0
        End If

        Dim ParsedResult As ParsedDataAddress = ParseAddress(startAddress)

        '* Save the address information, needed for extracting returned data
        PLCAddressByPDURef(PDUReference) = ParsedResult
        PLCAddressByPDURef(PDUReference).TargetNode = m_TargetNode
        PLCAddressByPDURef(PDUReference).InternallyRequested = InternalRequest

        '* Store this here for use by extract data
        ParsedResult.NumberOfElements = 1

        '* Invalid address?
        If ParsedResult.DataArea = 0 Then
            Throw New MfgControl.AdvancedHMI.Drivers.Common.PLCDriverException("Invalid Address")
        End If



        '*********************************
        '* Start building the byte stream
        '*********************************
        '********************************************************
        Dim VariableStream() As Byte = BuildANYPointer(ParsedResult)
        Dim ParameterBlock(VariableStream.Length + 2 - 1) As Byte
        '* Service ID, 4=Read, 5=Write
        ParameterBlock(0) = &H5
        '* Number of variables in Parameter block
        ParameterBlock(1) = 1
        VariableStream.CopyTo(ParameterBlock, 2)

        '********************************************************
        Dim PDUHeader(9) As Byte
        'ProtoID, 32=S7-200
        PDUHeader(0) = &H32
        'Remote operating Services Control (ROSCTR)
        PDUHeader(1) = 1
        '* RED_ID - Unused so Always 0
        PDUHeader(2) = 0
        PDUHeader(3) = 0
        '* PDU Reference
        PDUHeader(4) = PDUReference >> 8
        PDUHeader(5) = PDUReference And 255
        If PDUReference < 255 Then
            PDUReference += 1
        Else
            PDUReference = 0
        End If

        '* Parameter Length
        PDUHeader(6) = ParameterBlock.Length >> 8
        PDUHeader(7) = ParameterBlock.Length And 255
        '* Data Length
        PDUHeader(8) = 0
        PDUHeader(9) = 4 + ParsedResult.BytesPerElements
        '* Even number of bytes
        If (ParsedResult.BytesPerElements And 1) = 1 Then PDUHeader(9) += 1

        '* Data bytes must be even number
        Dim DataByteCount As Integer = ParsedResult.BytesPerElements
        If (DataByteCount And 1) = 1 Then DataByteCount += 1

        Dim PDU(PDUHeader.Length + ParameterBlock.Length + DataByteCount + 4 - 1) As Byte
        PDUHeader.CopyTo(PDU, 0)
        ParameterBlock.CopyTo(PDU, PDUHeader.Length)

        'TODO:
        '* Data Type - 03=Bit, 04=Byte, word, dword
        If ParsedResult.DataType = 1 Then
            PDU(PDU.Length - 5) = 3 '* bit
        Else
            PDU(PDU.Length - 5) = 4
        End If


        ' Size in bits
        If ParsedResult.DataType = 1 Then
            PDU(PDU.Length - 3) = 1 '* Bit Type
        Else
            PDU(PDU.Length - 3) = ParsedResult.BytesPerElements * 8
        End If

        '* TODO : handle string
        '* Add Value - if one byte, then use first byte
        PDU(PDU.Length - 2) = dataToWrite And 255
        If ParsedResult.BytesPerElements > 1 Then
            PDU(PDU.Length - 1) = dataToWrite And 255
            PDU(PDU.Length - 2) = (dataToWrite >> 8) And 255
        End If
        If ParsedResult.BytesPerElements > 2 Then
            PDU(PDU.Length - 3) = (dataToWrite >> 16) And 255
        End If
        If ParsedResult.BytesPerElements > 3 Then
            PDU(PDU.Length - 4) = (dataToWrite >> 24) And 255
        End If


        Dim result As Integer
        result = DLL.RequestData(PDU, m_MyNode, m_TargetNode)

        If result = 0 Then
            DLL.RequestResponse(m_MyNode, m_TargetNode)
        End If

        Return 0
    End Function


    Private Structure ParsedDataAddress
        Dim PLCAddress As String
        Dim DataArea As Integer
        Dim DataType As Integer
        Dim Offset As Integer
        Dim SubElement As Integer
        Dim BitNumber As Integer
        Dim BytesPerElements As Integer
        Dim TableSizeInBytes As Integer
        Dim NumberRead As Integer
        Dim InternallyRequested As Boolean
        Dim NumberOfElements As Integer
        Dim TargetNode As Integer
    End Structure

    '*********************************************************************************
    '* Parse the address string and validate, if invalid, Return 0 in FileType
    '* Convert the file type letter Type to the corresponding value
    '* Reference page 7-18
    '*********************************************************************************
    Private RE1 As New Regex("(?i)^\s*(?<DataArea>([IOMTCS]))(?<DataType>([BWD]))?(?<Offset>\d{1,3}).?(?<BitNumber>\d{1,3})?\s*$")
    'Private RE2 As New Regex("(?i)^\s*(?<FileType>[SBN])(?<FileNumber>\d{1,3})(/(?<BitNumber>\d{1,4}))\s*$")
    Private Function ParseAddress(ByVal DataAddress As String) As ParsedDataAddress
        Dim result As New ParsedDataAddress

        result.DataArea = 0  '* Let a 0 indicate an invalid address
        result.BitNumber = 99  '* Let a 99 indicate no bit level requested

        '*********************************
        '* Try all match patterns
        '*********************************
        Dim mc As MatchCollection = RE1.Matches(DataAddress)

        If mc.Count <= 0 Then
            'mc = RE2.Matches(DataAddress)
            'If mc.Count <= 0 Then
            Return result
            '    End If
        End If


        '*******************************************************************
        '* Keep the original address with the parsed values for later use
        '*******************************************************************
        result.PLCAddress = DataAddress


        '*********************************************
        '* Get elements extracted from match patterns
        '*********************************************
        If mc.Item(0).Groups("BitNumber").Length > 0 Then
            result.BitNumber = mc.Item(0).Groups("BitNumber").ToString
        End If

        '* Was an element number specified? 
        If mc.Item(0).Groups("Offset").Length > 0 Then
            result.Offset = mc.Item(0).Groups("Offset").ToString
        End If


        Dim DataSize As String = mc.Item(0).Groups("DataType").ToString.ToUpper(System.Globalization.CultureInfo.CurrentCulture)
        Select Case DataSize
            Case "B" : result.DataType = 2
                result.BytesPerElements = 1
            Case "W" : result.DataType = 4
                result.BytesPerElements = 2
            Case "D" : result.DataType = 6
                result.BytesPerElements = 4
            Case Else : result.DataType = 1 '* Bit
                result.BytesPerElements = 1
        End Select


        '***************************************
        '* Translate file type letter to number
        '***************************************
        Dim FileType As String = mc.Item(0).Groups("DataArea").ToString.ToUpper(System.Globalization.CultureInfo.CurrentCulture)
        Select Case FileType
            Case "I" : result.DataArea = &H81
            Case "Q" : result.DataArea = &H82
            Case "M" : result.DataArea = &H83
            Case "V" : result.DataArea = &H84
            Case "C" : result.DataArea = &H1E
                result.DataType = &H1E
            Case "T" : result.DataArea = &H1F
                result.DataType = &H1F
            Case "S" : result.DataArea = &H4
        End Select

        '**************************************************************************
        '* Was a bit number higher than 15 specified along with an element number?
        '**************************************************************************
        If result.BitNumber > 15 And result.BitNumber < 99 Then
            '* IO points can use higher bit numbers, so make sure it is not IO
            If result.DataArea <> &H8B And result.DataArea <> &H8C Then
                result.Offset += result.BitNumber >> 3
                result.BitNumber = result.BitNumber Mod 8
            End If
        End If

        Return result
    End Function

    '*************************************************************
    '* Build the data stream for a vairable in the Parameter Block
    '* Section 4.6 - Siemens PPI specification
    '*************************************************************
    Private Function BuildANYPointer(ByVal Address As ParsedDataAddress) As Byte()
        Dim Result(11) As Byte

        '* Variable Spec
        Result(0) = &H12
        '* V_ADDR_LG
        Result(1) = &HA
        '* Syntax ID
        Result(2) = &H10
        '* Element Type
        Result(3) = Address.DataType
        '* Number of Elements
        Result(4) = Address.NumberOfElements >> 8
        Result(5) = Address.NumberOfElements And 255
        '* SubArea
        Result(6) = 0
        Result(7) = 0 '=1 for V
        '* Data Area
        Result(8) = Address.DataArea
        '*Offset - shifted left by 3 bits (this sets offset to bit level by multiplying by 8)
        Result(9) = Address.Offset >> 13
        Result(10) = (Address.Offset >> 5) And 255
        Result(11) = (Address.Offset And 31) << 3
        If Address.BitNumber < 99 Then Result(11) += Address.BitNumber


        Return Result
    End Function


    ' TODO : Put this in a New event
    'Private EventHandleAdded As Boolean
    '* This is needed so the handler can be removed
    Private Dr As EventHandler = AddressOf DataLinkLayer_DataReceived
    'Private Function SendData(ByVal data() As Byte, ByVal MyNode As Byte, ByVal TargetNode As Byte) As Integer
    '    If DLL Is Nothing Then
    '        DLL = New PPIDataLinkLayer
    '        DLL.BaudRate = m_BaudRate
    '    End If


    '    'If Not EventHandleAdded Then
    '    '    AddHandler DLL.DataReceived, AddressOf DataLinkLayer_DataReceived
    '    '    EventHandleAdded = True
    '    'End If

    '    Return DLL.SendData(data, MyNode, TargetNode)
    'End Function

    '*********************************************************************************
    '*********************************************************************************
    Private Sub DataLinkLayer_DataReceived()
        '**************************************************************
        '* Only extract and send back if this response contained data
        '**************************************************************
        Dim PDURefReturned As Integer
        If DLL.DataPacket.Count > 20 Then
            PDURefReturned = DLL.DataPacket(8)
            DataPackets(PDURefReturned) = DLL.DataPacket
            Responded(PDURefReturned) = True

            '***************************************************
            '* Extract returned data into appropriate data type
            '* Transfer block of data read to the data table array
            '***************************************************
            '* TODO: Check array bounds
            Dim ReturnedData(DataPackets(PDURefReturned).Count - 22) As Byte
            For i As Integer = 0 To DLL.DataPacket.Count - 22
                ReturnedData(i) = DataPackets(PDURefReturned)(i + 21)
            Next

            '* Ignore if this was data that came back from a prior connection
            If PLCAddressByPDURef(PDURefReturned).BytesPerElements <= 0 Then
                Exit Sub
            End If

            Dim d() As String = ExtractData(PLCAddressByPDURef(PDURefReturned), ReturnedData)


            '* 27-APR-10 Make sure this corresponds to the target node requested from
            If PLCAddressByPDURef(PDURefReturned).TargetNode = DataPackets(PDURefReturned)(1) Then
                If Not PLCAddressByPDURef(PDURefReturned).InternallyRequested Then
                    If Not DisableEvent Then
                        Dim x As New MfgControl.AdvancedHMI.Drivers.Common.PlcComEventArgs(d, PLCAddressByPDURef(PDURefReturned).PLCAddress, 0)
                        If m_SynchronizingObject IsNot Nothing Then
                            Dim Parameters() As Object = {Me, x}
                            m_SynchronizingObject.BeginInvoke(drsd, Parameters)
                        Else
                            'RaiseEvent DataReceived(Me, System.EventArgs.Empty)
                            RaiseEvent DataReceived(Me, x)
                        End If
                    End If
                Else
                    '*********************************************************
                    '* Check to see if this is from the Polled variable list
                    '*********************************************************
                    Dim ParsedAddress As ParsedDataAddress = ParseAddress(PLCAddressByPDURef(PDURefReturned).PLCAddress)
                    Dim EnoughElements As Boolean
                    For i As Integer = 0 To PolledAddressList.Count - 1
                        EnoughElements = False                    '* Are there enought elements read for this request
                        If (PolledAddressList(i).Offset - PLCAddressByPDURef(PDURefReturned).Offset + PolledAddressList(i).ElementsToRead <= d.Length) And _
                            (PolledAddressList(i).DataType <> 134 And PolledAddressList(i).DataType <> 135 And PolledAddressList(i).DataType <> &H8B And PolledAddressList(i).DataType <> &H8C) Then
                            EnoughElements = True
                        End If
                        If (PolledAddressList(i).BitNumber < 16) And ((PolledAddressList(i).Offset - PLCAddressByPDURef(PDURefReturned).Offset + PolledAddressList(i).ElementsToRead) / 16 <= d.Length) Then
                            EnoughElements = True
                        End If
                        If (PolledAddressList(i).DataType = 134 Or PolledAddressList(i).DataType = 135) And (PolledAddressList(i).Offset - PLCAddressByPDURef(PDURefReturned).Offset + PolledAddressList(i).ElementsToRead) <= d.Length Then
                            EnoughElements = True
                        End If
                        '* IO addresses - be sure not to cross elements/card slots
                        If (PolledAddressList(i).DataType = &H8B Or PolledAddressList(i).DataType = &H8C And _
                                PolledAddressList(i).Offset = PLCAddressByPDURef(PDURefReturned).Offset) Then
                            Dim WordToUse As Integer = PolledAddressList(i).BitNumber >> 4
                            If (d.Length - 1) >= (PolledAddressList(i).Offset - PLCAddressByPDURef(PDURefReturned).Offset + (WordToUse)) Then
                                EnoughElements = True
                            End If
                        End If


                        If PolledAddressList(i).DataArea = PLCAddressByPDURef(PDURefReturned).DataArea And _
                            EnoughElements And _
                            PLCAddressByPDURef(PDURefReturned).Offset <= PolledAddressList(i).Offset Then ' And _
                            '((PLCAddressByTNS(TNSReturned).FileType <> &H8B And PLCAddressByTNS(TNSReturned).FileType <> &H8C) Or PLCAddressByTNS(TNSReturned).BitNumber = PolledAddressList(i).BitNumber) Then
                            'PolledAddressList(i).BitNumber = PLCAddressByTNS(TNSReturned).BitNumber Then
                            'PolledValueReturned(PLCAddressByTNS(TNSReturned).PLCAddress, d)

                            Dim BitResult(PolledAddressList(i).ElementsToRead - 1) As String
                            '* Handle timers and counters as exceptions because of the 3 subelements
                            If (PolledAddressList(i).DataType = 134 Or PolledAddressList(i).DataType = 135) Then
                                '* If this is a bit level address for a timer or counter, then handle appropriately
                                If PolledAddressList(i).BitNumber < 16 Then
                                    Try
                                        '                                    If d((PolledAddressList(i).ElementNumber - PLCAddressByTNS(TNSReturned).Element) * 3) Then
                                        BitResult(0) = (d((PolledAddressList(i).Offset - PLCAddressByPDURef(PDURefReturned).Offset) * 3) And 2 ^ PolledAddressList(i).BitNumber) > 0
                                        'End If
                                    Catch
                                        MsgBox("Error in returning data from datareceived BR0")
                                    End Try
                                Else
                                    Try
                                        For k As Integer = 0 To PolledAddressList(i).ElementsToRead - 1
                                            BitResult(k) = d((PolledAddressList(i).Offset - PLCAddressByPDURef(PDURefReturned).Offset + k) * 3 + PolledAddressList(i).SubElement)
                                        Next
                                    Catch
                                        MsgBox("Error in returning data from datareceived BRk")
                                    End Try
                                End If
                                'm_SynchronizingObject.BeginInvoke(PolledAddressList(i).dlgCallBack, CObj(BitResult))
                                m_SynchronizingObject.BeginInvoke(PolledAddressList(i).dlgCallBack, New Object() {BitResult})
                            Else
                                '* If its bit level, then return the individual bit
                                If PolledAddressList(i).BitNumber < 99 Then
                                    '*TODO : Make this handle a rquest for multiple bits
                                    Try
                                        '* Test to see if bits or integers returned
                                        Dim x As Integer
                                        Try
                                            x = d(0)
                                            If PolledAddressList(i).BitNumber < 16 Then
                                                BitResult(0) = (d(PolledAddressList(i).Offset - PLCAddressByPDURef(PDURefReturned).Offset) And 2 ^ PolledAddressList(i).BitNumber) > 0
                                            Else
                                                Dim WordToUse As Integer = PolledAddressList(i).BitNumber >> 4
                                                Dim ModifiedBitToUse As Integer = PolledAddressList(i).BitNumber Mod 16
                                                BitResult(0) = (d(PolledAddressList(i).Offset - PLCAddressByPDURef(PDURefReturned).Offset + (WordToUse)) And 2 ^ ModifiedBitToUse) > 0
                                            End If
                                        Catch ex As Exception
                                            BitResult(0) = d(0)
                                        End Try
                                    Catch ex As Exception
                                        MsgBox("Error in returning data from datareceived - " & ex.Message)
                                    End Try
                                    m_SynchronizingObject.BeginInvoke(PolledAddressList(i).dlgCallBack, New Object() {BitResult})
                                    'm_SynchronizingObject.BeginInvoke(PolledAddressList(i).dlgCallBack, CObj(BitResult))
                                Else
                                    '* All other data types
                                    For k As Integer = 0 To PolledAddressList(i).ElementsToRead - 1
                                        BitResult(k) = d((PolledAddressList(i).Offset - PLCAddressByPDURef(PDURefReturned).Offset + k))
                                    Next

                                    'm_SynchronizingObject.BeginInvoke(PolledAddressList(i).dlgCallBack, d(PolledAddressList(i).ElementNumber - PLCAddressByTNS(TNSReturned).Element))
                                    'm_SynchronizingObject.BeginInvoke(PolledAddressList(i).dlgCallBack, CObj(BitResult))
                                    m_SynchronizingObject.BeginInvoke(PolledAddressList(i).dlgCallBack, New Object() {BitResult})
                                End If
                            End If

                            'PolledAddressList(k).LastValue = values(0)
                        End If
                    Next
                End If
            End If
        End If
    End Sub

    '****************************************************************************
    '* This is required to sync the event back to the parent form's main thread
    '****************************************************************************
    Dim drsd As EventHandler = AddressOf DataReceivedSync
    'Delegate Sub DataReceivedSyncDel(ByVal sender As Object, ByVal e As EventArgs)
    Private Sub DataReceivedSync(ByVal sender As Object, ByVal e As MfgControl.AdvancedHMI.Drivers.Common.PlcComEventArgs)
        RaiseEvent DataReceived(sender, e)
    End Sub



    Private Function ExtractData(ByVal ParsedResult As ParsedDataAddress, ByVal ReturnedData() As Byte) As String()
        '***************************************************
        '* Extract returned data into appropriate data type
        '***************************************************
        If ParsedResult.BytesPerElements > 0 Then
            Dim result(Math.Floor(ReturnedData.Length / ParsedResult.BytesPerElements) - 1) As String

            '* If requesting 0 elements, then default to 1
            Dim ArrayElements As Int16 = Math.Max(result.Length - 1 - 1, 0)


            Dim i As Integer
            While i < result.Length
                Select Case ParsedResult.DataType
                    Case 1 'Bit
                        result(i) = (ReturnedData(i) <> 0)
                    Case 2 'Bytes
                        result(i) = ReturnedData(i * ParsedResult.BytesPerElements)
                    Case 4 'Word
                        result(i) = ReturnedData(i * ParsedResult.BytesPerElements) * 256 + ReturnedData(i * ParsedResult.BytesPerElements + 1)
                    Case 6 'DWord
                        result(i) = ReturnedData(i * ParsedResult.BytesPerElements) * 16777216 + ReturnedData(i * ParsedResult.BytesPerElements + 1) * 65536 + ReturnedData(i * ParsedResult.BytesPerElements + 2) * 256 + ReturnedData(i * ParsedResult.BytesPerElements + 3)
                End Select
                i += 1
            End While

            Return result
        Else
TODO:
            Dim d() As String = {0}
            Return d
        End If
    End Function
End Class

