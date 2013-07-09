'**********************************************************************************************
'* DF1 Data Link Layer & Application Layer
'*
'* Archie Jacobs
'* Manufacturing Automation, LLC
'* ajacobs@mfgcontrol.com
'* 22-NOV-06
'*
'* Copyright 2006, 2010 Archie Jacobs
'*
'* This class implements the two layers of the Allen Bradley DF1 protocol.
'* In terms of the AB documentation, the data link layer acts as the transmitter and receiver.
'* Communication commands in the format described in chapter 7, are passed to
'* the data link layer using the SendData method.
'*
'* Reference : Allen Bradley Publication 1770-6.5.16
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
'* 09-JUL-11  Split up in order to make a single class common for both DF1 and EthernetIP
'*******************************************************************************************************
Imports System.ComponentModel.Design
Imports System.ComponentModel

'<Assembly: system.Security.Permissions.SecurityPermissionAttribute(system.Security.Permissions.SecurityAction.RequestMinimum)> 
Public Class DF1Comm
    Inherits AllenBradleyPCCC
    Implements AdvancedHMIDrivers.IComComponent

    '* Create a common instance to share so multiple DF1Comms can be used in a project
    Private Shared DLL(10) As MfgControl.AdvancedHMI.Drivers.DF1DataLinkLayer
    Private MyDLLInstance As Integer

    Public Event AutoDetectTry As EventHandler


#Region "Constructor"
    Public Sub New(ByVal container As System.ComponentModel.IContainer)
        MyBase.New()

        'Required for Windows.Forms Class Composition Designer support
        container.Add(Me)
    End Sub

    '***************************************************************
    '* Create the Data Link Layer Instances
    '* if the IP Address is the same, then resuse a common instance
    '***************************************************************
    Friend Overrides Sub CreateDLLInstance()
        If DLL(0) IsNot Nothing Then
            '* At least one DLL instance already exists,
            '* so check to see if it has the same IP address
            '* if so, reuse the instance, otherwise create a new one
            Dim i As Integer
            While DLL(i) IsNot Nothing AndAlso DLL(i).ComPort <> m_ComPort AndAlso i < 11
                i += 1
            End While
            MyDLLInstance = i
        End If

        If DLL(MyDLLInstance) Is Nothing Then
            DLL(MyDLLInstance) = New MfgControl.AdvancedHMI.Drivers.DF1DataLinkLayer

            DLL(MyDLLInstance).ComPort = m_ComPort
            DLL(MyDLLInstance).Parity = m_Parity
            DLL(MyDLLInstance).ChecksumType = m_CheckSumType
        End If
        AddHandler DLL(MyDLLInstance).DataReceived, Dr
    End Sub

    Friend Dr As EventHandler = AddressOf DataLinkLayer_DataReceived

    'Component overrides dispose to clean up the component list.
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        '* The handle linked to the DataLink Layer has to be removed, otherwise it causes a problem when a form is closed
        If DLL(MyDLLInstance) IsNot Nothing Then RemoveHandler DLL(MyDLLInstance).DataReceived, Dr

        MyBase.Dispose(disposing)
    End Sub
#End Region

#Region "Properties"
    Private m_BaudRate As String = "AUTO"
    <EditorAttribute(GetType(BaudRateEditor), GetType(System.Drawing.Design.UITypeEditor))> _
Public Property BaudRate() As String
        Get
            Return m_BaudRate
        End Get
        Set(ByVal value As String)
            If value <> m_BaudRate Then
                If Not Me.DesignMode Then
                    '* If a new instance needs to be created, such as a different Com Port
                    CreateDLLInstance()

                    If DLL(MyDLLInstance) IsNot Nothing Then
                        DLL(MyDLLInstance).CloseComms()
                        Try
                            DLL(MyDLLInstance).BaudRate = value
                        Catch ex As Exception
                            '* 0 means AUTO to the data link layer
                            DLL(MyDLLInstance).BaudRate = 0
                        End Try
                    End If
                End If
                m_BaudRate = value
            End If
        End Set
    End Property

    '* This is need so the current value of Auto detect can be viewed
    Public ReadOnly Property ActualBaudRate() As String
        Get
            If DLL(MyDLLInstance) Is Nothing Then
                Return "0"
            Else
                Return DLL(MyDLLInstance).BaudRate
            End If
        End Get
    End Property

    Private m_ComPort As String = "COM1"
    Public Property ComPort() As String
        Get
            'Return DLL(MyDLLInstance).ComPort
            Return m_ComPort
        End Get
        Set(ByVal value As String)
            'If value <> DLL(MyDLLInstance).ComPort Then DLL(MyDLLInstance).CloseComms()
            'DLL(MyDLLInstance).ComPort = value
            m_ComPort = value

            '* If a new instance needs to be created, such as a different Com Port
            CreateDLLInstance()


            If DLL(MyDLLInstance) Is Nothing Then
            Else
                DLL(MyDLLInstance).ComPort = value
            End If
        End Set
    End Property

    Private m_Parity As System.IO.Ports.Parity = IO.Ports.Parity.None
    Public Property Parity() As System.IO.Ports.Parity
        Get
            Return m_Parity
        End Get
        Set(ByVal value As System.IO.Ports.Parity)
            m_Parity = value
        End Set
    End Property


    Public Enum CheckSumOptions
        Crc = 0
        Bcc = 1
    End Enum

    Private m_CheckSumType As CheckSumOptions
    Public Property CheckSumType() As CheckSumOptions
        Get
            Return m_CheckSumType
        End Get
        Set(ByVal value As CheckSumOptions)
            m_CheckSumType = value
            If DLL(MyDLLInstance) IsNot Nothing Then
                DLL(MyDLLInstance).ChecksumType = m_CheckSumType
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
    Public Overrides Property SynchronizingObject() As System.ComponentModel.ISynchronizeInvoke
        Get
            'If Me.Site.DesignMode Then

            Dim host1 As IDesignerHost
            Dim obj1 As Object
            If (m_SynchronizingObject Is Nothing) AndAlso Me.DesignMode Then
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

#Region "Public Methods"

    '***************************************************************
    '* This method is intended to make it easy to configure the
    '* comm port settings. It is similar to the auto configure
    '* in RSLinx.
    '* It uses the echo command and sends the character "A", then
    '* checks if it received a response.
    '**************************************************************
    ''' <summary>
    ''' This method is intended to make it easy to configure the
    ''' comm port settings. It is similar to the auto configure
    ''' in RSLinx. A successful configuration returns a 0 and sets the
    ''' properties to the discovered values.
    ''' It will fire the event "AutoDetectTry" for each setting attempt
    ''' It uses the echo command and sends the character "A", then
    ''' checks if it received a response.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function DetectCommSettings() As Integer
        'Dim rTNS As Integer

        Dim data() As Byte = {65}
        Dim BaudRates() As Integer = {38400, 19200, 9600, 1200}
        Dim BRIndex As Integer = 0
        Dim Parities() As System.IO.Ports.Parity = {System.IO.Ports.Parity.None, System.IO.Ports.Parity.Even}
        Dim PIndex As Integer
        Dim Checksums() As CheckSumOptions = {CheckSumOptions.Crc, CheckSumOptions.Bcc}
        Dim CSIndex As Integer
        Dim reply As Integer = -1

        DisableEvent = True
        '* We are sending a small amount of data, so speed up the response
        MaxTicks = 3
        While BRIndex < BaudRates.Length And reply <> 0
            PIndex = 0
            While PIndex < Parities.Length And reply <> 0
                CSIndex = 0
                While CSIndex < Checksums.Length And reply <> 0
                    DLL(MyDLLInstance).CloseComms()
                    'm_BaudRate = BaudRates(BRIndex)
                    DLL(MyDLLInstance).BaudRate = BaudRates(BRIndex)
                    DLL(MyDLLInstance).Parity = Parities(PIndex)
                    DLL(MyDLLInstance).ChecksumType = Checksums(CSIndex)

                    RaiseEvent AutoDetectTry(Me, System.EventArgs.Empty)


                    '* Send an ENQ sequence until we get a reply
                    reply = DLL(MyDLLInstance).SendENQ()

                    '* If we pass the ENQ test, then test an echo
                    '* send an "A" and look for echo back
                    If reply = 0 Then
                        Dim rTNS As Integer
                        reply = PrefixAndSend(&H6, &H0, data, True, rTNS)
                    End If

                    '* If port cannot be opened, do not retry
                    If reply = -6 Then Return reply

                    MaxTicks += 1
                    CSIndex += 1
                End While
                PIndex += 1
            End While
            BRIndex += 1
        End While

        DisableEvent = False
        MaxTicks = 85
        Return reply
    End Function


    'End of Public Methods
#End Region

#Region "Helper"
    '**************************************************************
    '* This method implements the common application routine
    '* as discussed in the Software Layer section of the AB manual
    '**************************************************************
    Friend Overrides Function PrefixAndSend(ByVal Command As Byte, ByVal Func As Byte, ByVal data() As Byte, ByVal Wait As Boolean, ByVal TNS As Integer) As Integer
        Dim PacketSize As Integer
        'PacketSize = data.Length + 6
        PacketSize = data.Length + 4 '* make this more generic for CIP Ethernet/IP encap


        Dim CommandPacke(PacketSize) As Byte
        Dim BytePos As Integer

        'CommandPacke(0) = TargetNode
        'CommandPacke(1) = MyNode
        'BytePos = 2
        BytePos = 0

        CommandPacke(BytePos) = Command
        CommandPacke(BytePos + 1) = 0       '* STS (status, always 0)

        CommandPacke(BytePos + 2) = (TNS And 255)
        CommandPacke(BytePos + 3) = (TNS >> 8)

        '*Mark whether this was requested by a subscription or not
        '* FIX
        PLCAddressByTNS(TNS And 255).InternallyRequested = InternalRequest


        CommandPacke(BytePos + 4) = Func

        If data.Length > 0 Then
            data.CopyTo(CommandPacke, BytePos + 5)
        End If

        Dim rTNS As Integer = TNS And &HFF
        Responded(rTNS) = False
        Dim result As Integer
        result = DLL(MyDLLInstance).SendData(CommandPacke, MyNode, TargetNode)


        If result = 0 And Wait Then
            result = WaitForResponse(rTNS)

            '* Return status byte that came from controller
            If result = 0 Then
                If DataPackets(rTNS) IsNot Nothing Then
                    If (DataPackets(rTNS).Count > 3) Then
                        result = DataPackets(rTNS)(3)  '* STS position in DF1 message
                        '* If its and EXT STS, page 8-4
                        If result = &HF0 Then
                            '* The EXT STS is the last byte in the packet
                            'result = DataPackets(rTNS)(DataPackets(rTNS).Count - 2) + &H100
                            result = DataPackets(rTNS)(DataPackets(rTNS).Count - 1) + &H100
                        End If
                    End If
                Else
                    result = -8 '* no response came back from PLC
                End If
            Else
                Dim DebugCheck As Integer = 0
            End If
        Else
            Dim DebugCheck As Integer = 0
        End If

        Return result
    End Function

    '**************************************************************
    '* This method Sends a response from an unsolicited msg
    '**************************************************************
    Private Function SendResponse(ByVal Command As Byte, ByVal rTNS As Integer) As Integer
        Dim PacketSize As Integer
        'PacketSize = Data.Length + 5
        PacketSize = 5
        PacketSize = 3    'Ethernet/IP Preparation


        Dim CommandPacke(PacketSize) As Byte
        Dim BytePos As Integer

        CommandPacke(1) = TargetNode
        CommandPacke(0) = MyNode
        BytePos = 2
        BytePos = 0

        CommandPacke(BytePos) = Command
        CommandPacke(BytePos + 1) = 0       '* STS (status, always 0)

        CommandPacke(BytePos + 2) = (rTNS And 255)
        CommandPacke(BytePos + 3) = (rTNS >> 8)


        Dim result As Integer
        result = DLL(MyDLLInstance).SendData(CommandPacke, MyNode, TargetNode)
    End Function


    '****************************************************
    '* Wait for a response from PLC before returning
    '****************************************************
    Private MaxTicks As Integer = 85  '* 50 ticks per second
    Friend Overrides Function WaitForResponse(ByVal rTNS As Integer) As Integer
        Dim Loops As Integer = 0
        While Not Responded(rTNS) And Loops < MaxTicks
            System.Threading.Thread.Sleep(20)
            Loops += 1
        End While

        If Loops >= MaxTicks Then
            Return -20
        ElseIf DLL(MyDLLInstance).LastResponseWasNAK Then
            Return -21
        Else
            Return 0
        End If
    End Function
#End Region
End Class

