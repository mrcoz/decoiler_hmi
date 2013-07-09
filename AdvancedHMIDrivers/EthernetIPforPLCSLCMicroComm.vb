'**********************************************************************************************
'* PCCC over Ethernet/IP
'*
'* Archie Jacobs
'* Manufacturing Automation, LLC
'* ajacobs@mfgcontrol.com
'* 01-DEC-09
'*
'* Copyright 2009, 2010 Archie Jacobs
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
'* 01-DEC-09  Adapted to use Ethernet/IP as the data layer 
'*******************************************************************************************************
Imports System.ComponentModel.Design


Public Class EthernetIPforPLCSLCMicroComm
    Inherits AllenBradleyPCCC
    Implements AdvancedHMIDrivers.IComComponent

    '* Create a common instance to share so multiple DF1Comms can be used in a project
    Private Shared DLL(10) As MfgControl.AdvancedHMI.Drivers.CIP
    Private MyDLLInstance As Integer

#Region "Constructor"
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
        'If DLL Is Nothing Then
        '    DLL = New CIP
        'End If
        'AddHandler DLL.DataReceived, Dr
    End Sub

    'Component overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        '* The handle linked to the DataLink Layer has to be removed, otherwise it causes a problem when a form is closed
        If DLL(MyDLLInstance) IsNot Nothing Then RemoveHandler DLL(MyDLLInstance).DataReceived, Dr

        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
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
            While DLL(i) IsNot Nothing AndAlso DLL(i).EIPEncap.IPAddress <> m_IPAddress AndAlso i < 11
                i += 1
            End While
            MyDLLInstance = i
        End If

        If DLL(MyDLLInstance) Is Nothing Then
            DLL(MyDLLInstance) = New MfgControl.AdvancedHMI.Drivers.CIP
            DLL(MyDLLInstance).EIPEncap.IPAddress = m_IPAddress
        End If
        AddHandler DLL(MyDLLInstance).DataReceived, Dr
    End Sub
#End Region

#Region "Properties"
    Private m_IPAddress As String = "192.168.0.10"
    <System.ComponentModel.Category("Communication Settings")> _
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

        Set(ByVal Value As System.ComponentModel.ISynchronizeInvoke)
            If Not Value Is Nothing Then
                m_SynchronizingObject = Value
            End If
        End Set
    End Property
#End Region

#Region "Helper"
    '****************************************************
    '* Wait for a response from PLC before returning
    '****************************************************
    Dim MaxTicks As Integer = 85  '* 50 ticks per second
    Friend Overrides Function WaitForResponse(ByVal rTNS As Integer) As Integer
        'Responded = False

        Dim Loops As Integer = 0
        While Not Responded(rTNS) And Loops < MaxTicks
            'Application.DoEvents()
            System.Threading.Thread.Sleep(20)
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
    Friend Overrides Function PrefixAndSend(ByVal Command As Byte, ByVal Func As Byte, ByVal data() As Byte, ByVal Wait As Boolean, ByVal TNS As Integer) As Integer
        '14-OCT-12, 16-OCT-12 Return a negative value, so it knows nothing was sent
        If m_IPAddress = "0.0.0.0" Then
            Return -10000
        End If

        Dim PacketSize As Integer
        'PacketSize = data.Length + 6
        PacketSize = data.Length + 4 '* make this more generic for CIP Ethernet/IP encap


        Dim CommandPacket(PacketSize) As Byte

        Dim TNSLowerByte As Integer = TNS And &HFF

        CommandPacket(0) = Command
        CommandPacket(1) = 0       '* STS (status, always 0)


        CommandPacket(2) = (TNSLowerByte)
        CommandPacket(3) = (TNS >> 8)

        '*Mark whether this was requested by a subscription or not
        '* FIX
        PLCAddressByTNS(TNSLowerByte).InternallyRequested = InternalRequest


        CommandPacket(4) = Func

        If data.Length > 0 Then
            data.CopyTo(CommandPacket, 5)
        End If

        Responded(TNSLowerByte) = False
        Dim result As Integer
        result = SendData(CommandPacket, TNS)


        If result = 0 And Wait Then
            result = WaitForResponse(TNSLowerByte)

            '* Return status byte that came from controller
            If result = 0 Then
                If DataPackets(TNSLowerByte) IsNot Nothing Then
                    If (DataPackets(TNSLowerByte).Count > 3) Then
                        result = DataPackets(TNSLowerByte)(3)  '* STS position in DF1 message
                        '* If its and EXT STS, page 8-4
                        If result = &HF0 Then
                            '* The EXT STS is the last byte in the packet
                            'result = DataPackets(rTNS)(DataPackets(rTNS).Count - 2) + &H100
                            result = DataPackets(TNSLowerByte)(DataPackets(TNSLowerByte).Count - 1) + &H100
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
        'PacketSize = 5
        PacketSize = 3    'Ethernet/IP Preparation


        Dim CommandPacket(PacketSize) As Byte
        Dim BytePos As Integer

        'CommandPacket(1) = m_TargetNode
        'CommandPacket(0) = m_MyNode
        'BytePos = 2
        BytePos = 0

        CommandPacket(BytePos) = Command
        CommandPacket(BytePos + 1) = 0       '* STS (status, always 0)

        CommandPacket(BytePos + 2) = (rTNS And 255)
        CommandPacket(BytePos + 3) = (rTNS >> 8)


        Dim result As Integer
        result = SendData(CommandPacket, rTNS)
    End Function

    '* This is needed so the handler can be removed
    Private Dr As EventHandler = AddressOf DataLinkLayer_DataReceived
    'Private Function SendData(ByVal data() As Byte, ByVal MyNode As Byte, ByVal TargetNode As Byte) As Integer
    Private Function SendData(ByVal data() As Byte, ByVal TNS As Integer) As Integer
        If DLL Is Nothing OrElse DLL(MyDLLInstance) Is Nothing Then
            CreateDLLInstance()
        End If

        Return DLL(MyDLLInstance).ExecutePCCC(data, TNS)
    End Function

#End Region

#Region "Public Methods"
    Public Sub CloseConnection()
        DLL(MyDLLInstance).ForwardClose()
    End Sub
#End Region

End Class