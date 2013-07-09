'****************************************************************************
'* Archie Jacobs
'* Manufacturing Automation, LLC
'* ajacobs@mfgcontrol.com
'* 12-JUN-11
'*
'* Copyright 2011 Archie Jacobs
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
'* 04-OCT-11 Created
'****************************************************************************
Public Class ThreeButtonPLC
    Inherits ThreeButtons


#Region "PLC Related Properties"
    '*****************************************************
    '* Property - Component to communicate to PLC through
    '*****************************************************
    Private _CommComponent As AdvancedHMIDrivers.IComComponent
    <System.ComponentModel.Category("PLC Properties")> _
    Public Property CommComponent() As AdvancedHMIDrivers.IComComponent
        Get
            Return _CommComponent
        End Get
        Set(ByVal value As AdvancedHMIDrivers.IComComponent)
            _CommComponent = value
        End Set
    End Property


    '*****************************************
    '* Property - Address in PLC to Link to
    '*****************************************
    Private NotificationIDHandStatus As Integer
    Private m_PLCaddressStatusHand As String = ""
    <System.ComponentModel.Category("PLC Properties")> _
    Public Property PLCaddressStatusHand() As String
        Get
            Return m_PLCaddressStatusHand
        End Get
        Set(ByVal value As String)
            If m_PLCaddressStatusHand <> value Then
                m_PLCaddressStatusHand = value

                '* When address is changed, re-subscribe to new address
                SubscribeToCommDriver()
            End If
        End Set
    End Property


    Private NotificationIDAutoStatus As Integer
    Private m_PLCaddressStatusAuto As String = ""
    <System.ComponentModel.Category("PLC Properties")> _
    Public Property PLCaddressStatusAuto() As String
        Get
            Return m_PLCaddressStatusAuto
        End Get
        Set(ByVal value As String)
            If m_PLCaddressStatusAuto <> value Then
                m_PLCaddressStatusAuto = value

                '* When address is changed, re-subscribe to new address
                SubscribeToCommDriver()
            End If
        End Set
    End Property

    Private NotificationIDOffStatus As Integer
    Private m_PLCaddressStatusOff As String = ""
    <System.ComponentModel.Category("PLC Properties")> _
    Public Property PLCaddressStatusOff() As String
        Get
            Return m_PLCaddressStatusOff
        End Get
        Set(ByVal value As String)
            If m_PLCaddressStatusOff <> value Then
                m_PLCaddressStatusOff = value

                '* When address is changed, re-subscribe to new address
                SubscribeToCommDriver()
            End If
        End Set
    End Property


    '*****************************************
    '* Property - Address in PLC to Link to
    '*****************************************
    Private NotificationIDVisibility As Integer
    Private InvertVisible As Boolean
    Private m_PLCaddressVisibility As String = ""
    <System.ComponentModel.Category("PLC Properties")> _
    Public Property PLCaddressVisibility() As String
        Get
            Return m_PLCaddressVisibility
        End Get
        Set(ByVal value As String)
            If m_PLCaddressVisibility <> value Then
                m_PLCaddressVisibility = value

                '* When address is changed, re-subscribe to new address
                SubscribeToCommDriver()
            End If
        End Set
    End Property



    '********************************************
    '* Property - Address in PLC for click event
    '********************************************
    Private m_PLCaddressClick1 As String = ""
    <System.ComponentModel.Category("PLC Properties")> _
    Public Property PLCaddressClickAuto() As String
        Get
            Return m_PLCaddressClick1
        End Get
        Set(ByVal value As String)
            m_PLCaddressClick1 = value
        End Set
    End Property

    Private m_PLCaddressClick2 As String = ""
    <System.ComponentModel.Category("PLC Properties")> _
    Public Property PLCaddressClickHand() As String
        Get
            Return m_PLCaddressClick2
        End Get
        Set(ByVal value As String)
            m_PLCaddressClick2 = value
        End Set
    End Property

    Private m_PLCaddressClick3 As String = ""
    <System.ComponentModel.Category("PLC Properties")> _
    Public Property PLCaddressClickOff() As String
        Get
            Return m_PLCaddressClick3
        End Get
        Set(ByVal value As String)
            m_PLCaddressClick3 = value
        End Set
    End Property

    '*****************************************
    '* Property - What to do to bit in PLC
    '*****************************************
    Private m_OutputType As MfgControl.AdvancedHMI.Controls.OutputType = MfgControl.AdvancedHMI.Controls.OutputType.MomentarySet
    <System.ComponentModel.Category("PLC Properties")> _
    Public Property OutputType() As MfgControl.AdvancedHMI.Controls.OutputType
        Get
            Return m_OutputType
        End Get
        Set(ByVal value As MfgControl.AdvancedHMI.Controls.OutputType)
            m_OutputType = value
        End Set
    End Property
#End Region

#Region "Events"
    '********************************************************************
    '* When an instance is added to the form, set the comm component
    '* property. If a comm component does not exist, add one to the form
    '********************************************************************
    Protected Overrides Sub OnCreateControl()
        MyBase.OnCreateControl()

        If Me.DesignMode Then
            '********************************************************
            '* Search for AdvancedHMIDrivers.IComComponent component in parent form
            '* If one exists, set the client of this component to it
            '********************************************************
            Dim i = 0
            Dim j As Integer = Me.Parent.Site.Container.Components.Count - 1
            While _CommComponent Is Nothing And i < j
                If Me.Parent.Site.Container.Components(i).GetType.GetInterface("AdvancedHMIDrivers.IComComponent") IsNot Nothing Then _CommComponent = Me.Parent.Site.Container.Components(i)
                i += 1
            End While

            '************************************************
            '* If no comm component was found, then add one and
            '* point the CommComponent property to it
            '*********************************************
            If _CommComponent Is Nothing Then
                Me.Parent.Site.Container.Add(New AdvancedHMIDrivers.EthernetIPforPLCSLCMicroComm)
                _CommComponent = Me.Parent.Site.Container.Components(Me.Parent.Site.Container.Components.Count - 1)
            End If
        Else
            SubscribeToCommDriver()
        End If
    End Sub

    '****************************
    '* Event - Button Click
    '****************************
    Private Sub _Click1(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles AutoButton.MouseDown
        MouseDownActon(m_PLCaddressClick1)
    End Sub

    Private Sub _MouseUp1(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles AutoButton.MouseUp
        MouseUpAction(m_PLCaddressClick1)
    End Sub

    Private Sub _MouseDown2(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles HandButton.MouseDown
        MouseDownActon(m_PLCaddressClick2)
    End Sub

    Private Sub _MouseUp2(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles HandButton.MouseUp
        MouseUpAction(m_PLCaddressClick2)
    End Sub

    Private Sub _click3(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles OffButton.MouseDown
        MouseUpAction(m_PLCaddressClick3)
    End Sub

    Private Sub _MouseUp3(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles OffButton.MouseUp
        MouseDownActon(m_PLCaddressClick3)
    End Sub

    Private Sub MouseDownActon(ByVal PLCAddress As String)
        If PLCAddress <> "" Then
            Try
                Select Case m_OutputType
                    Case MfgControl.AdvancedHMI.Controls.OutputType.MomentarySet : _CommComponent.WriteData(PLCAddress, 1)
                    Case MfgControl.AdvancedHMI.Controls.OutputType.MomentaryReset : _CommComponent.WriteData(PLCAddress, 0)
                    Case MfgControl.AdvancedHMI.Controls.OutputType.SetTrue : _CommComponent.WriteData(PLCAddress, 1)
                    Case MfgControl.AdvancedHMI.Controls.OutputType.SetFalse : _CommComponent.WriteData(PLCAddress, 0)
                    Case MfgControl.AdvancedHMI.Controls.OutputType.Toggle
                        Dim CurrentValue As Boolean
                        CurrentValue = _CommComponent.ReadAny(PLCAddress)
                        If CurrentValue Then
                            _CommComponent.WriteData(PLCAddress, 0)
                        Else
                            _CommComponent.WriteData(PLCAddress, 1)
                        End If
                End Select
            Catch ex As MfgControl.AdvancedHMI.Drivers.common.PLCDriverException
                If ex.ErrorCode = 1808 Then
                    DisplayError("""" & PLCAddress & """ PLC Address not found")
                Else
                    DisplayError(ex.Message)
                End If
            End Try
        End If
    End Sub


    Private Sub MouseUpAction(ByVal PLCAddress As String)
        If PLCAddress <> "" And Enabled Then
            Try
                Select Case OutputType
                    Case MfgControl.AdvancedHMI.Controls.OutputType.MomentarySet : _CommComponent.WriteData(PLCAddress, 0)
                    Case MfgControl.AdvancedHMI.Controls.OutputType.MomentaryReset : _CommComponent.WriteData(PLCAddress, 1)
                End Select
            Catch ex As MfgControl.AdvancedHMI.Drivers.common.PLCDriverException
                If ex.ErrorCode = 1808 Then
                    DisplayError("""" & PLCAddress & """ PLC Address not found")
                Else
                    DisplayError(ex.Message)
                End If
            End Try
        End If
    End Sub




    '****************************************************************
    '* UserControl overrides dispose to clean up the component list.
    '****************************************************************
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing And _CommComponent IsNot Nothing Then
                _CommComponent.UnSubscribe(NotificationIDAutoStatus)
                _CommComponent.UnSubscribe(NotificationIDHandStatus)
                _CommComponent.UnSubscribe(NotificationIDOffStatus)
                _CommComponent.UnSubscribe(NotificationIDVisibility)
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub
#End Region

#Region "Subscribing and PLC data receiving"
    '**************************************************
    '* Subscribe to addresses in the Comm(PLC) Driver
    '**************************************************
    Private SubscribedPLCAddressAutoStatus As String
    Private SubscribedPLCAddressHandStatus As String
    Private SubscribedPLCAddressOffStatus As String
    Private SubscribedPLCAddressVisibility As String
    Private Sub SubscribeToCommDriver()
        If Not DesignMode And IsHandleCreated Then
            '*******************************
            '* Subscription for Auto Status
            '*******************************
            SubscribeTo(m_PLCaddressStatusAuto, SubscribedPLCAddressAutoStatus, NotificationIDAutoStatus, AddressOf PolledDataReturnedAutoStatus)

            '*******************************
            '* Subscription for Hand Status
            '*******************************
            SubscribeTo(m_PLCaddressStatusHand, SubscribedPLCAddressHandStatus, NotificationIDHandStatus, AddressOf PolledDataReturnedHandStatus)

            '*******************************
            '* Subscription for Off Status
            '*******************************
            SubscribeTo(m_PLCaddressStatusOff, SubscribedPLCAddressOffStatus, NotificationIDOffStatus, AddressOf PolledDataReturnedOffStatus)

            '*************************
            '* Visbility Subscription
            '*************************
            If m_PLCaddressVisibility <> "" Then
                Dim PLCaddress As String = m_PLCaddressVisibility
                If PLCaddress.ToUpper.IndexOf("NOT ") = 0 Then
                    PLCaddress = m_PLCaddressVisibility.Substring(4).Trim
                    InvertVisible = True
                Else
                    InvertVisible = False
                End If
                SubscribeTo(PLCaddress, SubscribedPLCAddressVisibility, NotificationIDVisibility, AddressOf PolledDataReturnedVisibility)
            End If
        End If
    End Sub

    '******************************************************
    '* Attempt to create a subscription to the PLC driver
    '******************************************************
    Private Sub SubscribeTo(ByVal PLCAddress As String, ByRef SubscribedPLCAddress As String, ByRef NotificationID As Integer, ByVal callBack As AdvancedHMIDrivers.IComComponent.ReturnValues)
        If SubscribedPLCAddress <> PLCAddress Then
            '* If already subscribed, but address is changed, then unsubscribe first
            If NotificationID > 0 Then
                _CommComponent.UnSubscribe(NotificationID)
            End If
            '* Is there an address to subscribe to?
            If PLCAddress <> "" Then
                Try
                    If _CommComponent IsNot Nothing Then
                        NotificationID = _CommComponent.Subscribe(PLCAddress, 1, 250, callBack)

                        '* If subscription succeedded, save the address
                        SubscribedPLCAddress = PLCAddress
                    Else
                        DisplayError("CommComponent Property not set")
                    End If
                Catch ex As MfgControl.AdvancedHMI.Drivers.common.PLCDriverException
                    '* If subscribe fails, set up for retry
                    InitializeSubscribeRetry(ex, PLCAddress)
                End Try
            End If
        End If

    End Sub

    '********************************************
    '* Show the error and start the retry time
    '********************************************
    Private Sub InitializeSubscribeRetry(ByVal ex As MfgControl.AdvancedHMI.Drivers.common.PLCDriverException, ByVal PLCAddress As String)
        If ex.ErrorCode = 1808 Then
            DisplayError("""" & PLCAddress & """ PLC Address not found")
        Else
            DisplayError(ex.Message)
        End If

        If SubscribeRetryTimer Is Nothing Then
            SubscribeRetryTimer = New Windows.Forms.Timer
            SubscribeRetryTimer.Interval = 10000
            AddHandler SubscribeRetryTimer.Tick, AddressOf Retry_Tick
        End If

        SubscribeRetryTimer.Enabled = True
    End Sub


    '********************************************
    '* Keep retrying to subscribe if it failed
    '********************************************
    Private SubscribeRetryTimer As Windows.Forms.Timer
    Private Sub Retry_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs)
        SubscribeRetryTimer.Enabled = False
        SubscribeRetryTimer.Dispose()
        SubscribeRetryTimer = Nothing

        SubscribeToCommDriver()
    End Sub

    '***************************************
    '* Call backs for returned data
    '***************************************
    Private OriginalText As String
    Private Sub PolledDataReturnedAutoStatus(ByVal Values() As String)
        Try
            If Values(0) Then
                AutoButton.BackColor = Color.Green
            Else
                AutoButton.BackColor = Color.LightGray
            End If
        Catch
            DisplayError("INVALID Auto Status RETURNED!")
        End Try
    End Sub

    Private Sub PolledDataReturnedHandStatus(ByVal Values() As String)
        Try
            If Values(0) Then
                HandButton.BackColor = Color.Green
            Else
                HandButton.BackColor = Color.LightGray
            End If
        Catch
            DisplayError("INVALID Hand Status RETURNED!")
        End Try
    End Sub

    Private Sub PolledDataReturnedOffStatus(ByVal Values() As String)
        Try
            If Values(0) Then
                OffButton.BackColor = Color.Green
            Else
                OffButton.BackColor = Color.LightGray
            End If
        Catch
            DisplayError("INVALID Off Status RETURNED!")
        End Try
    End Sub

    Private Sub PolledDataReturnedVisibility(ByVal Values() As String)
        Try
            If InvertVisible Then
                MyBase.Visible = Not CBool(Values(0))
            Else
                MyBase.Visible = Values(0)
            End If
        Catch
            DisplayError("INVALID Visibilty VALUE RETURNED!")
        End Try
    End Sub
#End Region

#Region "Error Display"
    '********************************************************
    '* Show an error via the text property for a short time
    '********************************************************
    Private WithEvents ErrorDisplayTime As System.Windows.Forms.Timer
    Private Sub DisplayError(ByVal ErrorMessage As String)
        If ErrorDisplayTime Is Nothing Then
            ErrorDisplayTime = New System.Windows.Forms.Timer
            AddHandler ErrorDisplayTime.Tick, AddressOf ErrorDisplay_Tick
            ErrorDisplayTime.Interval = 6000
        End If

        '* Save the text to return to
        If Not ErrorDisplayTime.Enabled Then
            OriginalText = Me.Text
        End If

        ErrorDisplayTime.Enabled = True

        Text = ErrorMessage
    End Sub


    '**************************************************************************************
    '* Return the text back to its original after displaying the error for a few seconds.
    '**************************************************************************************
    Private Sub ErrorDisplay_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Text = OriginalText

        If ErrorDisplayTime IsNot Nothing Then
            ErrorDisplayTime.Enabled = False
            ErrorDisplayTime.Dispose()
            ErrorDisplayTime = Nothing
        End If
    End Sub
#End Region

End Class
