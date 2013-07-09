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
'* 12-JUN-11 Created
'****************************************************************************
Public Class TempController
    Inherits MfgControl.AdvancedHMI.Controls.TempController


#Region "Properties"
    Private SavedBackColor As System.Drawing.Color

    Private _ScaleFactor As Decimal = 1
    <System.ComponentModel.Category("Numeric Display")> _
    Public Property ScaleFactor() As Decimal
        Get
            Return _ScaleFactor
        End Get
        Set(ByVal value As Decimal)
            _ScaleFactor = value
        End Set
    End Property
#End Region

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
    Private m_PLCaddressText As String = ""
    <System.ComponentModel.Category("PLC Properties")> _
    Public Property PLCaddressText() As String
        Get
            Return m_PLCaddressText
        End Get
        Set(ByVal value As String)
            If m_PLCaddressText <> value Then
                m_PLCaddressText = value

                '* When address is changed, re-subscribe to new address
                SubscribeToCommDriver()
            End If
        End Set
    End Property

    '*****************************************
    '* Property - Address in PLC to Link to
    '*****************************************
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

    '*****************************************
    '* Property - Address in PLC to Link to
    '*****************************************
    Private m_PLCaddressValuePV As String = ""
    <System.ComponentModel.Category("PLC Properties")> _
    Public Property PLCaddressValuePV() As String
        Get
            Return m_PLCaddressValuePV
        End Get
        Set(ByVal value As String)
            If m_PLCaddressValuePV <> value Then
                m_PLCaddressValuePV = value

                '* When address is changed, re-subscribe to new address
                SubscribeToCommDriver()
            End If
        End Set
    End Property

    '*****************************************
    '* Property - Address in PLC to Link to
    '*****************************************
    Private m_PLCaddressValueSP As String = ""
    <System.ComponentModel.Category("PLC Properties")> _
    Public Property PLCaddressValueSP() As String
        Get
            Return m_PLCaddressValueSP
        End Get
        Set(ByVal value As String)
            If m_PLCaddressValueSP <> value Then
                m_PLCaddressValueSP = value

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
    Public Property PLCaddressClick1() As String
        Get
            Return m_PLCaddressClick1
        End Get
        Set(ByVal value As String)
            m_PLCaddressClick1 = value
        End Set
    End Property

    Private m_PLCaddressClick2 As String = ""
    <System.ComponentModel.Category("PLC Properties")> _
    Public Property PLCaddressClick2() As String
        Get
            Return m_PLCaddressClick2
        End Get
        Set(ByVal value As String)
            m_PLCaddressClick2 = value
        End Set
    End Property

    Private m_PLCaddressClick3 As String = ""
    <System.ComponentModel.Category("PLC Properties")> _
    Public Property PLCaddressClick3() As String
        Get
            Return m_PLCaddressClick3
        End Get
        Set(ByVal value As String)
            m_PLCaddressClick3 = value
        End Set
    End Property

    Private m_PLCaddressClick4 As String = ""
    <System.ComponentModel.Category("PLC Properties")> _
    Public Property PLCaddressClick4() As String
        Get
            Return m_PLCaddressClick4
        End Get
        Set(ByVal value As String)
            m_PLCaddressClick4 = value
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
    'Protected Overrides Sub OnHandleCreated(ByVal e As EventArgs)
    'End Sub

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
    Private Sub _Click1(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles MyBase.Button1MouseDown
        MouseDownActon(m_PLCaddressClick1)
    End Sub

    Private Sub _MouseUp1(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles MyBase.Button1MouseUp
        MouseUpAction(m_PLCaddressClick1)
    End Sub

    Private Sub _click2(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles MyBase.Button2MouseDown
        MouseDownActon(m_PLCaddressClick2)
    End Sub

    Private Sub _MouseUp2(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles MyBase.Button2MouseUp
        MouseUpAction(m_PLCaddressClick2)
    End Sub

    Private Sub _click3(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles MyBase.Button3MouseDown
        MouseDownActon(m_PLCaddressClick3)
    End Sub

    Private Sub _MouseUp3(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles MyBase.Button3MouseUp
        MouseUpAction(m_PLCaddressClick3)
    End Sub

    Private Sub _click4(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles MyBase.Button4MouseDown
        MouseDownActon(m_PLCaddressClick4)
    End Sub

    Private Sub _MouseUp4(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles MyBase.Button4MouseUp
        MouseUpAction(m_PLCaddressClick4)
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
    '* Control overrides dispose to clean up the component list.
    '****************************************************************
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing Then
                '* Unsubscribe from the subscriptions
                If _CommComponent IsNot Nothing Then
                    '* Unsubscribe from all
                    For i As Integer = 0 To SuccessfulSubscriptions.Count - 1
                        _CommComponent.UnSubscribe(SuccessfulSubscriptions(i).NotificationID)
                    Next
                    SuccessfulSubscriptions.Clear()
                End If
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
    Private Sub SubscribeToCommDriver()
        If Not DesignMode And IsHandleCreated Then
            '*************************
            '* Text Subscription
            '*************************
            SubscribeTo(m_PLCaddressText, AddressOf PolledDataReturnedText)

            '*************************
            '* Visibility Subscription
            '*************************
            If m_PLCaddressVisibility <> "" Then
                Dim PLCaddress As String = m_PLCaddressVisibility
                If PLCaddress.ToUpper.IndexOf("NOT ") = 0 Then
                    PLCaddress = m_PLCaddressVisibility.Substring(4).Trim
                    InvertVisible = True
                Else
                    InvertVisible = False
                End If
                SubscribeTo(PLCaddress, AddressOf PolledDataReturnedVisibility)
            End If

            '*************************
            '* PV Value Subscription
            '*************************
            SubscribeTo(m_PLCaddressValuePV, AddressOf PolledDataReturnedValuePV)

            '*************************
            '* ValueSP Subscription
            '*************************
            SubscribeTo(m_PLCaddressValueSP, AddressOf PolledDataReturnedValueSP)
        End If
    End Sub

    Private SuccessfulSubscriptions As New List(Of SubscriptionDetail)

    '******************************************************
    '* Attempt to create a subscription to the PLC driver
    '******************************************************
    Private Sub SubscribeTo(ByVal PLCAddress As String, ByVal callBack As AdvancedHMIDrivers.IComComponent.ReturnValues)
        '* Check to see if the subscription has already been created
        Dim index As Integer
        While index < SuccessfulSubscriptions.Count AndAlso SuccessfulSubscriptions(index).Callback <> callBack
            index += 1
        End While

        '* Already subscribed and PLCAddress was changed, so unsubscribe
        If (index < SuccessfulSubscriptions.Count) AndAlso SuccessfulSubscriptions(index).PLCAddress <> PLCAddress Then
            _CommComponent.UnSubscribe(SuccessfulSubscriptions(index).NotificationID)
            SuccessfulSubscriptions.RemoveAt(index)
        End If

        '* Is there an address to subscribe to?
        If PLCAddress <> "" Then
            Try
                If _CommComponent IsNot Nothing Then
                    Dim NotificationID As Integer = _CommComponent.Subscribe(PLCAddress, 1, 250, callBack)

                    '* If subscription succeedded, save the subscription details
                    Dim temp As New SubscriptionDetail(PLCAddress, NotificationID, callBack)
                    SuccessfulSubscriptions.Add(temp)
                Else
                    DisplayError("CommComponent Property not set")
                End If
            Catch ex As MfgControl.AdvancedHMI.Drivers.common.PLCDriverException
                '* If subscribe fails, set up for retry
                InitializeSubscribeRetry(ex, PLCAddress)
            End Try
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
            SubscribeRetryTimer = New System.Timers.Timer
            SubscribeRetryTimer.Interval = 10000
            AddHandler SubscribeRetryTimer.Elapsed, AddressOf Retry_Tick
        End If

        SubscribeRetryTimer.Enabled = True
    End Sub


    '********************************************
    '* Keep retrying to subscribe if it failed
    '********************************************
    Private SubscribeRetryTimer As System.Timers.Timer
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
    Private Sub PolledDataReturnedText(ByVal Values() As String)
        Try
            Text = Values(0)
        Catch
            DisplayError("INVALID Text VALUE RETURNED!")
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


    Private Sub PolledDataReturnedValuePV(ByVal Values() As String)
        Try
            MyBase.ValuePV = Values(0)
        Catch
            DisplayError("INVALID VALUE RETURNED!" & Values(0))
        End Try
    End Sub

    Private Sub PolledDataReturnedValueSP(ByVal Values() As String)
        Try
            MyBase.ValueSP = Values(0)
        Catch
            DisplayError("INVALID VALUEsp RETURNED!" & Values(0))
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
