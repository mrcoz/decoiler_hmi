'**********************************************************************************************
'* Selector Switch for AdvancedHMI
'*
'* Archie Jacobs
'* Manufacturing Automation, LLC
'* ajacobs@mfgcontrol.com
'* 
'*
'* Copyright 2011 Archie Jacobs
'*
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
'**********************************************************************************************

Public Class SelectorSwitch3Pos
    Inherits MfgControl.AdvancedHMI.Controls.SelectorSwitch3Pos


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
    Private m_PLCaddressValueLeft As String = ""
    <System.ComponentModel.Category("PLC Properties")> _
    Public Property PLCaddressValueLeft() As String
        Get
            Return m_PLCaddressValueLeft
        End Get
        Set(ByVal value As String)
            If m_PLCaddressValueLeft <> value Then
                m_PLCaddressValueLeft = value

                '* When address is changed, re-subscribe to new address
                SubscribeToCommDriver()
            End If
        End Set
    End Property

    '*****************************************
    '* Property - Address in PLC to Link to
    '*****************************************
    Private m_PLCaddressValueRight As String = ""
    <System.ComponentModel.Category("PLC Properties")> _
    Public Property PLCaddressValueRight() As String
        Get
            Return m_PLCaddressValueRight
        End Get
        Set(ByVal value As String)
            If m_PLCaddressValueRight <> value Then
                m_PLCaddressValueRight = value

                '* When address is changed, re-subscribe to new address
                SubscribeToCommDriver()
            End If
        End Set
    End Property

    '*****************************************
    '* Property - Address in PLC to Link to
    '*****************************************
    Private m_PLCaddressClickLeft As String = ""
    <System.ComponentModel.Category("PLC Properties")> _
    Public Property PLCaddressClickLeft() As String
        Get
            Return m_PLCaddressClickLeft
        End Get
        Set(ByVal value As String)
            If m_PLCaddressClickLeft <> value Then
                m_PLCaddressClickLeft = value

                '* When address is changed, re-subscribe to new address
                SubscribeToCommDriver()
            End If
        End Set
    End Property

    '*****************************************
    '* Property - Address in PLC to Link to
    '*****************************************
    Private m_PLCaddressClickRight As String = ""
    <System.ComponentModel.Category("PLC Properties")> _
    Public Property PLCaddressClickRight() As String
        Get
            Return m_PLCaddressClickRight
        End Get
        Set(ByVal value As String)
            If m_PLCaddressClickRight <> value Then
                m_PLCaddressClickRight = value

                '* When address is changed, re-subscribe to new address
                SubscribeToCommDriver()
            End If
        End Set
    End Property
#End Region


#Region "Events"
    '****************************************************************
    '* UserControl overrides dispose to clean up the component list.
    '* Release all subscriptions
    '****************************************************************
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing Then
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

    '* This is part of the transparent background code and it stops flicker
    'Protected Overrides Sub OnPaintBackground(ByVal e As System.Windows.Forms.PaintEventArgs)
    'MyBase.OnPaintBackground(e)
    'End Sub

    Private WithEvents tmrError As New System.Windows.Forms.Timer

    Private Sub ClickedLeft() Handles MyBase.SwitchLeft
        Try
            If m_PLCaddressClickLeft <> "" Then
                _CommComponent.WriteData(m_PLCaddressClickLeft, 1)
            End If

            If m_PLCaddressClickRight <> "" Then
                _CommComponent.WriteData(m_PLCaddressClickRight, 0)
            End If
        Catch ex As Exception
            DisplayError("WRITE FAILED!" & ex.Message)
        End Try
    End Sub

    Private Sub ClickedCenter() Handles MyBase.SwitchCenter
        Try
            If m_PLCaddressClickLeft <> "" Then
                _CommComponent.WriteData(m_PLCaddressClickLeft, 0)
            End If

            If m_PLCaddressClickRight <> "" Then
                _CommComponent.WriteData(m_PLCaddressClickRight, 0)
            End If
        Catch ex As Exception
            DisplayError("WRITE FAILED!" & ex.Message)
        End Try
    End Sub

    Private Sub ClickedRight() Handles MyBase.SwitchRight
        Try
            If m_PLCaddressClickLeft <> "" Then
                _CommComponent.WriteData(m_PLCaddressClickLeft, 0)
            End If

            If m_PLCaddressClickRight <> "" Then
                _CommComponent.WriteData(m_PLCaddressClickRight, 1)
            End If
        Catch ex As Exception
            DisplayError("WRITE FAILED!" & ex.Message)
        End Try
    End Sub


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
            If _CommComponent Is Nothing And (PLCaddressClickLeft <> "" Or PLCaddressClickRight <> "") Then
                Me.Parent.Site.Container.Add(New AdvancedHMIDrivers.EthernetIPforPLCSLCMicroComm)
                _CommComponent = Me.Parent.Site.Container.Components(Me.Parent.Site.Container.Components.Count - 1)
            End If
        Else
            SubscribeToCommDriver()
        End If
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
                SubscribeTo(PLCaddress, AddressOf PolledDataReturnedVisibility)
            End If

            '*************************
            '* Value Subscriptions
            '*************************
            SubscribeTo(m_PLCaddressValueLeft, AddressOf PolledDataReturnedValueLeft)
            SubscribeTo(m_PLCaddressValueRight, AddressOf PolledDataReturnedValueRight)
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
            _CommComponent.Unsubscribe(SuccessfulSubscriptions(index).NotificationID)
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
            Catch ex As MfgControl.AdvancedHMI.Drivers.Common.PLCDriverException
                '* If subscribe fails, set up for retry
                InitializeSubscribeRetry(ex, PLCAddress)
            End Try
        End If
    End Sub


    '********************************************
    '* Show the error and start the retry time
    '********************************************
    Private Sub InitializeSubscribeRetry(ByVal ex As MfgControl.AdvancedHMI.Drivers.Common.PLCDriverException, ByVal PLCAddress As String)
        '* Error 1808 is from TwinCAT ADS
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
    Private Sub PolledDataReturnedText(ByVal Values() As String)
        Try
            Text = Values(0)
        Catch
            DisplayError(Values(0) & "-INVALID Text VALUE RETURNED!")
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
            DisplayError(Values(0) & "-INVALID Visibilty VALUE RETURNED!")
        End Try
    End Sub


    Private Sub PolledDataReturnedValueLeft(ByVal Values() As String)
        Try
            If Values(0) Then
                Value = 0
            ElseIf Value = 0 Then
                '* Move back to center if we are left and left drops out
                Value = 1
            End If
        Catch
            DisplayError(Values(0) & "-INVALID Value RETURNED!")
        End Try
    End Sub

    Private Sub PolledDataReturnedValueRight(ByVal Values() As String)
        Try
            If Values(0) Then
                Value = 2
            ElseIf Value = 2 Then
                '* Move back to center if we are right and right drops out
                Value = 1
            End If
        Catch
            DisplayError(Values(0) & "-INVALID Value RETURNED!")
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
            ErrorDisplayTime.Interval = 5000
        End If

        '* Save the text to return to
        If Not ErrorDisplayTime.Enabled Then
            OriginalText = Me.Text
        End If

        ErrorDisplayTime.Enabled = True

        Me.Text = ErrorMessage
    End Sub


    '**************************************************************************************
    '* Return the text back to its original after displaying the error for a few seconds.
    '**************************************************************************************
    Private Sub ErrorDisplay_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ErrorDisplayTime.Tick
        Text = OriginalText

        If ErrorDisplayTime IsNot Nothing Then
            ErrorDisplayTime.Enabled = False
            ErrorDisplayTime.Dispose()
            ErrorDisplayTime = Nothing
        End If
    End Sub
#End Region

End Class
