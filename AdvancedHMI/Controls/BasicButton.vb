'*****************************************************************************************************
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
'* 23-NOV-11 Added MinimumHoldTime property to allow bit in PLC to be set for longer than button held
'* 23-NOV-11 Added TextAlternate and SelectTextAlternate properties to allow single bit selection of text
'* 21-FEB-11 Added MaximumHoldTime property to ensure button always releases even if windows misses
'* 29-MAY-12 Fixed a problem in MaxHoldTime property where it stored to MinHoldTime
'* 10-JAN-13 Added "AndAlso Me.Parent.Site IsNot Nothing" to OnCreateControl to prevent startup error
'* 10-MAY-13 Added ValueToWrite property
'******************************************************************************************************
Public Class BasicButton
    Inherits System.Windows.Forms.Button


#Region "Basic Properties"
    '***************************************************************
    '* Property - Highlight Color
    '***************************************************************
    Private _Highlightcolor As Drawing.Color = Drawing.Color.Green
    Public Property HighlightColor() As Drawing.Color
        Get
            Return _Highlightcolor
        End Get
        Set(ByVal value As Drawing.Color)
            _Highlightcolor = value
        End Set
    End Property

    '***************************************************************
    '* Property - If value from PLC is true, then highlight button
    '***************************************************************
    Private OriginalBackcolor As Drawing.Color
    Private _Highlight As Boolean
    Public Property Highlight() As Boolean
        Get
            Return _Highlight
        End Get
        Set(ByVal value As Boolean)
            If OriginalBackcolor.R = 0 And OriginalBackcolor.G = 0 And OriginalBackcolor.B = 0 Then
                OriginalBackcolor = MyBase.BackColor
            End If

            If value Then
                MyBase.BackColor = _Highlightcolor
            Else
                MyBase.BackColor = OriginalBackcolor
            End If

            _Highlight = value
        End Set
    End Property

    '******************************
    '* Property - Text
    '******************************
    Private m_Text As String
    Public Shadows Property Text() As String
        Get
            Return m_Text
        End Get
        Set(ByVal value As String)
            m_Text = value

            If m_SelectTextAlternate Then
                MyBase.Text = m_TextAlternate
            Else
                MyBase.Text = m_Text
            End If
        End Set
    End Property

    '******************************
    '* Property - Alternate Text
    '******************************
    Private m_TextAlternate As String
    Public Property TextAlternate() As String
        Get
            Return m_TextAlternate
        End Get
        Set(ByVal value As String)
            m_TextAlternate = value

            If m_SelectTextAlternate Then
                MyBase.Text = m_TextAlternate
            Else
                MyBase.Text = m_Text
            End If

        End Set
    End Property

    '***********************************
    '* Property - Select Alternate Text
    '***********************************
    Private m_SelectTextAlternate As Boolean
    Public Property SelectTextAlternate() As Boolean
        Get
            Return m_SelectTextAlternate
        End Get
        Set(ByVal value As Boolean)
            If value <> m_SelectTextAlternate Then
                m_SelectTextAlternate = value
                If value Then
                    MyBase.Text = m_TextAlternate
                Else
                    MyBase.Text = m_Text
                End If
                'Me.Invalidate()
            End If
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

    '********************************************
    '* Property - Address in PLC for click event
    '********************************************
    Private _PLCaddressClick As String = ""
    <System.ComponentModel.Category("PLC Properties")> _
    Public Property PLCaddressClick() As String
        Get
            Return _PLCaddressClick
        End Get
        Set(ByVal value As String)
            _PLCaddressClick = value
        End Set
    End Property

    '*****************************************
    '* Property - Address in PLC to Link to
    '*****************************************
    Private _PLCaddressHighlight As String = ""
    <System.ComponentModel.Category("PLC Properties")> _
    Public Property PLCaddressHighlight() As String
        Get
            Return _PLCaddressHighlight
        End Get
        Set(ByVal value As String)
            If _PLCaddressHighlight <> value Then
                _PLCaddressHighlight = value

                '* When address is changed, re-subscribe to new address
                SubscribeToCommDriver()
            End If
        End Set
    End Property

    '*****************************************
    '* Property - Address in PLC to Link to
    '*****************************************
    Private InvertVisible As Boolean
    Private _PLCaddressVisibility As String = ""
    <System.ComponentModel.Category("PLC Properties")> _
    Public Property PLCaddressVisibility() As String
        Get
            Return _PLCaddressVisibility
        End Get
        Set(ByVal value As String)
            If _PLCaddressVisibility <> value Then
                _PLCaddressVisibility = value

                '* When address is changed, re-subscribe to new address
                SubscribeToCommDriver()
            End If
        End Set
    End Property

    '*****************************************
    '* Property - Address in PLC to Link to
    '*****************************************
    Private m_PLCaddressSelectTextAlternate As String = ""
    <System.ComponentModel.Category("PLC Properties")> _
    Public Property PLCaddressSelectTextAlternate() As String
        Get
            Return m_PLCaddressSelectTextAlternate
        End Get
        Set(ByVal value As String)
            If m_PLCaddressSelectTextAlternate <> value Then
                m_PLCaddressSelectTextAlternate = value

                '* When address is changed, re-subscribe to new address
                SubscribeToCommDriver()
            End If
        End Set
    End Property


    '*****************************************
    '* Property - What to do to bit in PLC
    '*****************************************
    Private _OutputType As MfgControl.AdvancedHMI.Controls.OutputType = MfgControl.AdvancedHMI.Controls.OutputType.MomentarySet
    <System.ComponentModel.Category("PLC Properties")> _
    Public Property OutputType() As MfgControl.AdvancedHMI.Controls.OutputType
        Get
            Return _OutputType
        End Get
        Set(ByVal value As MfgControl.AdvancedHMI.Controls.OutputType)
            _OutputType = value
        End Set
    End Property

    '*****************************************
    '* Property - Hold time before bit reset
    '*****************************************
    Private WithEvents MinHoldTimer As New System.Windows.Forms.Timer
    Private m_MinimumHoldTime As Integer = 500
    <System.ComponentModel.Category("PLC Properties")> _
    Public Property MinimumHoldTime() As Integer
        Get
            Return m_MinimumHoldTime
        End Get
        Set(ByVal value As Integer)
            m_MinimumHoldTime = value
            If value > 0 Then MinHoldTimer.Interval = value
        End Set
    End Property

    '*****************************************
    '* Property - Hold time before bit reset
    '*****************************************
    Private WithEvents MaxHoldTimer As New System.Windows.Forms.Timer
    Private m_MaximumHoldTime As Integer = 3000
    <System.ComponentModel.Category("PLC Properties")> _
    Public Property MaximumHoldTime() As Integer
        Get
            Return m_MaximumHoldTime
        End Get
        Set(ByVal value As Integer)
            m_MaximumHoldTime = value
            If value > 0 Then MaxHoldTimer.Interval = value
        End Set
    End Property

    '**********************************************************************
    '* If output type is set to write value, then write this value to PLC
    '**********************************************************************
    Private m_ValueToWrite As Integer
    <System.ComponentModel.Category("PLC Properties")> _
    Public Property ValueToWrite As Integer
        Get
            Return m_ValueToWrite
        End Get
        Set(value As Integer)
            m_ValueToWrite = value
        End Set
    End Property
#End Region

#Region "Events"
    '****************************
    '* Event - Mouse Down
    '****************************
    Private Sub MomentaryButton_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseDown
        MouseIsDown = True
        HoldTimeMet = False

        If _PLCaddressClick <> "" Then
            Try
                Select Case _OutputType
                    Case MfgControl.AdvancedHMI.Controls.OutputType.MomentarySet
                        _CommComponent.WriteData(_PLCaddressClick, 1)
                        If m_MinimumHoldTime > 0 Then MinHoldTimer.Enabled = True
                        If m_MaximumHoldTime > 0 Then MaxHoldTimer.Enabled = True
                    Case MfgControl.AdvancedHMI.Controls.OutputType.MomentaryReset
                        _CommComponent.WriteData(_PLCaddressClick, 0)
                        If m_MinimumHoldTime > 0 Then MinHoldTimer.Enabled = True
                        If m_MaximumHoldTime > 0 Then MaxHoldTimer.Enabled = True
                    Case MfgControl.AdvancedHMI.Controls.OutputType.SetTrue : _CommComponent.WriteData(_PLCaddressClick, 1)
                    Case MfgControl.AdvancedHMI.Controls.OutputType.SetFalse : _CommComponent.WriteData(_PLCaddressClick, 0)
                    Case MfgControl.AdvancedHMI.Controls.OutputType.Toggle
                        Dim CurrentValue As Boolean
                        CurrentValue = _CommComponent.ReadSynchronous(_PLCaddressClick, 1)(0)
                        If CurrentValue Then
                            _CommComponent.WriteData(_PLCaddressClick, 0)
                        Else
                            _CommComponent.WriteData(_PLCaddressClick, 1)
                        End If
                    Case MfgControl.AdvancedHMI.Controls.OutputType.WriteValue
                        _CommComponent.WriteData(_PLCaddressClick, m_ValueToWrite)
                End Select
            Catch ex As MfgControl.AdvancedHMI.Drivers.Common.PLCDriverException
                If ex.ErrorCode = 1808 Then
                    DisplayError("""" & _PLCaddressClick & """ PLC Address not found")
                Else
                    DisplayError(ex.Message)
                End If
            End Try
        End If
    End Sub


    '****************************
    '* Event - Mouse Up
    '****************************
    Private Sub MomentaryButton_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseUp
        MouseIsDown = False
        If _PLCaddressClick <> "" Then
            If HoldTimeMet Or m_MinimumHoldTime <= 0 Then
                MaxHoldTimer.Enabled = False
                ReleaseValue()
            End If
        End If
    End Sub

    Private Sub ReleaseValue()
        Try
            Select Case _OutputType
                Case MfgControl.AdvancedHMI.Controls.OutputType.MomentarySet : _CommComponent.WriteData(_PLCaddressClick, 0)
                Case MfgControl.AdvancedHMI.Controls.OutputType.MomentaryReset : _CommComponent.WriteData(_PLCaddressClick, 1)
            End Select
        Catch ex As MfgControl.AdvancedHMI.Drivers.Common.PLCDriverException
            If ex.ErrorCode = 1808 Then
                DisplayError("""" & _PLCaddressClick & """ PLC Address not found")
            Else
                DisplayError(ex.Message)
            End If
        End Try
    End Sub

    Private MouseIsDown, HoldTimeMet As Boolean
    Private Sub HoldTimer_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles MinHoldTimer.Tick
        MinHoldTimer.Enabled = False
        HoldTimeMet = True
        If Not MouseIsDown Then
            ReleaseValue()
        End If
    End Sub

    Private Sub MaxHoldTimer_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles MaxHoldTimer.Tick
        MaxHoldTimer.Enabled = False
        ReleaseValue()
    End Sub



    '********************************************************************
    '* When an instance is added to the form, set the comm component
    '* property. If a comm component does not exist, add one to the form
    '********************************************************************
    Protected Overrides Sub OnCreateControl()
        MyBase.OnCreateControl()


        If Me.DesignMode AndAlso Me.Parent.Site IsNot Nothing Then
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

    '****************************************************************
    '* UserControl overrides dispose to clean up the component list.
    '****************************************************************
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing Then
                If _CommComponent IsNot Nothing Then
                    '* Unsubscribe from all
                    For i As Integer = 0 To SuccessfulSubscriptions.Count - 1
                        _CommComponent.Unsubscribe(SuccessfulSubscriptions(i).NotificationID)
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
            '* Highlight Subscription
            '*************************
            SubscribeTo(_PLCaddressHighlight, AddressOf PolledDataReturnedHighlight)

            '************************************
            '* SelectAlternate Text Subscription
            '************************************
            SubscribeTo(m_PLCaddressSelectTextAlternate, AddressOf PolledDataReturnedSelectTextAlternate)

            '*************************
            '* Visbility Subscription
            '*************************
            If _PLCaddressVisibility <> "" Then
                Dim PLCaddress As String = _PLCaddressVisibility
                If PLCaddress.ToUpper.IndexOf("NOT ") = 0 Then
                    PLCaddress = _PLCaddressVisibility.Substring(4).Trim
                    InvertVisible = True
                Else
                    InvertVisible = False
                End If
                SubscribeTo(PLCaddress, AddressOf PolledDataReturnedVisibility)
            End If
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


    Private SubscribeRetryTimer As Windows.Forms.Timer
    Private Sub Retry_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs)
        SubscribeToCommDriver()
    End Sub

    '********************************************
    '* Show the error and start the retry time
    '********************************************
    Private Sub InitializeSubscribeRetry(ByVal ex As MfgControl.AdvancedHMI.Drivers.Common.PLCDriverException, ByVal PLCAddress As String)
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


    '***************************************
    '* Call backs for returned data
    '***************************************
    Private OriginalText As String
    Private Sub PolledDataReturnedHighlight(ByVal Values() As String)
        Try
            If _OutputType = AdvancedHMI.Controls.OutputType.WriteValue Then
                Highlight = (Values(0) = m_ValueToWrite)
            Else
                Highlight = Values(0)
            End If
        Catch
            DisplayError("INVALID Highlight VALUE RETURNED!")
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

    Private Sub PolledDataReturnedBackColorChange(ByVal Values() As String)
        Try
            If Values(0) Then
                MyBase.BackColor = Color.Lime
            Else
                MyBase.BackColor = Color.LightGray
            End If
        Catch
            DisplayError("INVALID BackColorChange VALUE RETURNED!")
        End Try
    End Sub

    Private Sub PolledDataReturnedSelectTextAlternate(ByVal Values() As String)
        Try
            SelectTextAlternate = Values(0)
        Catch
            DisplayError("INVALID SelectAlternateText VALUE RETURNED!")
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

        MyBase.Text = ErrorMessage
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