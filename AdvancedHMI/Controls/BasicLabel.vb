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
'* 31-DEC-11 Added BooleanDisplay property
'* 28-SEP-12 Catch specific PLCDriverException when trying to subscribe
'* 29-JAN-13 Added KeypadMinValue and KeypadMaxValue
'****************************************************************************
Public Class BasicLabel
    Inherits System.Windows.Forms.Label


#Region "Basic Properties"
    Private SavedBackColor As System.Drawing.Color

    '******************************************************************************************
    '* Use the base control's text property and make it visible as a property on the designer
    '******************************************************************************************
    <System.ComponentModel.Browsable(True)> _
<System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Visible)> _
Public Overrides Property Text() As String
        Get
            Return MyBase.Text
        End Get
        Set(ByVal value As String)
            '* True/False comes from driver, change if BooleanDisplay is different 31-DEC-11
            If (value = "True" Or value = "False") And m_BooleanDisplay <> BooleanDisplayOption.TrueFalse Then
                If value = "True" Then
                    If m_BooleanDisplay = BooleanDisplayOption.OnOff Then value = "On"
                    If m_BooleanDisplay = BooleanDisplayOption.YesNo Then value = "Yes"
                Else
                    If m_BooleanDisplay = BooleanDisplayOption.OnOff Then value = "Off"
                    If m_BooleanDisplay = BooleanDisplayOption.YesNo Then value = "No"
                End If
            End If

            '* If suffix has already been added, then removed 17-OCT-11
            If _Suffix IsNot Nothing AndAlso _Suffix <> "" AndAlso value.IndexOf(_Suffix) > 0 Then value = value.Substring(0, value.IndexOf(_Suffix))

            If m_Format <> "" And (Not DesignMode) Then
                Try
                    '* 31-MAY-13
                    Dim v As Single
                    If Single.TryParse(value, v) Then
                        MyBase.Text = _Prefix & Format(CSng(v) * m_ScaleFactor, m_Format) & _Suffix
                    End If
                Catch exC As InvalidCastException
                    If MyBase.Text = "Check NumericFormat and variable type" Then
                        MyBase.Text = "----"
                    Else
                        MyBase.Text = value
                    End If
                Catch ex As Exception
                    MyBase.Text = "Check NumericFormat and variable type"
                End Try
            Else
                '* Highlight in red if a Highlightcharacter found mark is in text
                If InStr(value, _HighlightKeyChar) > 0 Then
                    If MyBase.BackColor <> _Highlightcolor Then SavedBackColor = MyBase.BackColor
                    MyBase.BackColor = _Highlightcolor
                Else
                    If SavedBackColor <> Nothing Then MyBase.BackColor = SavedBackColor
                End If

                If m_ScaleFactor = 1 Then
                    MyBase.Text = _Prefix & value & _Suffix
                Else
                    Try
                        MyBase.Text = value * m_ScaleFactor
                    Catch ex As Exception
                        DisplayError("Scale Factor Error - " & ex.Message)
                    End Try
                End If
            End If
        End Set
    End Property

    '**********************************
    '* Prefix and suffixes to text
    '**********************************
    Private _Prefix As String
    Public Property TextPrefix() As String
        Get
            Return _Prefix
        End Get
        Set(ByVal value As String)
            _Prefix = value
            Invalidate()
        End Set
    End Property

    Private _Suffix As String
    Public Property TextSuffix() As String
        Get
            Return _Suffix
        End Get
        Set(ByVal value As String)
            _Suffix = value
            Invalidate()
        End Set
    End Property


    '***************************************************************
    '* Property - Highlight Color
    '***************************************************************
    Private _Highlightcolor As Drawing.Color = Drawing.Color.Red
    <System.ComponentModel.Category("Appearance")> _
    Public Property HighlightColor() As Drawing.Color
        Get
            Return _Highlightcolor
        End Get
        Set(ByVal value As Drawing.Color)
            _Highlightcolor = value
        End Set
    End Property

    Private _HighlightKeyChar As String = "!"
    <System.ComponentModel.Category("Appearance")> _
    Public Property HighlightKeyCharacter() As String
        Get
            Return _HighlightKeyChar
        End Get
        Set(ByVal value As String)
            _HighlightKeyChar = value
        End Set
    End Property


    Private m_Format As String
    Public Property NumericFormat() As String
        Get
            Return m_Format
        End Get
        Set(ByVal value As String)
            m_Format = value
        End Set
    End Property

    Private m_ScaleFactor As Decimal = 1
    Public Property ScaleFactor() As Decimal
        Get
            Return m_ScaleFactor
        End Get
        Set(ByVal value As Decimal)
            m_ScaleFactor = value
            'TODO: Does not refresh in designmode
            'Text = MyBase.Text
        End Set
    End Property

    Public Enum BooleanDisplayOption
        TrueFalse
        YesNo
        OnOff
    End Enum

    Private m_BooleanDisplay As BooleanDisplayOption
    Public Property BooleanDisplay() As BooleanDisplayOption
        Get
            Return m_BooleanDisplay
        End Get
        Set(ByVal value As BooleanDisplayOption)
            m_BooleanDisplay = value
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

    Private _PollRate As Integer
    Public Property PollRate() As Integer
        Get
            Return _PollRate
        End Get
        Set(ByVal value As Integer)
            _PollRate = value
        End Set
    End Property

    Private m_KeypadText As String
    Public Property KeypadText() As String
        Get
            Return m_KeypadText
        End Get
        Set(ByVal value As String)
            m_KeypadText = value
        End Set
    End Property

    Private m_KeypadFont As Font = New Font("Arial", 10)
    Public Property KeypadFont() As Font
        Get
            Return m_KeypadFont
        End Get
        Set(ByVal value As Font)
            m_KeypadFont = value
        End Set
    End Property

    Private m_KeypadForeColor As Color = Color.WhiteSmoke
    Public Property KeypadFontColor() As Color
        Get
            Return m_KeypadForeColor
        End Get
        Set(ByVal value As Color)
            m_KeypadForeColor = value
        End Set
    End Property

    Private m_KeypadWidth As Integer = 300
    Public Property KeypadWidth() As Integer
        Get
            Return m_KeypadWidth
        End Get
        Set(ByVal value As Integer)
            m_KeypadWidth = value
        End Set
    End Property

    '* 29-JAN-13
    Private m_KeypadMinValue As Double
    Public Property KeypadMinValue As Double
        Get
            Return m_KeypadMinValue
        End Get
        Set(value As Double)
            m_KeypadMinValue = value
        End Set
    End Property

    Private m_KeypadMaxValue As Double
    Public Property KeypadMaxValue As Double
        Get
            Return m_KeypadMaxValue
        End Get
        Set(value As Double)
            m_KeypadMaxValue = value
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
    '* Property - Address in PLC to Write Data To
    '*****************************************
    Private m_PLCaddressEntry As String = ""
    <System.ComponentModel.Category("PLC Properties")> _
    Public Property PLCaddressEntry() As String
        Get
            Return m_PLCaddressEntry
        End Get
        Set(ByVal value As String)
            If m_PLCaddressEntry <> value Then
                m_PLCaddressEntry = value
            End If
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

    Public Sub New()
        MyBase.new()

        'If Me.Parent.BackColor = System.Drawing.Color.Black And _
        If (MyBase.ForeColor = System.Drawing.Color.FromKnownColor(System.Drawing.KnownColor.ControlText) Or ForeColor = Color.FromArgb(0, 0, 0, 0)) Then
            ForeColor = System.Drawing.Color.WhiteSmoke
        End If
    End Sub

    '****************************************************************
    '* UserControl overrides dispose to clean up the component list.
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

                If KeypadPopUp IsNot Nothing Then
                    KeypadPopUp.Dispose()
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
            Catch ex As MfgControl.AdvancedHMI.Drivers.Common.PLCDriverException
                'Me.Text = ex.Message
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
            DisplayError("INVALID VALUE RETURNED!")
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
    Private Sub ErrorDisplay_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Text = OriginalText

        If ErrorDisplayTime IsNot Nothing Then
            ErrorDisplayTime.Enabled = False
            ErrorDisplayTime.Dispose()
            ErrorDisplayTime = Nothing
        End If
    End Sub
#End Region

#Region "Keypad popup for data entry"
    Private WithEvents KeypadPopUp As MfgControl.AdvancedHMI.Controls.Keypad

    Private Sub KeypadPopUp_ButtonClick(ByVal sender As Object, ByVal e As MfgControl.AdvancedHMI.Controls.KeyPadEventArgs) Handles KeypadPopUp.ButtonClick
        If e.Key = "Quit" Then
            KeypadPopUp.Visible = False
        ElseIf e.Key = "Enter" Then
            If _CommComponent Is Nothing Then
                DisplayError("CommComponent Property not set")
            Else
                If KeypadPopUp.Value <> "" Then
                    '* 29-JAN-13 - Validate value if a Min/ax was specified
                    Try
                        If m_KeypadMaxValue <> m_KeypadMinValue Then
                            If KeypadPopUp.Value < m_KeypadMinValue Or KeypadPopUp.Value > m_KeypadMaxValue Then
                                MsgBox("Value must be >" & m_KeypadMinValue & " and <" & m_KeypadMaxValue)
                                Exit Sub
                            End If
                        End If
                    Catch ex As Exception
                        MsgBox("Failed to validate value. " & ex.Message)
                        Exit Sub
                    End Try
                    Try
                        '* 29-JAN-13 - reduced code and checked for divide by 0
                        If ScaleFactor <> 0 Then
                            _CommComponent.WriteData(m_PLCaddressEntry, KeypadPopUp.Value / ScaleFactor)
                        End If
                    Catch ex As Exception
                        MsgBox("Failed to write value - " & ex.Message)
                    End Try
                Else
                    'DisplayError("CommComponent Property not set")
                End If
                KeypadPopUp.Visible = False
            End If
        End If
    End Sub

    '***********************************************************
    '* If labeled is clicked, pop up a keypad for data entry
    '***********************************************************
    Private Sub BasicLabelWithEntry_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Click
        If m_PLCaddressEntry <> "" And Enabled Then
            If KeypadPopUp Is Nothing Then
                KeypadPopUp = New MfgControl.AdvancedHMI.Controls.Keypad(m_KeypadWidth)
                KeypadPopUp.StartPosition = Windows.Forms.FormStartPosition.CenterScreen
                KeypadPopUp.TopMost = True
            End If

            '***************************
            '*Set the font and forecolor
            '****************************
            'KeypadPopUp.Font = New Font("Arial", 16, FontStyle.Bold, GraphicsUnit.Point)
            If m_KeypadFont IsNot Nothing Then KeypadPopUp.Font = m_KeypadFont
            'If m_KeypadForeColor IsNot Nothing Then
            KeypadPopUp.ForeColor = m_KeypadForeColor
            'End If


            KeypadPopUp.Text = m_KeypadText
            KeypadPopUp.Value = ""
            KeypadPopUp.Show()
        End If
    End Sub
#End Region
End Class
