Imports System.Drawing
'****************************************************************************
'* 12-DEC-08 Added line in OnPaint to exit is LegendText is nothing
'* 05-OCT-09 Exit OnPaint if GrayPen is Nothing
'****************************************************************************
Public Class GraphicIndicator
    Inherits MfgControl.AdvancedHMI.Controls.GraphicIndicator

    Private StaticImage As Bitmap
    Private TextRect As New Rectangle
    Private ImageRatio As Single

#Region "Properties"

    '    <System.ComponentModel.Browsable(True)> _
    '<System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Visible)> _
    'Public Overrides Property Text() As String
    '        Get
    '            Return MyBase.Text
    '        End Get
    '        Set(ByVal value As String)
    '            MyBase.Text = value
    '        End Set
    '    End Property

    '    'Private m_font As New Font("Arial", 22, FontStyle.Regular, GraphicsUnit.Point)
    '    '* These next properties are overriden so that we can refresh the image when changed
    '    <System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Visible)> _
    '    Public Overrides Property Font() As Font
    '        Get
    '            Return MyBase.Font
    '        End Get
    '        Set(ByVal value As Font)
    '            MyBase.Font = value
    '        End Set
    '    End Property


    '*****************************************************
    '* Property - Component to communicate to PLC through
    '*****************************************************
    Private _CommComponent As AdvancedHMIDrivers.IComComponent
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
    Private m_PLCaddressSelect1 As String = ""
    Public Property PLCaddressSelect1() As String
        Get
            Return m_PLCaddressSelect1
        End Get
        Set(ByVal value As String)
            If m_PLCaddressSelect1 <> value Then
                m_PLCaddressSelect1 = value
                '* When address is changed, re-subscribe to new address
                SubscribeToCommDriver()
            End If
        End Set
    End Property

    '*****************************************
    '* Property - Address in PLC to Link to
    '*****************************************
    Private m_PLCaddressSelect2 As String = ""
    Public Property PLCaddressSelect2() As String
        Get
            Return m_PLCaddressSelect2
        End Get
        Set(ByVal value As String)
            If m_PLCaddressSelect2 <> value Then
                m_PLCaddressSelect2 = value
                '* When address is changed, re-subscribe to new address
                SubscribeToCommDriver()
            End If
        End Set
    End Property

    '*****************************************
    '* Property - Address in PLC to Link to
    '*****************************************
    Private m_PLCaddressVisibility As String = ""
    Private InvertVisible As Boolean
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
    Private m_PLCaddressText2 As String = ""
    Public Property PLCaddressText2() As String
        Get
            Return m_PLCaddressText2
        End Get
        Set(ByVal value As String)
            If m_PLCaddressText2 <> value Then
                m_PLCaddressText2 = value
                '* When address is changed, re-subscribe to new address
                SubscribeToCommDriver()
            End If
        End Set
    End Property

    '*****************************************
    '* Property - Address in PLC to Link to
    '*****************************************
    Private m_PLCaddressClick As String = ""
    Public Property PLCaddressClick() As String
        Get
            Return m_PLCaddressClick
        End Get
        Set(ByVal value As String)
            m_PLCaddressClick = value
        End Set
    End Property

#End Region


#Region "Events"
    'UserControl overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
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

    '****************************
    '* Event - Mouse Down
    '****************************
    Private Sub MomentaryButton_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseDown
        If m_PLCaddressClick <> "" Then
            Try
                Select Case OutputType
                    Case OutputTypes.MomentarySet : _CommComponent.WriteData(m_PLCaddressClick, 1)
                    Case OutputTypes.MomentaryReset : _CommComponent.WriteData(m_PLCaddressClick, 0)
                    Case OutputTypes.SetTrue : _CommComponent.WriteData(m_PLCaddressClick, 1)
                    Case OutputTypes.SetFalse : _CommComponent.WriteData(m_PLCaddressClick, 0)
                    Case OutputTypes.Toggle
                        Dim CurrentValue As Boolean
                        CurrentValue = _CommComponent.ReadSynchronous(m_PLCaddressClick, 1)(0)
                        If CurrentValue Then
                            _CommComponent.WriteData(m_PLCaddressClick, 0)
                        Else
                            _CommComponent.WriteData(m_PLCaddressClick, 1)
                        End If
                End Select
            Catch
            End Try
        End If
    End Sub


    '****************************
    '* Event - Mouse Up
    '****************************
    Private Sub MomentaryButton_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseUp
        If m_PLCaddressClick <> "" Then
            Try
                Select Case OutputType
                    Case OutputTypes.MomentarySet : _CommComponent.WriteData(m_PLCaddressClick, 0)
                    Case OutputTypes.MomentaryReset : _CommComponent.WriteData(m_PLCaddressClick, 1)
                End Select
                'tmrError.Enabled = False
            Catch
            End Try
        End If
    End Sub

    Private Sub GraphicIndicator_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.TextChanged
        Me.Invalidate()
    End Sub


    '* This is part of the transparent background code and it stops flicker
    'Protected Overrides Sub OnPaintBackground(ByVal e As System.Windows.Forms.PaintEventArgs)
    '    'MyBase.OnPaintBackground(e)
    'End Sub


    '********************************************************************
    '* When an instance is added to the form, set the comm component
    '* property. If a comm component does not exist, add one to the form
    '********************************************************************
    Protected Overrides Sub OnCreateControl()
        MyBase.OnCreateControl()

        If Me.DesignMode Then
            'If Me Is Nothing OrElse Me.Parent Is Nothing OrElse Me.Parent.Site Is Nothing Then Exit Sub

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


    '**************************************************
    '* Subscribe to addresses in the Comm(PLC) Driver
    '**************************************************
    Private Sub SubscribeToCommDriver()
        If Not DesignMode And IsHandleCreated Then
            '*************************
            '* Text Subscription
            '*************************
            SubscribeTo(m_PLCaddressText2, AddressOf PolledDataReturnedText2)

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
            '* Select1 Subscription
            '*************************
            SubscribeTo(m_PLCaddressSelect1, AddressOf PolledDataReturnedSelect1)

            '*************************
            '* Select2 Subscription
            '*************************
            SubscribeTo(m_PLCaddressSelect2, AddressOf PolledDataReturnedSelect2)
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



    Private OriginalText, OriginalText2 As String
    Private Sub PolledDataReturnedSelect1(ByVal Values() As String)
        Try
            ValueSelect1 = Values(0)
        Catch
            DisplayError("Invalid Select1 Value")
        End Try
    End Sub


    Private Sub PolledDataReturnedSelect2(ByVal Values() As String)
        Try
            ValueSelect2 = Values(0)
        Catch
            DisplayError("Invalid Select2 Value")
        End Try
    End Sub

    '************************************************************
    Private Sub PolledDataReturnedVisibility(ByVal Values() As String)
        Try
            Me.Visible = (Values(0) <> 0)
            'LegendText = Values(0)
        Catch
            If Values(0).Length < 10 Then
                DisplayError("INVALID Visibility VALUE!")
            Else
                DisplayError("Invalid Visibility Value")
            End If
        End Try
    End Sub

    '************************************************************
    Private Sub PolledDataReturnedText2(ByVal Values() As String)
        Try
            Text2 = Values(0)
            'LegendText = Values(0)
        Catch
            DisplayError("Invalid Text2 Value")
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
