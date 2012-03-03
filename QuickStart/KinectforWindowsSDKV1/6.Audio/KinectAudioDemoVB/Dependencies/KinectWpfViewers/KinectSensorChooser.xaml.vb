Imports Microsoft.Kinect
Imports System.Windows.Media.Animation
Imports System.Globalization
Imports System.ComponentModel

'------------------------------------------------------------------------------
' <copyright file="KinectSensorChooser.xaml.cs" company="Microsoft">
'     Copyright (c) Microsoft Corporation.  All rights reserved.
' </copyright>
'------------------------------------------------------------------------------

Namespace Microsoft.Samples.Kinect.WpfViewers


	''' <summary>
	''' Interaction logic for KinectSensorChooser.xaml
	''' </summary>
	Partial Public Class KinectSensorChooser
		Inherits UserControl
		Public Shared ReadOnly KinectProperty As DependencyProperty = DependencyProperty.Register("Kinect", GetType(KinectSensor), GetType(KinectSensorChooser), New UIPropertyMetadata(Nothing, New PropertyChangedCallback(AddressOf KinectPropertyChanged)))

		Public Shared ReadOnly MessageProperty As DependencyProperty = DependencyProperty.Register("Message", GetType(String), GetType(KinectSensorChooser), New UIPropertyMetadata(Nothing))

		Public Shared ReadOnly MoreInfoProperty As DependencyProperty = DependencyProperty.Register("MoreInfo", GetType(String), GetType(KinectSensorChooser), New UIPropertyMetadata(Nothing))

		Public Shared ReadOnly MoreInfoUriProperty As DependencyProperty = DependencyProperty.Register("MoreInfoUri", GetType(Uri), GetType(KinectSensorChooser), New UIPropertyMetadata(Nothing))

		Public Shared ReadOnly ShowRetryProperty As DependencyProperty = DependencyProperty.Register("ShowRetry", GetType(Boolean), GetType(KinectSensorChooser), New UIPropertyMetadata(False))

		Public Shared ReadOnly KinectSensorChangedEvent As RoutedEvent = EventManager.RegisterRoutedEvent("KinectSensorChanged", RoutingStrategy.Bubble, GetType(DependencyPropertyChangedEventHandler), GetType(ImageViewer))

		Private sensorConflict As Boolean

		Public Sub New()
			InitializeComponent()
			AddHandler Me.Loaded, AddressOf KinectSensorChooserLoaded

			Me.IsRequired = True

			' Setup bindings via code
			Dim binding As New Binding("Message") With {.Source = Me}
			MessageTextBlock.SetBinding(TextBlock.TextProperty, binding)
			Dim binding2 As New Binding("MoreInfo") With {.Source = Me}
			TellMeMoreLink.SetBinding(TextBlock.ToolTipProperty, binding2)
			Dim binding3 As New Binding("MoreInfo") With {.Source = Me, .Converter = New NullToVisibilityConverter()}
			TellMeMore.SetBinding(TextBlock.VisibilityProperty, binding3)
			Dim binding4 As New Binding("ShowRetry") With {.Source = Me, .Converter = New BoolToVisibilityConverter()}
			RetryButton.SetBinding(Button.VisibilityProperty, binding4)
			Dim binding5 As New Binding("MoreInfoUri") With {.Source = Me}
			TellMeMoreLink.SetBinding(Hyperlink.NavigateUriProperty, binding5)

			Me.UpdateMessage(KinectStatus.Undefined, "Required", "This application needs a Kinect for Windows sensor in order to function. Please plug one into the PC.", New Uri("http://go.microsoft.com/fwlink/?LinkID=239815"), False)
		End Sub

		Public Event KinectSensorChanged As DependencyPropertyChangedEventHandler

		Public Property IsRequired() As Boolean

		Public Property Kinect() As KinectSensor
			Get
				Return CType(GetValue(KinectProperty), KinectSensor)
			End Get
			Set(ByVal value As KinectSensor)
				SetValue(KinectProperty, value)
			End Set
		End Property

		Public Property Message() As String
			Get
				Return CStr(GetValue(MessageProperty))
			End Get
			Set(ByVal value As String)
				SetValue(MessageProperty, value)
			End Set
		End Property

		Public Property MoreInfo() As String
			Get
				Return CStr(GetValue(MoreInfoProperty))
			End Get
			Set(ByVal value As String)
				SetValue(MoreInfoProperty, value)
			End Set
		End Property

		Public Property MoreInfoUri() As Uri
			Get
				Return CType(GetValue(MoreInfoUriProperty), Uri)
			End Get
			Set(ByVal value As Uri)
				SetValue(MoreInfoUriProperty, value)
			End Set
		End Property

		Public Property ShowRetry() As Boolean
			Get
				Return CBool(GetValue(ShowRetryProperty))
			End Get
			Set(ByVal value As Boolean)
				SetValue(ShowRetryProperty, value)
			End Set
		End Property

		Public Sub AppConflictOccurred()
			If Me.Kinect IsNot Nothing Then
				Me.sensorConflict = True

				' sensorConflict is considered when handling status updates,
				' so we call UpdateStatus with the current Status to ensure that the
				' logic takes the sensorConflict into account.
				Me.UpdateStatus(Me.Kinect.Status)
			End If
		End Sub

		Private Shared Sub KinectPropertyChanged(ByVal d As DependencyObject, ByVal args As DependencyPropertyChangedEventArgs)
			Dim sensorChooser As KinectSensorChooser = CType(d, KinectSensorChooser)
			RaiseEvent sensorChooser.KinectSensorChanged(sensorChooser, args)
		End Sub

		Private Sub KinectSensorChooserLoaded(ByVal sender As Object, ByVal e As RoutedEventArgs)
			AddHandler KinectSensor.KinectSensors.StatusChanged, AddressOf KinectSensorsStatusChanged
			Me.DiscoverSensor()
		End Sub

		Private Sub DiscoverSensor()
			' When this control is running in a designer, it should not discover a KinectSensor
			If Not DesignerProperties.GetIsInDesignMode(Me) Then
				' Find first sensor that is connected.
				For Each sensor As KinectSensor In KinectSensor.KinectSensors
					If sensor.Status = KinectStatus.Connected Then
						Me.UpdateStatus(sensor.Status)
						Me.Kinect = sensor
						Exit For
					End If
				Next sensor

				' If we didn't find a connected Sensor
				If Me.Kinect Is Nothing Then
					' NOTE: this doesn't handle the multiple Kinect sensor case very well.
					For Each sensor As KinectSensor In KinectSensor.KinectSensors
						Me.UpdateStatus(sensor.Status)
					Next sensor
				End If
			End If
		End Sub

		Private Sub KinectSensorsStatusChanged(ByVal sender As Object, ByVal e As StatusChangedEventArgs)
			Dim status = e.Status
			If Me.Kinect Is Nothing Then
				Me.UpdateStatus(status)
				If e.Status = KinectStatus.Connected Then
					Me.Kinect = e.Sensor
				End If
			Else
				If Me.Kinect Is e.Sensor Then
					Me.UpdateStatus(status)
					If e.Status = KinectStatus.Disconnected OrElse e.Status = KinectStatus.NotPowered Then
						Me.Kinect = Nothing
						Me.sensorConflict = False
						Me.DiscoverSensor()
					End If
				End If
			End If
		End Sub

		Private Sub UpdateStatus(ByVal status As KinectStatus)
			Dim message As String = Nothing
			Dim moreInfo As String = Nothing
			Dim moreInfoUri As Uri = Nothing
			Dim showRetry As Boolean = False

			Select Case status
				Case KinectStatus.Connected
					' If there's a sensor conflict, we wish to display all of the normal 
					' states and statuses, with the exception of Connected.
					If Me.sensorConflict Then
						message = "This Kinect is being used by another application."
						moreInfo = "This application needs a Kinect for Windows sensor in order to function. However, another application is using the Kinect Sensor."
						moreInfoUri = New Uri("http://go.microsoft.com/fwlink/?LinkID=239812")
						showRetry = True
					Else
						message = "All set!"
						moreInfo = Nothing
						moreInfoUri = Nothing
					End If

				Case KinectStatus.DeviceNotGenuine
					message = "This sensor is not genuine!"
					moreInfo = "This application needs a genuine Kinect for Windows sensor in order to function. Please plug one into the PC."
					moreInfoUri = New Uri("http://go.microsoft.com/fwlink/?LinkID=239813")

				Case KinectStatus.DeviceNotSupported
					message = "Kinect for Xbox not supported."
					moreInfo = "This application needs a Kinect for Windows sensor in order to function. Please plug one into the PC."
					moreInfoUri = New Uri("http://go.microsoft.com/fwlink/?LinkID=239814")

				Case KinectStatus.Disconnected
					If Me.IsRequired Then
						message = "Required"
						moreInfo = "This application needs a Kinect for Windows sensor in order to function. Please plug one into the PC."
						moreInfoUri = New Uri("http://go.microsoft.com/fwlink/?LinkID=239815")
					Else
						message = "Get the full experience by plugging in a Kinect for Windows sensor."
						moreInfo = "This application will use a Kinect for Windows sensor if one is plugged into the PC."
						moreInfoUri = New Uri("http://go.microsoft.com/fwlink/?LinkID=239816")
					End If

				Case KinectStatus.NotReady, KinectStatus.Error
					message = "Oops, there is an error."
					moreInfo = "The Kinect Sensor is plugged in, however an error has occured. For steps to resolve, please click the ""Tell me more"" link."
					moreInfoUri = New Uri("http://go.microsoft.com/fwlink/?LinkID=239817")
				Case KinectStatus.Initializing
					message = "Initializing..."
					moreInfo = Nothing
					moreInfoUri = Nothing
				Case KinectStatus.InsufficientBandwidth
					message = "Too many USB devices! Please unplug one or more."
					moreInfo = "The Kinect Sensor needs the majority of the USB Bandwidth of a USB Controller. If other devices are in contention for that bandwidth, the Kinect Sensor may not be able to function."
					moreInfoUri = New Uri("http://go.microsoft.com/fwlink/?LinkID=239818")
				Case KinectStatus.NotPowered
					message = "Plug my power cord in!"
					moreInfo = "The Kinect Sensor is plugged into the computer with its USB connection, but the power plug appears to be not powered."
					moreInfoUri = New Uri("http://go.microsoft.com/fwlink/?LinkID=239819")
			End Select

			Me.UpdateMessage(status, message, moreInfo, moreInfoUri, showRetry)
		End Sub

		Private Sub UpdateMessage(ByVal status As KinectStatus, ByVal message As String, ByVal moreInfo As String, ByVal moreInfoUri As Uri, ByVal showRetry As Boolean)
			Me.Message = message
			Me.MoreInfo = moreInfo
			Me.MoreInfoUri = moreInfoUri
			Me.ShowRetry = showRetry

			If (status = KinectStatus.Connected) AndAlso (Not Me.sensorConflict) Then
				Dim fadeAnimation = New DoubleAnimation(0, New Duration(TimeSpan.FromMilliseconds(500)))

				AddHandler fadeAnimation.Completed, Sub(sender, args)
					' If we've reached the end of the animation and achieved total transparency, 
					' the chooser should no longer be clickable - hide it.
					If Me.Opacity = 0 Then
						Me.Visibility = Visibility.Hidden
					End If
				End Sub

				Me.BeginAnimation(UserControl.OpacityProperty, fadeAnimation, HandoffBehavior.SnapshotAndReplace)
			Else
				' The chooser is heading towards opaque - as long as it's not completely transparent, 
				' it should be Visible and clickable.
				Me.Visibility = Visibility.Visible

				Dim fadeAnimation = New DoubleAnimation(1.0, New Duration(TimeSpan.FromMilliseconds(500)))
				Me.BeginAnimation(UserControl.OpacityProperty, fadeAnimation, HandoffBehavior.SnapshotAndReplace)
			End If
		End Sub

		Private Sub RetryButtonClick(ByVal sender As Object, ByVal e As RoutedEventArgs)
			Me.sensorConflict = False
			Dim sensorToRetry = Me.Kinect

			' Necessary to null the Kinect value first. Otherwise, no property change will be detected.
			Me.Kinect = Nothing
			Me.Kinect = sensorToRetry

			Me.UpdateStatus(KinectStatus.Connected)
		End Sub

		Private Sub TellMeMoreLinkRequestNavigate(ByVal sender As Object, ByVal e As RequestNavigateEventArgs)
			Dim hyperlink As Hyperlink = TryCast(e.OriginalSource, Hyperlink)
			If hyperlink IsNot Nothing Then
				' Careful - ensure that this NavigateUri comes from a trusted source, as in this sample, before launching a process using it.
				Process.Start(New ProcessStartInfo(hyperlink.NavigateUri.ToString()))
			End If

			e.Handled = True
		End Sub
	End Class

	Public Class NullToVisibilityConverter
		Implements IValueConverter
		Public Function Convert(ByVal value As Object, ByVal targetType As Type, ByVal parameter As Object, ByVal culture As CultureInfo) As Object Implements IValueConverter.Convert
			Dim isVisible As Boolean = value IsNot Nothing
			Return If(isVisible, Visibility.Visible, Visibility.Collapsed)
		End Function

		Public Function ConvertBack(ByVal value As Object, ByVal targetType As Type, ByVal parameter As Object, ByVal culture As CultureInfo) As Object Implements IValueConverter.ConvertBack
			Throw New NotImplementedException()
		End Function
	End Class

	Public Class BoolToVisibilityConverter
		Implements IValueConverter
		Public Function Convert(ByVal value As Object, ByVal targetType As Type, ByVal parameter As Object, ByVal culture As CultureInfo) As Object Implements IValueConverter.Convert
			Dim isVisible As Boolean = CBool(value)
			Return If(isVisible, Visibility.Visible, Visibility.Collapsed)
		End Function

		Public Function ConvertBack(ByVal value As Object, ByVal targetType As Type, ByVal parameter As Object, ByVal culture As CultureInfo) As Object Implements IValueConverter.ConvertBack
			Throw New NotImplementedException()
		End Function
	End Class
End Namespace
