Imports Microsoft.Kinect

'------------------------------------------------------------------------------
' <copyright file="KinectDepthViewer.xaml.cs" company="Microsoft">
'     Copyright (c) Microsoft Corporation.  All rights reserved.
' </copyright>
'------------------------------------------------------------------------------

Namespace Microsoft.Samples.Kinect.WpfViewers


	''' <summary>
	''' Interaction logic for KinectDepthViewer.xaml
	''' </summary>
	Partial Public Class KinectDepthViewer
		Inherits ImageViewer
		' color divisors for tinting depth pixels
		Private Shared ReadOnly IntensityShiftByPlayerR() As Integer = { 1, 2, 0, 2, 0, 0, 2, 0 }
		Private Shared ReadOnly IntensityShiftByPlayerG() As Integer = { 1, 2, 2, 0, 2, 0, 0, 1 }
		Private Shared ReadOnly IntensityShiftByPlayerB() As Integer = { 1, 0, 2, 2, 0, 2, 0, 2 }

		Private Const RedIndex As Integer = 2
		Private Const GreenIndex As Integer = 1
		Private Const BlueIndex As Integer = 0
		Private Shared ReadOnly Bgr32BytesPerPixel As Integer = (PixelFormats.Bgr32.BitsPerPixel + 7) / 8

		Private lastImageFormat As DepthImageFormat
		Private pixelData() As Short

		' We want to control how depth data gets converted into false-color data
		' for more intuitive visualization, so we keep 32-bit color frame buffer versions of
		' these, to be updated whenever we receive and process a 16-bit frame.
		Private depthFrame32() As Byte
		Private outputBitmap As WriteableBitmap

		Public Sub New()
			InitializeComponent()
		End Sub

		Protected Overrides Sub OnKinectChanged(ByVal oldKinectSensor As KinectSensor, ByVal newKinectSensor As KinectSensor)
			If oldKinectSensor IsNot Nothing Then
				RemoveHandler oldKinectSensor.DepthFrameReady, AddressOf DepthImageReady
				kinectDepthImage.Source = Nothing
				Me.lastImageFormat = DepthImageFormat.Undefined
			End If

			If newKinectSensor IsNot Nothing AndAlso newKinectSensor.Status = KinectStatus.Connected Then
				ResetFrameRateCounters()

				AddHandler newKinectSensor.DepthFrameReady, AddressOf DepthImageReady
			End If
		End Sub

		Private Sub DepthImageReady(ByVal sender As Object, ByVal e As DepthImageFrameReadyEventArgs)
			Using imageFrame As DepthImageFrame = e.OpenDepthImageFrame()
				If imageFrame IsNot Nothing Then
					' We need to detect if the format has changed.
					Dim haveNewFormat As Boolean = Me.lastImageFormat <> imageFrame.Format

					If haveNewFormat Then
						Me.pixelData = New Short(imageFrame.PixelDataLength - 1){}
						Me.depthFrame32 = New Byte(imageFrame.Width * imageFrame.Height * Bgr32BytesPerPixel - 1){}
					End If

					imageFrame.CopyPixelDataTo(Me.pixelData)

					Dim convertedDepthBits() As Byte = Me.ConvertDepthFrame(Me.pixelData, (CType(sender, KinectSensor)).DepthStream)

					' A WriteableBitmap is a WPF construct that enables resetting the Bits of the image.
					' This is more efficient than creating a new Bitmap every frame.
					If haveNewFormat Then
						Me.outputBitmap = New WriteableBitmap(imageFrame.Width, imageFrame.Height, 96, 96, PixelFormats.Bgr32, Nothing) ' DpiY -  DpiX

						Me.kinectDepthImage.Source = Me.outputBitmap
					End If

					Me.outputBitmap.WritePixels(New Int32Rect(0, 0, imageFrame.Width, imageFrame.Height), convertedDepthBits, imageFrame.Width * Bgr32BytesPerPixel, 0)

					Me.lastImageFormat = imageFrame.Format

					UpdateFrameRate()
				End If
			End Using
		End Sub

		' Converts a 16-bit grayscale depth frame which includes player indexes into a 32-bit frame
		' that displays different players in different colors
		Private Function ConvertDepthFrame(ByVal depthFrame() As Short, ByVal depthStream As DepthImageStream) As Byte()
			Dim tooNearDepth As Integer = depthStream.TooNearDepth
			Dim tooFarDepth As Integer = depthStream.TooFarDepth
			Dim unknownDepth As Integer = depthStream.UnknownDepth

			Dim i16 As Integer = 0
			Dim i32 As Integer = 0
			Do While i16 < depthFrame.Length AndAlso i32 < Me.depthFrame32.Length
				Dim player As Integer = depthFrame(i16) And DepthImageFrame.PlayerIndexBitmask
				Dim realDepth As Integer = depthFrame(i16) >> DepthImageFrame.PlayerIndexBitmaskWidth




				' transform 13-bit depth information into an 8-bit intensity appropriate
				' for display (we disregard information in most significant bit)
				Dim intensity As Byte = CByte(Not(realDepth >> 4))

				If player = 0 AndAlso realDepth = 0 Then
					' white 
					Me.depthFrame32(i32 + RedIndex) = 255
					Me.depthFrame32(i32 + GreenIndex) = 255
					Me.depthFrame32(i32 + BlueIndex) = 255
				ElseIf player = 0 AndAlso realDepth = tooFarDepth Then
					' dark purple
					Me.depthFrame32(i32 + RedIndex) = 66
					Me.depthFrame32(i32 + GreenIndex) = 0
					Me.depthFrame32(i32 + BlueIndex) = 66
				ElseIf player = 0 AndAlso realDepth = unknownDepth Then
					' dark brown
					Me.depthFrame32(i32 + RedIndex) = 66
					Me.depthFrame32(i32 + GreenIndex) = 66
					Me.depthFrame32(i32 + BlueIndex) = 33
				Else
					' tint the intensity by dividing by per-player values
					Me.depthFrame32(i32 + RedIndex) = CByte(intensity >> IntensityShiftByPlayerR(player))
					Me.depthFrame32(i32 + GreenIndex) = CByte(intensity >> IntensityShiftByPlayerG(player))
					Me.depthFrame32(i32 + BlueIndex) = CByte(intensity >> IntensityShiftByPlayerB(player))
				End If

				i16 += 1
				i32 += 4
			Loop

			Return Me.depthFrame32
		End Function
	End Class
End Namespace
