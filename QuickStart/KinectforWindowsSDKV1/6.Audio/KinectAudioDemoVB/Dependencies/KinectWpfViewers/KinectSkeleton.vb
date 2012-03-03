Imports Microsoft.Kinect

'------------------------------------------------------------------------------
' <copyright file="KinectSkeleton.cs" company="Microsoft">
'     Copyright (c) Microsoft Corporation.  All rights reserved.
' </copyright>
'------------------------------------------------------------------------------

Namespace Microsoft.Samples.Kinect.WpfViewers

	''' <summary>
	''' This control is used to render a player's skeleton.
	''' If the ClipToBounds is set to "false", it will be allowed to overdraw
	''' it's bounds.
	''' </summary>
	Public Class KinectSkeleton
		Inherits Control
		Public Shared ReadOnly ShowClippedEdgesProperty As DependencyProperty = DependencyProperty.Register("ShowClippedEdges", GetType(Boolean), GetType(KinectSkeleton), New FrameworkPropertyMetadata(True, FrameworkPropertyMetadataOptions.AffectsRender))

		Public Shared ReadOnly ShowJointsProperty As DependencyProperty = DependencyProperty.Register("ShowJoints", GetType(Boolean), GetType(KinectSkeleton), New FrameworkPropertyMetadata(True, FrameworkPropertyMetadataOptions.AffectsRender))

		Public Shared ReadOnly ShowBonesProperty As DependencyProperty = DependencyProperty.Register("ShowBones", GetType(Boolean), GetType(KinectSkeleton), New FrameworkPropertyMetadata(True, FrameworkPropertyMetadataOptions.AffectsRender))

		Public Shared ReadOnly ShowCenterProperty As DependencyProperty = DependencyProperty.Register("ShowCenter", GetType(Boolean), GetType(KinectSkeleton), New FrameworkPropertyMetadata(True, FrameworkPropertyMetadataOptions.AffectsRender))

		Private Const JointThickness As Double = 3
		Private Const BodyCenterThickness As Double = 10
		Private Const TrackedBoneThickness As Double = 6
		Private Const InferredBoneThickness As Double = 1
		Private Const ClipBoundsThickness As Double = 10

		Private ReadOnly centerPointBrush As Brush = Brushes.Blue
		Private ReadOnly trackedJointBrush As Brush = New SolidColorBrush(Color.FromArgb(255, 68, 192, 68))
		Private ReadOnly inferredJointBrush As Brush = Brushes.Yellow
		Private ReadOnly trackedBonePen As New Pen(Brushes.Green, TrackedBoneThickness)
		Private ReadOnly inferredBonePen As New Pen(Brushes.Gray, InferredBoneThickness)

		Private ReadOnly bottomClipBrush As Brush = New LinearGradientBrush(Color.FromArgb(0, 255, 0, 0), Color.FromArgb(255, 255, 0, 0), New Point(0, 0), New Point(0, 1))

		Private ReadOnly topClipBrush As Brush = New LinearGradientBrush(Color.FromArgb(0, 255, 0, 0), Color.FromArgb(255, 255, 0, 0), New Point(0, 1), New Point(0, 0))

		Private ReadOnly leftClipBrush As Brush = New LinearGradientBrush(Color.FromArgb(0, 255, 0, 0), Color.FromArgb(255, 255, 0, 0), New Point(1, 0), New Point(0, 0))

		Private ReadOnly rightClipBrush As Brush = New LinearGradientBrush(Color.FromArgb(0, 255, 0, 0), Color.FromArgb(255, 255, 0, 0), New Point(0, 0), New Point(1, 0))

		Private jointMappings As Dictionary(Of JointType, JointMapping)
		Private centerPoint As Point
		Private currentSkeleton As Skeleton
		Private scale As Double = 1.0

		Public Property ShowClippedEdges() As Boolean
			Get
				Return CBool(GetValue(ShowClippedEdgesProperty))
			End Get
			Set(ByVal value As Boolean)
				SetValue(ShowClippedEdgesProperty, value)
			End Set
		End Property

		Public Property ShowJoints() As Boolean
			Get
				Return CBool(GetValue(ShowJointsProperty))
			End Get
			Set(ByVal value As Boolean)
				SetValue(ShowJointsProperty, value)
			End Set
		End Property

		Public Property ShowBones() As Boolean
			Get
				Return CBool(GetValue(ShowBonesProperty))
			End Get
			Set(ByVal value As Boolean)
				SetValue(ShowBonesProperty, value)
			End Set
		End Property

		Public Property ShowCenter() As Boolean
			Get
				Return CBool(GetValue(ShowCenterProperty))
			End Get
			Set(ByVal value As Boolean)
				SetValue(ShowCenterProperty, value)
			End Set
		End Property

		''' <summary>
		''' This method should be called every skeleton frame.  It will force the properties to update and 
		''' trigger the control to render.
		''' </summary>
		''' <param name="skeleton">This is the current frame's skeleton data.</param>
		''' <param name="mappings">This is a list of pre-mapped joints.  See KinectSkeletonViewerer.xaml.cs</param>
		''' <param name="center">This is a pre-mapped point to the skeleton's center position.</param>
		''' <param name="scaleFactor">1 will render the bones and joints at a good size for a 640x480 image.
		''' The method would expect 0.5 to render a 320x240 scaled image.</param>
		Public Sub RefreshSkeleton(ByVal skeleton As Skeleton, ByVal mappings As Dictionary(Of JointType, JointMapping), ByVal center As Point, ByVal scaleFactor As Double)
			Me.centerPoint = center
			Me.jointMappings = mappings
			Me.scale = scaleFactor
			Me.currentSkeleton = skeleton

			Me.InvalidateVisual()
		End Sub

		Public Sub Reset()
			Me.currentSkeleton = Nothing
			Me.InvalidateVisual()
		End Sub

		Protected Overrides Function MeasureOverride(ByVal constraint As Size) As Size
			Return New Size()
		End Function

		Protected Overrides Function ArrangeOverride(ByVal arrangeBounds As Size) As Size
			Return arrangeBounds
		End Function

		Protected Overrides Sub OnRender(ByVal drawingContext As DrawingContext)
			MyBase.OnRender(drawingContext)

			' Don't render if we don't have a skeleton, or it isn't tracked
			If Me.currentSkeleton Is Nothing OrElse Me.currentSkeleton.TrackingState = SkeletonTrackingState.NotTracked Then
				Return
			End If

			' Displays a gradient near the edge of the display where the skeleton is leaving the screen
			Me.RenderClippedEdges(drawingContext)

			Select Case Me.currentSkeleton.TrackingState
				Case SkeletonTrackingState.PositionOnly
					If Me.ShowCenter Then
						drawingContext.DrawEllipse(Me.centerPointBrush, Nothing, Me.centerPoint, BodyCenterThickness * Me.scale, BodyCenterThickness * Me.scale)
					End If

				Case SkeletonTrackingState.Tracked
					Me.DrawBonesAndJoints(drawingContext)
			End Select
		End Sub

		Private Sub RenderClippedEdges(ByVal drawingContext As DrawingContext)
			If (Not Me.ShowClippedEdges) OrElse Me.currentSkeleton.ClippedEdges.Equals(FrameEdges.None) Then
				Return
			End If

			Dim scaledThickness As Double = ClipBoundsThickness * Me.scale
			If Me.currentSkeleton.ClippedEdges.HasFlag(FrameEdges.Bottom) Then
				drawingContext.DrawRectangle(Me.bottomClipBrush, Nothing, New Rect(0, Me.RenderSize.Height - scaledThickness, Me.RenderSize.Width, scaledThickness))
			End If

			If Me.currentSkeleton.ClippedEdges.HasFlag(FrameEdges.Top) Then
				drawingContext.DrawRectangle(Me.topClipBrush, Nothing, New Rect(0, 0, Me.RenderSize.Width, scaledThickness))
			End If

			If Me.currentSkeleton.ClippedEdges.HasFlag(FrameEdges.Left) Then
				drawingContext.DrawRectangle(Me.leftClipBrush, Nothing, New Rect(0, 0, scaledThickness, Me.RenderSize.Height))
			End If

			If Me.currentSkeleton.ClippedEdges.HasFlag(FrameEdges.Right) Then
				drawingContext.DrawRectangle(Me.rightClipBrush, Nothing, New Rect(Me.RenderSize.Width - scaledThickness, 0, scaledThickness, Me.RenderSize.Height))
			End If
		End Sub

		Private Sub DrawBonesAndJoints(ByVal drawingContext As DrawingContext)
			If Me.ShowBones Then
				' Render Torso
				Me.DrawBone(drawingContext, JointType.Head, JointType.ShoulderCenter)
				Me.DrawBone(drawingContext, JointType.ShoulderCenter, JointType.ShoulderLeft)
				Me.DrawBone(drawingContext, JointType.ShoulderCenter, JointType.ShoulderRight)
				Me.DrawBone(drawingContext, JointType.ShoulderCenter, JointType.Spine)
				Me.DrawBone(drawingContext, JointType.Spine, JointType.HipCenter)
				Me.DrawBone(drawingContext, JointType.HipCenter, JointType.HipLeft)
				Me.DrawBone(drawingContext, JointType.HipCenter, JointType.HipRight)

				' Left Arm
				Me.DrawBone(drawingContext, JointType.ShoulderLeft, JointType.ElbowLeft)
				Me.DrawBone(drawingContext, JointType.ElbowLeft, JointType.WristLeft)
				Me.DrawBone(drawingContext, JointType.WristLeft, JointType.HandLeft)

				' Right Arm
				Me.DrawBone(drawingContext, JointType.ShoulderRight, JointType.ElbowRight)
				Me.DrawBone(drawingContext, JointType.ElbowRight, JointType.WristRight)
				Me.DrawBone(drawingContext, JointType.WristRight, JointType.HandRight)

				' Left Leg
				Me.DrawBone(drawingContext, JointType.HipLeft, JointType.KneeLeft)
				Me.DrawBone(drawingContext, JointType.KneeLeft, JointType.AnkleLeft)
				Me.DrawBone(drawingContext, JointType.AnkleLeft, JointType.FootLeft)

				' Right Leg
				Me.DrawBone(drawingContext, JointType.HipRight, JointType.KneeRight)
				Me.DrawBone(drawingContext, JointType.KneeRight, JointType.AnkleRight)
				Me.DrawBone(drawingContext, JointType.AnkleRight, JointType.FootRight)
			End If

			If Me.ShowJoints Then
				' Render Joints
				For Each joint As JointMapping In Me.jointMappings.Values
					Dim drawBrush As Brush = Nothing
					Select Case joint.Joint.TrackingState
						Case JointTrackingState.Tracked
							drawBrush = Me.trackedJointBrush
						Case JointTrackingState.Inferred
							drawBrush = Me.inferredJointBrush
					End Select

					If drawBrush IsNot Nothing Then
						drawingContext.DrawEllipse(drawBrush, Nothing, joint.MappedPoint, JointThickness * Me.scale, JointThickness * Me.scale)
					End If
				Next joint
			End If
		End Sub

		Private Sub DrawBone(ByVal drawingContext As DrawingContext, ByVal jointType1 As JointType, ByVal jointType2 As JointType)
			Dim joint1 As JointMapping
			Dim joint2 As JointMapping

			' If we can't find either of these joints, exit
			If (Not Me.jointMappings.TryGetValue(jointType1, joint1)) OrElse joint1.Joint.TrackingState = JointTrackingState.NotTracked OrElse (Not Me.jointMappings.TryGetValue(jointType2, joint2)) OrElse joint2.Joint.TrackingState = JointTrackingState.NotTracked Then
				Return
			End If

			' Don't draw if both points are inferred
			If joint1.Joint.TrackingState = JointTrackingState.Inferred AndAlso joint2.Joint.TrackingState = JointTrackingState.Inferred Then
				Return
			End If

			' We assume all drawn bones are inferred unless BOTH joints are tracked
			Dim drawPen As Pen = Me.inferredBonePen
			drawPen.Thickness = InferredBoneThickness * Me.scale
			If joint1.Joint.TrackingState = JointTrackingState.Tracked AndAlso joint2.Joint.TrackingState = JointTrackingState.Tracked Then
				drawPen = Me.trackedBonePen
				drawPen.Thickness = TrackedBoneThickness * Me.scale
			End If

			drawingContext.DrawLine(drawPen, joint1.MappedPoint, joint2.MappedPoint)
		End Sub
	End Class
End Namespace
