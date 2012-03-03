Imports Microsoft.Kinect
Imports System.ComponentModel

'------------------------------------------------------------------------------
' <copyright file="KinectSkeletonViewer.xaml.cs" company="Microsoft">
'     Copyright (c) Microsoft Corporation.  All rights reserved.
' </copyright>
'------------------------------------------------------------------------------

Namespace Microsoft.Samples.Kinect.WpfViewers

	Public Enum ImageType
		Color
		Depth
	End Enum

	Friend Enum TrackingMode
		DefaultSystemTracking
		Closest1Player
		Closest2Player
		Sticky1Player
		Sticky2Player
		MostActive1Player
		MostActive2Player
	End Enum

	''' <summary>
	''' Interaction logic for KinectSkeletonViewer.xaml
	''' </summary>
	Partial Public Class KinectSkeletonViewer
		Inherits ImageViewer
		Implements INotifyPropertyChanged
		Private Const ActivityFalloff As Single = 0.98f
		Private ReadOnly recentActivity As New List(Of ActivityWatcher)()
		Private ReadOnly activeList As New List(Of Integer)()
		Private skeletonCanvases As List(Of KinectSkeleton)
		Private jointMappings As New List(Of Dictionary(Of JointType, JointMapping))()
		Private skeletonData() As Skeleton

		Public Sub New()
			InitializeComponent()
			Me.ShowJoints = True
			Me.ShowBones = True
			Me.ShowCenter = True
		End Sub

		Public Property ShowBones() As Boolean

		Public Property ShowJoints() As Boolean

		Public Property ShowCenter() As Boolean

		Public Property ImageType() As ImageType

		Friend Property TrackingMode() As TrackingMode

		Public Sub HideAllSkeletons()
			If Me.skeletonCanvases IsNot Nothing Then
				For Each skeletonCanvas As KinectSkeleton In Me.skeletonCanvases
					skeletonCanvas.Reset()
				Next skeletonCanvas
			End If
		End Sub

		Protected Overrides Sub OnKinectChanged(ByVal oldKinectSensor As KinectSensor, ByVal newKinectSensor As KinectSensor)
			If oldKinectSensor IsNot Nothing Then
				RemoveHandler oldKinectSensor.AllFramesReady, AddressOf KinectAllFramesReady
				Me.HideAllSkeletons()
			End If

			If newKinectSensor IsNot Nothing AndAlso newKinectSensor.Status = KinectStatus.Connected Then
				AddHandler newKinectSensor.AllFramesReady, AddressOf KinectAllFramesReady
			End If
		End Sub

		Private Sub KinectAllFramesReady(ByVal sender As Object, ByVal e As AllFramesReadyEventArgs)
			' Have we already been "shut down" by the user of this viewer, 
			' or has the SkeletonStream been disabled since this event was posted?
			If (Me.Kinect Is Nothing) OrElse Not(CType(sender, KinectSensor)).SkeletonStream.IsEnabled Then
				Return
			End If

			Dim haveSkeletonData As Boolean = False

			Using skeletonFrame As SkeletonFrame = e.OpenSkeletonFrame()
				If skeletonFrame IsNot Nothing Then
					If Me.skeletonCanvases Is Nothing Then
						Me.CreateListOfSkeletonCanvases()
					End If

					If (Me.skeletonData Is Nothing) OrElse (Me.skeletonData.Length <> skeletonFrame.SkeletonArrayLength) Then
						Me.skeletonData = New Skeleton(skeletonFrame.SkeletonArrayLength - 1){}
					End If

					skeletonFrame.CopySkeletonDataTo(Me.skeletonData)

					haveSkeletonData = True
				End If
			End Using

			If haveSkeletonData Then
				Using depthImageFrame As DepthImageFrame = e.OpenDepthImageFrame()
					If depthImageFrame IsNot Nothing Then
						Dim trackedSkeletons As Integer = 0

						For Each skeleton As Skeleton In Me.skeletonData
							Dim jointMapping As Dictionary(Of JointType, JointMapping) = Me.jointMappings(trackedSkeletons)
							jointMapping.Clear()

							Dim skeletonCanvas As KinectSkeleton = Me.skeletonCanvases(trackedSkeletons)
							trackedSkeletons += 1
							skeletonCanvas.ShowBones = Me.ShowBones
							skeletonCanvas.ShowJoints = Me.ShowJoints
							skeletonCanvas.ShowCenter = Me.ShowCenter

							' Transform the data into the correct space
							' For each joint, we determine the exact X/Y coordinates for the target view
							For Each joint As Joint In skeleton.Joints
								Dim mappedPoint As Point = Me.GetPosition2DLocation(depthImageFrame, joint.Position)
								jointMapping(joint.JointType) = New JointMapping With {.Joint = joint, .MappedPoint = mappedPoint}
							Next joint

							' Look up the center point
							Dim centerPoint As Point = Me.GetPosition2DLocation(depthImageFrame, skeleton.Position)

							' Scale the skeleton thickness
							' 1.0 is the desired size at 640 width
							Dim scale As Double = Me.RenderSize.Width \ 640

							skeletonCanvas.RefreshSkeleton(skeleton, jointMapping, centerPoint, scale)
						Next skeleton

						If ImageType = ImageType.Depth Then
							Me.ChooseTrackedSkeletons(Me.skeletonData)
						End If
					End If
				End Using
			End If
		End Sub

		Private Function GetPosition2DLocation(ByVal depthFrame As DepthImageFrame, ByVal skeletonPoint As SkeletonPoint) As Point
			Dim depthPoint As DepthImagePoint = depthFrame.MapFromSkeletonPoint(skeletonPoint)

			Select Case ImageType
				Case ImageType.Color
					Dim colorPoint As ColorImagePoint = depthFrame.MapToColorImagePoint(depthPoint.X, depthPoint.Y, Me.Kinect.ColorStream.Format)

					' map back to skeleton.Width & skeleton.Height
					Return New Point(CInt(Fix(Me.RenderSize.Width * colorPoint.X / Me.Kinect.ColorStream.FrameWidth)), CInt(Fix(Me.RenderSize.Height * colorPoint.Y / Me.Kinect.ColorStream.FrameHeight)))
				Case ImageType.Depth
					Return New Point(CInt(Me.RenderSize.Width * depthPoint.X \ depthFrame.Width), CInt(Me.RenderSize.Height * depthPoint.Y \ depthFrame.Height))
				Case Else
					Throw New ArgumentOutOfRangeException("ImageType was a not expected value: " & ImageType.ToString())
			End Select
		End Function

		Private Sub CreateListOfSkeletonCanvases()
			Me.skeletonCanvases = New List(Of KinectSkeleton) From {Me.skeletonCanvas1, Me.skeletonCanvas2, Me.skeletonCanvas3, Me.skeletonCanvas4, Me.skeletonCanvas5, Me.skeletonCanvas6}

			Me.skeletonCanvases.ForEach(Function(s) Me.jointMappings.Add(New Dictionary(Of JointType, JointMapping)()))
		End Sub

		' NOTE: The ChooseTrackedSkeletons part of the KinectSkeletonViewer would be useful
		' separate from the SkeletonViewer.
		Private Sub ChooseTrackedSkeletons(ByVal skeletonDataValue As IEnumerable(Of Skeleton))
			Select Case TrackingMode
				Case TrackingMode.Closest1Player
					Me.ChooseClosestSkeletons(skeletonDataValue, 1)
				Case TrackingMode.Closest2Player
					Me.ChooseClosestSkeletons(skeletonDataValue, 2)
				Case TrackingMode.Sticky1Player
					Me.ChooseOldestSkeletons(skeletonDataValue, 1)
				Case TrackingMode.Sticky2Player
					Me.ChooseOldestSkeletons(skeletonDataValue, 2)
				Case TrackingMode.MostActive1Player
					Me.ChooseMostActiveSkeletons(skeletonDataValue, 1)
				Case TrackingMode.MostActive2Player
					Me.ChooseMostActiveSkeletons(skeletonDataValue, 2)
			End Select
		End Sub

		Private Sub ChooseClosestSkeletons(ByVal skeletonDataValue As IEnumerable(Of Skeleton), ByVal count As Integer)
			Dim depthSorted As New SortedList(Of Single, Integer)()

			For Each s As Skeleton In skeletonDataValue
				If s.TrackingState <> SkeletonTrackingState.NotTracked Then
					Dim valueZ As Single = s.Position.Z
					Do While depthSorted.ContainsKey(valueZ)
						valueZ += 0.0001f
					Loop

					depthSorted.Add(valueZ, s.TrackingId)
				End If
			Next s

			Me.ChooseSkeletonsFromList(depthSorted.Values, count)
		End Sub

		Private Sub ChooseOldestSkeletons(ByVal skeletonDataValue As IEnumerable(Of Skeleton), ByVal count As Integer)
			Dim newList As New List(Of Integer)()

			For Each s As Skeleton In skeletonDataValue
				If s.TrackingState <> SkeletonTrackingState.NotTracked Then
					newList.Add(s.TrackingId)
				End If
			Next s

			' Remove all elements from the active list that are not currently present
			Me.activeList.RemoveAll(Function(k) (Not newList.Contains(k)))

			' Add all elements that aren't already in the activeList
			Me.activeList.AddRange(newList.FindAll(Function(k) (Not Me.activeList.Contains(k))))

			Me.ChooseSkeletonsFromList(Me.activeList, count)
		End Sub

		Private Sub ChooseMostActiveSkeletons(ByVal skeletonDataValue As IEnumerable(Of Skeleton), ByVal count As Integer)
			For Each watcher As ActivityWatcher In Me.recentActivity
				watcher.NewPass()
			Next watcher

			For Each s As Skeleton In skeletonDataValue
				If s.TrackingState <> SkeletonTrackingState.NotTracked Then
					Dim watcher As ActivityWatcher = Me.recentActivity.Find(Function(w) w.TrackingId = s.TrackingId)
					If watcher IsNot Nothing Then
						watcher.Update(s)
					Else
						Me.recentActivity.Add(New ActivityWatcher(s))
					End If
				End If
			Next s

			' Remove any skeletons that are gone
			Me.recentActivity.RemoveAll(Function(aw) (Not aw.Updated))

			Me.recentActivity.Sort()
			Me.ChooseSkeletonsFromList(Me.recentActivity.ConvertAll(Function(f) f.TrackingId), count)
		End Sub

		Private Sub ChooseSkeletonsFromList(ByVal list As IList(Of Integer), ByVal max As Integer)
			If Me.Kinect.SkeletonStream.IsEnabled Then
				Dim argCount As Integer = Math.Min(list.Count, max)

				If argCount = 0 Then
					Me.Kinect.SkeletonStream.ChooseSkeletons()
				End If

				If argCount = 1 Then
					Me.Kinect.SkeletonStream.ChooseSkeletons(list(0))
				End If

				If argCount >= 2 Then
					Me.Kinect.SkeletonStream.ChooseSkeletons(list(0), list(1))
				End If
			End If
		End Sub

		Private Class ActivityWatcher
			Implements IComparable(Of ActivityWatcher)
			Private activityLevel As Single
			Private previousPosition As SkeletonPoint
			Private previousDelta As SkeletonPoint

			Friend Sub New(ByVal s As Skeleton)
				Me.activityLevel = 0.0f
				Me.TrackingId = s.TrackingId
				Me.Updated = True
				Me.previousPosition = s.Position
				Me.previousDelta = New SkeletonPoint()
			End Sub

			Private privateTrackingId As Integer
			Friend Property TrackingId() As Integer
				Get
					Return privateTrackingId
				End Get
				Private Set(ByVal value As Integer)
					privateTrackingId = value
				End Set
			End Property

			Private privateUpdated As Boolean
			Friend Property Updated() As Boolean
				Get
					Return privateUpdated
				End Get
				Private Set(ByVal value As Boolean)
					privateUpdated = value
				End Set
			End Property

			Public Function CompareTo(ByVal other As ActivityWatcher) As Integer Implements IComparable(Of ActivityWatcher).CompareTo
				' Use the existing CompareTo on float, but reverse the arguments,
				' since we wish to have larger activityLevels sort ahead of smaller values.
				Return other.activityLevel.CompareTo(Me.activityLevel)
			End Function

			Friend Sub NewPass()
				Me.Updated = False
			End Sub

			Friend Sub Update(ByVal s As Skeleton)
				Dim newPosition As SkeletonPoint = s.Position
				Dim newDelta As SkeletonPoint = New SkeletonPoint With {.X = newPosition.X - Me.previousPosition.X, .Y = newPosition.Y - Me.previousPosition.Y, .Z = newPosition.Z - Me.previousPosition.Z}

				Dim deltaV As SkeletonPoint = New SkeletonPoint With {.X = newDelta.X - Me.previousDelta.X, .Y = newDelta.Y - Me.previousDelta.Y, .Z = newDelta.Z - Me.previousDelta.Z}

				Me.previousPosition = newPosition
				Me.previousDelta = newDelta

				Dim deltaVLengthSquared As Single = (deltaV.X * deltaV.X) + (deltaV.Y * deltaV.Y) + (deltaV.Z * deltaV.Z)
				Dim deltaVLength As Single = CSng(Math.Sqrt(deltaVLengthSquared))

				Me.activityLevel = Me.activityLevel * ActivityFalloff
				Me.activityLevel += deltaVLength

				Me.Updated = True
			End Sub
		End Class
	End Class
End Namespace
