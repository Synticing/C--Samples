Imports Microsoft.Kinect

'------------------------------------------------------------------------------
' <copyright file="JointMapping.cs" company="Microsoft">
'     Copyright (c) Microsoft Corporation.  All rights reserved.
' </copyright>
'------------------------------------------------------------------------------

Namespace Microsoft.Samples.Kinect.WpfViewers

	''' <summary>
	''' This class is used to map points between skeleton and color/depth
	''' </summary>
	Public Class JointMapping
		''' <summary>
		''' This is the joint we are looking at
		''' </summary>
		Public Property Joint() As Joint

		''' <summary>
		''' This is the point mapped into the target displays
		''' </summary>
		Public Property MappedPoint() As Point
	End Class
End Namespace
