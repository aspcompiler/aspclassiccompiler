'
' Visual Basic .NET Parser
'
' Copyright (C) 2005, Microsoft Corporation. All rights reserved.
'
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER 
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF 
' MERCHANTIBILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
'

''' <summary>
''' A parse tree for an intrinsic type name.
''' </summary>
Public NotInheritable Class IntrinsicTypeName
    Inherits TypeName

    Private _IntrinsicType As IntrinsicType

    ''' <summary>
    ''' The intrinsic type.
    ''' </summary>
    Public ReadOnly Property IntrinsicType() As IntrinsicType
        Get
            Return _IntrinsicType
        End Get
    End Property

    ''' <summary>
    ''' Constructs a new intrinsic type parse tree.
    ''' </summary>
    ''' <param name="intrinsicType">The intrinsic type.</param>
    ''' <param name="span">The location of the parse tree.</param>
    Public Sub New(ByVal intrinsicType As IntrinsicType, ByVal span As Span)
        MyBase.New(TreeType.IntrinsicType, span)

        If intrinsicType < intrinsicType.Boolean OrElse intrinsicType > intrinsicType.Object Then
            Throw New ArgumentOutOfRangeException("intrinsicType")
        End If

        _IntrinsicType = intrinsicType
    End Sub
End Class