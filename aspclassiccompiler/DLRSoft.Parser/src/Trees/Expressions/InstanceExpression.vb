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
''' A parse tree for Me, MyBase or MyClass.
''' </summary>
Public NotInheritable Class InstanceExpression
    Inherits Expression

    Private _InstanceType As InstanceType

    ''' <summary>
    ''' The type of the instance expression.
    ''' </summary>
    Public ReadOnly Property InstanceType() As InstanceType
        Get
            Return _InstanceType
        End Get
    End Property

    ''' <summary>
    ''' Constructs a new parse tree for My, MyBase or MyClass.
    ''' </summary>
    ''' <param name="instanceType">The type of the instance expression.</param>
    ''' <param name="span">The location of the parse tree.</param>
    Public Sub New(ByVal instanceType As InstanceType, ByVal span As Span)
        MyBase.New(TreeType.InstanceExpression, span)

        If instanceType < instanceType.Me OrElse instanceType > instanceType.MyBase Then
            Throw New ArgumentOutOfRangeException("instanceType")
        End If

        _InstanceType = instanceType
    End Sub
End Class