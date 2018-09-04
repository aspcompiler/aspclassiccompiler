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
''' A collection of trees that are colon delimited.
''' </summary>
Public MustInherit Class ColonDelimitedTreeCollection(Of T As Tree)
    Inherits TreeCollection(Of T)

    Private ReadOnly _ColonLocations As ReadOnlyCollection(Of Location)

    ''' <summary>
    ''' The locations of the colons in the collection.
    ''' </summary>
    Public ReadOnly Property ColonLocations() As ReadOnlyCollection(Of Location)
        Get
            Return _ColonLocations
        End Get
    End Property

    Protected Sub New(ByVal type As TreeType, ByVal trees As IList(Of T), ByVal colonLocations As IList(Of Location), ByVal span As Span)
        MyBase.New(type, trees, span)

        Debug.Assert(type = TreeType.StatementCollection OrElse type = TreeType.DeclarationCollection)

        If colonLocations IsNot Nothing AndAlso colonLocations.Count > 0 Then
            _ColonLocations = New ReadOnlyCollection(Of Location)(colonLocations)
        End If
    End Sub
End Class
