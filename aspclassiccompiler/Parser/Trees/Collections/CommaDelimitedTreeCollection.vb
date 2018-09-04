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
''' A collection of trees that are delimited by commas.
''' </summary>
Public MustInherit Class CommaDelimitedTreeCollection(Of T As Tree)
    Inherits TreeCollection(Of T)

    Private ReadOnly _CommaLocations As ReadOnlyCollection(Of Location)

    ''' <summary>
    ''' The location of the commas in the list.
    ''' </summary>
    Public ReadOnly Property CommaLocations() As ReadOnlyCollection(Of Location)
        Get
            Return _CommaLocations
        End Get
    End Property

    Protected Sub New(ByVal type As TreeType, ByVal trees As IList(Of T), ByVal commaLocations As IList(Of Location), ByVal span As Span)
        MyBase.New(type, trees, span)

        Debug.Assert(type >= TreeType.ArgumentCollection AndAlso type <= TreeType.ImportCollection)

        If commaLocations IsNot Nothing AndAlso commaLocations.Count > 0 Then
            _CommaLocations = New ReadOnlyCollection(Of Location)(commaLocations)
        End If
    End Sub
End Class
