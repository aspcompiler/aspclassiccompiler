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
''' A region marked in the source code.
''' </summary>
Public NotInheritable Class SourceRegion
    Private ReadOnly _Start, _Finish As Location
    Private ReadOnly _Description As String

    ''' <summary>
    ''' The start location of the region.
    ''' </summary>
    Public ReadOnly Property Start() As Location
        Get
            Return _Start
        End Get
    End Property

    ''' <summary>
    ''' The end location of the region.
    ''' </summary>
    Public ReadOnly Property Finish() As Location
        Get
            Return _Finish
        End Get
    End Property

    ''' <summary>
    ''' The description of the region.
    ''' </summary>
    Public ReadOnly Property Description() As String
        Get
            Return _Description
        End Get
    End Property

    ''' <summary>
    ''' Constructs a new source region.
    ''' </summary>
    ''' <param name="start">The start location of the region.</param>
    ''' <param name="finish">The end location of the region.</param>
    ''' <param name="description">The description of the region.</param>
    Public Sub New(ByVal start As Location, ByVal finish As Location, ByVal description As String)
        _Start = start
        _Finish = finish
        _Description = description
    End Sub
End Class
