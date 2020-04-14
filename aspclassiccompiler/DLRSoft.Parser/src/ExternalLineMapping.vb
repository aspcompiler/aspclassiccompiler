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
''' A line mapping from a source span to an external file and line.
''' </summary>
Public NotInheritable Class ExternalLineMapping
    Private ReadOnly _Start, _Finish As Location
    Private ReadOnly _File As String
    Private ReadOnly _Line As Long

    ''' <summary>
    ''' The start location of the mapping in the source.
    ''' </summary>
    Public ReadOnly Property Start() As Location
        Get
            Return _Start
        End Get
    End Property

    ''' <summary>
    ''' The end location of the mapping in the source.
    ''' </summary>
    Public ReadOnly Property Finish() As Location
        Get
            Return _Finish
        End Get
    End Property

    ''' <summary>
    ''' The external file the source maps to.
    ''' </summary>
    Public ReadOnly Property File() As String
        Get
            Return _File
        End Get
    End Property

    ''' <summary>
    ''' The external line number the source maps to.
    ''' </summary>
    Public ReadOnly Property Line() As Long
        Get
            Return _Line
        End Get
    End Property

    ''' <summary>
    ''' Constructs a new external line mapping.
    ''' </summary>
    ''' <param name="start">The start location in the source.</param>
    ''' <param name="finish">The end location in the source.</param>
    ''' <param name="file">The name of the external file.</param>
    ''' <param name="line">The line number in the external file.</param>
    Public Sub New(ByVal start As Location, ByVal finish As Location, ByVal file As String, ByVal line As Long)
        _Start = start
        _Finish = finish
        _File = file
        _Line = line
    End Sub
End Class