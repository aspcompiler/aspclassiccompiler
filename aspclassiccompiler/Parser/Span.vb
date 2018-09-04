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
''' Stores the location of a span of text.
''' </summary>
''' <remarks>The end location is exclusive.</remarks>
Public Structure Span
    Private ReadOnly _Start As Location
    Private ReadOnly _Finish As Location

    ''' <summary>
    ''' The start location of the span.
    ''' </summary>
    Public ReadOnly Property Start() As Location
        Get
            Return _Start
        End Get
    End Property

    ''' <summary>
    ''' The end location of the span.
    ''' </summary>
    Public ReadOnly Property Finish() As Location
        Get
            Return _Finish
        End Get
    End Property

    ''' <summary>
    ''' Whether the locations in the span are valid.
    ''' </summary>
    Public ReadOnly Property IsValid() As Boolean
        Get
            Return Start.IsValid AndAlso Finish.IsValid
        End Get
    End Property

    ''' <summary>
    ''' Constructs a new span with a specific start and end location.
    ''' </summary>
    ''' <param name="start">The beginning of the span.</param>
    ''' <param name="finish">The end of the span.</param>
    Public Sub New(ByVal start As Location, ByVal finish As Location)
        _Start = Start
        _Finish = Finish
    End Sub

    ''' <summary>
    ''' Compares two specified Span values to see if they are equal.
    ''' </summary>
    ''' <param name="left">One span to compare.</param>
    ''' <param name="right">The other span to compare.</param>
    ''' <returns>True if the spans are the same, False otherwise.</returns>
    Public Shared Operator =(ByVal left As Span, ByVal right As Span) As Boolean
        Return left.Start.Index = right.Start.Index AndAlso left.Finish.Index = right.Finish.Index
    End Operator

    ''' <summary>
    ''' Compares two specified Span values to see if they are not equal.
    ''' </summary>
    ''' <param name="left">One span to compare.</param>
    ''' <param name="right">The other span to compare.</param>
    ''' <returns>True if the spans are not the same, False otherwise.</returns>
    Public Shared Operator <>(ByVal left As Span, ByVal right As Span) As Boolean
        Return left.Start.Index <> right.Start.Index OrElse left.Finish.Index <> right.Finish.Index
    End Operator

    Public Overrides Function ToString() As String
        Return Start.ToString() & " - " & Finish.ToString()
    End Function

    Public Overrides Function Equals(ByVal obj As Object) As Boolean
        If TypeOf obj Is Span Then
            Return Me = DirectCast(obj, Span)
        Else
            Return False
        End If
    End Function

    Public Overrides Function GetHashCode() As Integer
        Return CInt((Start.Index Xor Finish.Index) And &HFFFFFFFFL)
    End Function
End Structure