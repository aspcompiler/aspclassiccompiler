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
''' Stores source code line and column information.
''' </summary>
Public Structure Location
    Private ReadOnly _Index As Integer
    Private ReadOnly _Line As Integer
    Private ReadOnly _Column As Integer

    ''' <summary>
    ''' The index in the stream (0-based).
    ''' </summary>
    Public ReadOnly Property Index() As Integer
        Get
            Return _Index
        End Get
    End Property

    ''' <summary>
    ''' The physical line number (1-based).
    ''' </summary>
    Public ReadOnly Property Line() As Integer
        Get
            Return _Line
        End Get
    End Property

    ''' <summary>
    ''' The physical column number (1-based).
    ''' </summary>
    Public ReadOnly Property Column() As Integer
        Get
            Return _Column
        End Get
    End Property

    ''' <summary>
    ''' Whether the location is a valid location.
    ''' </summary>
    Public ReadOnly Property IsValid() As Boolean
        Get
            Return Line <> 0 AndAlso Column <> 0
        End Get
    End Property

    ''' <summary>
    ''' Constructs a new Location for a particular source location.
    ''' </summary>
    ''' <param name="index">The index in the stream (0-based).</param>
    ''' <param name="line">The physical line number (1-based).</param>
    ''' <param name="column">The physical column number (1-based).</param>
    Public Sub New(ByVal index As Integer, ByVal line As Integer, ByVal column As Integer)
        _Index = index
        _Line = line
        _Column = column
    End Sub

    ''' <summary>
    ''' Compares two specified Location values to see if they are equal.
    ''' </summary>
    ''' <param name="left">One location to compare.</param>
    ''' <param name="right">The other location to compare.</param>
    ''' <returns>True if the locations are the same, False otherwise.</returns>
    Public Shared Operator =(ByVal left As Location, ByVal right As Location) As Boolean
        Return left.Index = right.Index
    End Operator

    ''' <summary>
    ''' Compares two specified Location values to see if they are not equal.
    ''' </summary>
    ''' <param name="left">One location to compare.</param>
    ''' <param name="right">The other location to compare.</param>
    ''' <returns>True if the locations are not the same, False otherwise.</returns>
    Public Shared Operator <>(ByVal left As Location, ByVal right As Location) As Boolean
        Return left.Index <> right.Index
    End Operator

    ''' <summary>
    ''' Compares two specified Location values to see if one is before the other.
    ''' </summary>
    ''' <param name="left">One location to compare.</param>
    ''' <param name="right">The other location to compare.</param>
    ''' <returns>True if the first location is before the other location, False otherwise.</returns>
    Public Shared Operator <(ByVal left As Location, ByVal right As Location) As Boolean
        Return left.Index < right.Index
    End Operator

    ''' <summary>
    ''' Compares two specified Location values to see if one is after the other.
    ''' </summary>
    ''' <param name="left">One location to compare.</param>
    ''' <param name="right">The other location to compare.</param>
    ''' <returns>True if the first location is after the other location, False otherwise.</returns>
    Public Shared Operator >(ByVal left As Location, ByVal right As Location) As Boolean
        Return left.Index > right.Index
    End Operator

    ''' <summary>
    ''' Compares two specified Location values to see if one is before or the same as the other.
    ''' </summary>
    ''' <param name="left">One location to compare.</param>
    ''' <param name="right">The other location to compare.</param>
    ''' <returns>True if the first location is before or the same as the other location, False otherwise.</returns>
    Public Shared Operator <=(ByVal left As Location, ByVal right As Location) As Boolean
        Return left.Index <= right.Index
    End Operator

    ''' <summary>
    ''' Compares two specified Location values to see if one is after or the same as the other.
    ''' </summary>
    ''' <param name="left">One location to compare.</param>
    ''' <param name="right">The other location to compare.</param>
    ''' <returns>True if the first location is after or the same as the other location, False otherwise.</returns>
    Public Shared Operator >=(ByVal left As Location, ByVal right As Location) As Boolean
        Return left.Index >= right.Index
    End Operator

    ''' <summary>
    ''' Compares two specified Location values.
    ''' </summary>
    ''' <param name="left">One location to compare.</param>
    ''' <param name="right">The other location to compare.</param>
    ''' <returns>0 if the locations are equal, -1 if the left one is less than the right one, 1 otherwise.</returns>
    Public Shared Function Compare(ByVal left As Location, ByVal right As Location) As Integer
        If left = right Then
            Return 0
        ElseIf left < right Then
            Return -1
        Else
            Return 1
        End If
    End Function

    Public Overrides Function ToString() As String
        Return "(" & Me.Column & "," & Me.Line & ")"
    End Function

    Public Overrides Function Equals(ByVal obj As Object) As Boolean
        If TypeOf obj Is Location Then
            Return Me = DirectCast(obj, Location)
        Else
            Return False
        End If
    End Function

    Public Overrides Function GetHashCode() As Integer
        ' Mask off the upper 32 bits of the index and use that as
        ' the hash code.
        Return CInt(Index And &HFFFFFFFFL)
    End Function
End Structure