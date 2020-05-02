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
''' A parse tree for an End statement of a block.
''' </summary>
Public NotInheritable Class EndBlockStatement
    Inherits Statement

    Private ReadOnly _EndType As BlockType
    Private ReadOnly _EndArgumentLocation As Location

    ''' <summary>
    ''' The type of block the statement ends.
    ''' </summary>
    Public ReadOnly Property EndType() As BlockType
        Get
            Return _EndType
        End Get
    End Property

    ''' <summary>
    ''' The location of the end block argument.
    ''' </summary>
    Public ReadOnly Property EndArgumentLocation() As Location
        Get
            Return _EndArgumentLocation
        End Get
    End Property

    ''' <summary>
    ''' Creates a new parse tree for an End block statement.
    ''' </summary>
    ''' <param name="endType">The type of the block the statement ends.</param>
    ''' <param name="endArgumentLocation">The location of the end block argument.</param>
    ''' <param name="span">The location of the parse tree.</param>
    ''' <param name="comments">The comments for the parse tree.</param>
    Public Sub New(ByVal endType As BlockType, ByVal endArgumentLocation As Location, ByVal span As Span, ByVal comments As IList(Of Comment))
        MyBase.New(TreeType.EndBlockStatement, span, comments)

        If endType < BlockType.While OrElse endType > BlockType.Namespace Then
            Throw New ArgumentOutOfRangeException("endType")
        End If

        _EndType = endType
        _EndArgumentLocation = endArgumentLocation
    End Sub
End Class