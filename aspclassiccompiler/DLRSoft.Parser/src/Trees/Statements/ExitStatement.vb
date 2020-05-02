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
''' A parse tree for an Exit statement.
''' </summary>
Public NotInheritable Class ExitStatement
    Inherits Statement

    Private ReadOnly _ExitType As BlockType
    Private ReadOnly _ExitArgumentLocation As Location

    ''' <summary>
    ''' The type of tree this statement exits.
    ''' </summary>
    Public ReadOnly Property ExitType() As BlockType
        Get
            Return _ExitType
        End Get
    End Property

    ''' <summary>
    ''' The location of the exit statement type.
    ''' </summary>
    Public ReadOnly Property ExitArgumentLocation() As Location
        Get
            Return _ExitArgumentLocation
        End Get
    End Property

    ''' <summary>
    ''' Constructs a parse tree for an Exit statement.
    ''' </summary>
    ''' <param name="exitType">The type of tree this statement exits.</param>
    ''' <param name="exitArgumentLocation">The location of the exit statement type.</param>
    ''' <param name="span">The location of the parse tree.</param>
    ''' <param name="comments">The comments for the parse tree.</param>
    Public Sub New(ByVal exitType As BlockType, ByVal exitArgumentLocation As Location, ByVal span As Span, ByVal comments As IList(Of Comment))
        MyBase.New(TreeType.ExitStatement, span, comments)

        Select Case exitType
            Case BlockType.Do, BlockType.For, BlockType.While, BlockType.Select, BlockType.Sub, BlockType.Function, _
                 BlockType.Property, BlockType.Try, BlockType.None
                ' OK

            Case Else
                Throw New ArgumentOutOfRangeException("exitType")
        End Select

        _ExitType = exitType
        _ExitArgumentLocation = exitArgumentLocation
    End Sub
End Class