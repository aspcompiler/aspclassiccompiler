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
''' A parse tree for an Continue statement.
''' </summary>
Public NotInheritable Class ContinueStatement
    Inherits Statement

    Private ReadOnly _ContinueType As BlockType
    Private ReadOnly _ContinueArgumentLocation As Location

    ''' <summary>
    ''' The type of tree this statement continues.
    ''' </summary>
    Public ReadOnly Property ContinueType() As BlockType
        Get
            Return _ContinueType
        End Get
    End Property

    ''' <summary>
    ''' The location of the Continue statement type.
    ''' </summary>
    Public ReadOnly Property ContinueArgumentLocation() As Location
        Get
            Return _ContinueArgumentLocation
        End Get
    End Property

    ''' <summary>
    ''' Constructs a parse tree for an Continue statement.
    ''' </summary>
    ''' <param name="continueType">The type of tree this statement continues.</param>
    ''' <param name="continueArgumentLocation">The location of the Continue statement type.</param>
    ''' <param name="span">The location of the parse tree.</param>
    ''' <param name="comments">The comments for the parse tree.</param>
    Public Sub New(ByVal continueType As BlockType, ByVal continueArgumentLocation As Location, ByVal span As Span, ByVal comments As IList(Of Comment))
        MyBase.New(TreeType.ContinueStatement, span, comments)

        Select Case continueType
            Case BlockType.Do, BlockType.For, BlockType.While, BlockType.None
                ' OK

            Case Else
                Throw New ArgumentOutOfRangeException("continueType")
        End Select

        _ContinueType = continueType
        _ContinueArgumentLocation = continueArgumentLocation
    End Sub
End Class