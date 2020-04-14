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
''' A parse tree for a Case Else statement.
''' </summary>
Public NotInheritable Class CaseElseStatement
    Inherits Statement

    Private ReadOnly _ElseLocation As Location

    ''' <summary>
    ''' The location of the 'Else'.
    ''' </summary>
    Public ReadOnly Property ElseLocation() As Location
        Get
            Return _ElseLocation
        End Get
    End Property

    ''' <summary>
    ''' Constructs a new parse tree for a Case Else statement.
    ''' </summary>
    ''' <param name="elseLocation">The location of the 'Else'.</param>
    ''' <param name="span">The location of the parse tree.</param>
    ''' <param name="comments">The comments for the parse tree.</param>
    Public Sub New(ByVal elseLocation As Location, ByVal span As Span, ByVal comments As IList(Of Comment))
        MyBase.New(TreeType.CaseElseStatement, span, comments)

        _ElseLocation = elseLocation
    End Sub
End Class