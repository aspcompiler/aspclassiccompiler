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
''' A parse tree for an Option declaration.
''' </summary>
Public NotInheritable Class OptionDeclaration
    Inherits Declaration

    Private ReadOnly _OptionType As OptionType
    Private ReadOnly _OptionTypeLocation As Location
    Private ReadOnly _OptionArgumentLocation As Location

    ''' <summary>
    ''' The type of Option statement.
    ''' </summary>
    Public ReadOnly Property OptionType() As OptionType
        Get
            Return _OptionType
        End Get
    End Property

    ''' <summary>
    ''' The location of the Option type (e.g. "Strict"), if any.
    ''' </summary>
    Public ReadOnly Property OptionTypeLocation() As Location
        Get
            Return _OptionTypeLocation
        End Get
    End Property

    ''' <summary>
    ''' The location of the Option argument (e.g. "On"), if any.
    ''' </summary>
    Public ReadOnly Property OptionArgumentLocation() As Location
        Get
            Return _OptionArgumentLocation
        End Get
    End Property

    ''' <summary>
    ''' Constructs a new parse tree for an Option declaration.
    ''' </summary>
    ''' <param name="optionType">The type of the Option declaration.</param>
    ''' <param name="optionTypeLocation">The location of the Option type, if any.</param>
    ''' <param name="optionArgumentLocation">The location of the Option argument, if any.</param>
    ''' <param name="span">The location of the parse tree.</param>
    ''' <param name="comments">The comments for the parse tree.</param>
    Public Sub New(ByVal optionType As OptionType, ByVal optionTypeLocation As Location, ByVal optionArgumentLocation As Location, ByVal span As Span, ByVal comments As IList(Of Comment))
        MyBase.New(TreeType.OptionDeclaration, span, comments)

        If optionType < optionType.SyntaxError OrElse optionType > optionType.CompareText Then
            Throw New ArgumentOutOfRangeException("optionType")
        End If

        _OptionType = optionType
        _OptionTypeLocation = optionTypeLocation
        _OptionArgumentLocation = optionArgumentLocation
    End Sub
End Class