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
''' A line terminator.
''' </summary>
Public NotInheritable Class LineTerminatorToken
    Inherits Token

    ''' <summary>
    ''' Create a new line terminator token.
    ''' </summary>
    ''' <param name="span">The location of the line terminator.</param>
    Public Sub New(ByVal span As Span)
        MyBase.New(TokenType.LineTerminator, span)
    End Sub
End Class