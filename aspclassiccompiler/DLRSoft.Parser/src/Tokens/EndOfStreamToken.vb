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
''' A token representing the end of the file.
''' </summary>
Public NotInheritable Class EndOfStreamToken
    Inherits Token

    ''' <summary>
    ''' Creates a new end-of-stream token.
    ''' </summary>
    ''' <param name="span">The location of the end of the stream.</param>
    Public Sub New(ByVal span As Span)
        MyBase.New(TokenType.EndOfStream, Span)
    End Sub
End Class

