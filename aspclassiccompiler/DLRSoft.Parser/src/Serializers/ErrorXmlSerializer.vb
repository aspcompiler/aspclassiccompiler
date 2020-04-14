'
' Visual Basic .NET Parser
'
' Copyright (C) 2005, Microsoft Corporation. All rights reserved.
'
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER 
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF 
' MERCHANTIBILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
'

Imports System.Xml

Public Class ErrorXmlSerializer
    Private ReadOnly Writer As XmlWriter

    Public Sub New(ByVal Writer As XmlWriter)
        Me.Writer = Writer
    End Sub

    Private Sub Serialize(ByVal Span As Span)
        Writer.WriteAttributeString("startLine", CStr(Span.Start.Line))
        Writer.WriteAttributeString("startCol", CStr(Span.Start.Column))
        Writer.WriteAttributeString("endLine", CStr(Span.Finish.Line))
        Writer.WriteAttributeString("endCol", CStr(Span.Finish.Column))
    End Sub

    Public Sub Serialize(ByVal SyntaxError As SyntaxError)
        Writer.WriteStartElement(SyntaxError.Type.ToString())
        Serialize(SyntaxError.Span)
        Writer.WriteString(SyntaxError.ToString())
        Writer.WriteEndElement()
    End Sub

    Public Sub Serialize(ByVal SyntaxErrors As List(Of SyntaxError))
        For Each SyntaxError As SyntaxError In SyntaxErrors
            Serialize(SyntaxError)
        Next
    End Sub

End Class
