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
''' An external checksum for a file.
''' </summary>
Public NotInheritable Class ExternalChecksum
    Private ReadOnly _Filename As String
    Private ReadOnly _Guid As String
    Private ReadOnly _Checksum As String

    ''' <summary>
    ''' The filename that the checksum is for.
    ''' </summary>
    Public ReadOnly Property Filename() As String
        Get
            Return _Filename
        End Get
    End Property

    ''' <summary>
    ''' The guid of the file.
    ''' </summary>
    Public ReadOnly Property Guid() As String
        Get
            Return _Guid
        End Get
    End Property

    ''' <summary>
    ''' The checksum for the file.
    ''' </summary>
    Public ReadOnly Property Checksum() As String
        Get
            Return _Checksum
        End Get
    End Property

    ''' <summary>
    ''' Constructs a new external checksum.
    ''' </summary>
    ''' <param name="filename">The filename that the checksum is for.</param>
    ''' <param name="guid">The guid of the file.</param>
    ''' <param name="checksum">The checksum for the file.</param>
    Public Sub New(ByVal filename As String, ByVal guid As String, ByVal checksum As String)
        _Filename = filename
        _Guid = guid
        _Checksum = checksum
    End Sub
End Class