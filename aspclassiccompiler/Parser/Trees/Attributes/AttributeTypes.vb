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
''' The type of an attribute usage.
''' </summary>
<Flags()> _
Public Enum AttributeTypes
    ''' <summary>Regular application.</summary>
    Regular = &H1

    ''' <summary>Applied to the netmodule.</summary>
    [Module] = &H2

    ''' <summary>Applied to the assembly.</summary>
    Assembly = &H4
End Enum