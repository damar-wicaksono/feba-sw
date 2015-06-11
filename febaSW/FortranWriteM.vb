Module FortranWriteM

    ''' <summary>
    ''' Pad a string with spaces from the left
    ''' Think of it as fixed-column, right justified
    ''' </summary>
    ''' <param name="strBuffer">The input string</param>
    ''' <param name="l">the number of columns</param>
    ''' <returns>String with extra spaces</returns>
    ''' <remarks>Converted from VBA to VB.NET</remarks>
    ''' <author>
    ''' CI41, LRS/PSI, 2014
    ''' WD41, LRS/EPFL/PSI, 2015
    ''' </author>
    Function leftPad(ByVal strBuffer As String,
                     ByVal l As Integer) As String

        Dim intCount As Integer
        intCount = l - strBuffer.Length
        If intCount > 0 Then
            leftPad = String.Format("{0}{1}", Space(intCount), strBuffer)
        Else
            leftPad = strBuffer
        End If
        Return leftPad

    End Function
End Module
