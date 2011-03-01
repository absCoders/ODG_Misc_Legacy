Namespace Extensions


    Public Module ExtensionMethods

        <System.Runtime.CompilerServices.Extension()> _
                Public Function ToByteArray(ByVal aString As String) As Byte()
            Dim encoding As New System.Text.ASCIIEncoding()
            Return encoding.GetBytes(aString)
        End Function

        <System.Runtime.CompilerServices.Extension()> _
       Public Function ToStringX(ByVal bytes As Byte()) As String
            Dim enc As System.Text.ASCIIEncoding = New System.Text.ASCIIEncoding()
            Return enc.GetString(bytes)
        End Function


        <System.Runtime.CompilerServices.Extension()> _
        Public Function Append(ByVal byte1 As Byte(), ByVal byte2 As Byte()) As Byte()
            Dim len = byte1.Length + byte2.Length
            Dim newBytes As Byte() = New Byte(len - 1) {}

            Buffer.BlockCopy(byte1, 0, newBytes, 0, byte1.Length)
            Buffer.BlockCopy(byte2, 0, newBytes, byte1.Length, byte2.Length)

            Return newBytes
        End Function

        <System.Runtime.CompilerServices.Extension()> _
        Public Function ToHexBytes(ByVal hex As String) As Byte()
            Dim NumberChars As Integer = hex.Length / 2
            Dim bytes As Byte() = New Byte(NumberChars - 1) {}
            For i As Integer = 0 To NumberChars Step 2
                bytes(i / 2) = Convert.ToByte(hex.Substring(i, 2), 16)
            Next
            Return bytes
        End Function

        <System.Runtime.CompilerServices.Extension()> _
        Public Function ToHexString(ByVal num As Integer) As String
            Return num.ToString("X").PadLeft(4, "0")
        End Function

        <System.Runtime.CompilerServices.Extension()> _
        Public Function ToCommand(ByVal bytes As Byte()) As Byte()
            Return BitConverter.GetBytes(bytes.Length).Append(bytes)
        End Function

        <System.Runtime.CompilerServices.Extension()> _
        Public Function ToCommand(ByVal str As String) As Byte()
            Return str.ToByteArray.ToCommand()
        End Function

    End Module

End Namespace