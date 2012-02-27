Imports System.IO
Imports System.Xml.Serialization

Public Class XMLHelper
    Public Shared Sub SerializeXml( _
    ByVal toSerialize As Object, _
    ByVal targetFileName As String)

        Dim sw As New StreamWriter(targetFileName)
        Dim ser As New XmlSerializer(toSerialize.GetType())
        ser.Serialize(sw, toSerialize)
        sw.Close()

    End Sub

    Public Shared Function DeSerializeXml( _
    ByVal toDeSerializeType As Type, _
    ByVal sourceFileName As String) As Object

        Dim sr As New StreamReader(sourceFileName)
        Dim ser As New XmlSerializer(toDeSerializeType)
        Dim o As Object = ser.Deserialize(sr)
        sr.Close()

        Return o
    End Function
End Class
