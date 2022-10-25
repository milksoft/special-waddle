Imports System.Runtime.InteropServices

Public Class AcadXRecordDataDinamicBinding
    Private Const MyPropertyKey As String = "{90CE424B-0E50-454d-80D1-103A3D746B08}"

    Private Const progIDstr As String = "AutoCAD.Application"

    Public Shared Function GetCustomXRecordDataDinamicBindingAcad(ByVal obj As Object) As String
        Dim db = obj.Database
        Dim dt1 As Object = Nothing, dt2 As Object = Nothing

        If obj.HasExtensionDictionary Then
            Dim dict = obj.GetExtensionDictionary()
            Dim XRecord = dict.GetObject(MyPropertyKey)
            XRecord.GetXRecordData(dt1, dt2)
            Dim arrdata = TryCast(dt2, Object())
            If arrdata.Length > 0 Then Return arrdata.ToString()
        End If

        Return String.Empty
    End Function

    Public Shared Sub SetCustomXRecordDataDinamicBindingAcad(ByVal obj As Object, ByVal value As String)
        Dim ret As String = String.Empty
        Dim db = obj.Database
        Dim dict = obj.GetExtensionDictionary()
        Dim XRecord = dict.AddXRecord(MyPropertyKey)
        XRecord.SetXRecordData(New Short() {1}, New Object() {value})
    End Sub

    Private Shared Sub Test()
        Dim acadapp As Object = Nothing
        Dim acaddoc As Object = Nothing
        acadapp = Marshal.GetActiveObject(progIDstr)
        If acadapp IsNot Nothing Then acaddoc = acadapp.ActiveDocument

        If acaddoc IsNot Nothing Then
            Console.WriteLine(acaddoc.Name)
        End If

        Dim ss = acaddoc.PickfirstSelectionSet()
        Dim item = ss.Item(0)
        Console.WriteLine(item.ObjectName)
        Dim val = $"Тесетq{New Random().[Next](100)}"
        Console.WriteLine(GetCustomXRecordDataDinamicBindingAcad(item))
        SetCustomXRecordDataDinamicBindingAcad(item, val)
        Console.WriteLine(GetCustomXRecordDataDinamicBindingAcad(item))
        Console.ReadLine()
    End Sub

End Class