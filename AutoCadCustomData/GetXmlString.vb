Imports System.Xml

Public Class GetXmlString

    Public Shared Sub Test() ' пример использования

        Dim obj = New With {Key .Name = "name1", 'тестовые данные загоняем
                        Key .Id = "1",
                        Key .Type = "typesf",
                        Key .Materials = {
                                            New With {Key .MaterialName = "m1", Key .Parametres =
                                                                                    {New With {Key .Name = "p1", Key .Value = "p1", Key .MUnit = "p1"}}},
                                            New With {Key .MaterialName = "m1", Key .Parametres =
                                                                                    {New With {Key .Name = "p2", Key .Value = "p2", Key .MUnit = "p2"}}}
                                         },
                                        Key .Parametres = {New With {Key .Name = "p3", Key .Value = "p4", Key .MUnit = "p5"
                            }}
                        }

        Dim s = GetXmlStringToObject(obj)
        Console.WriteLine(s)
        Console.ReadLine()
    End Sub

    Public Shared Function GetXmlStringToObject(objectdata As Object) As String ' Функция на вход принимает данные возвращает строку xml
        Dim doc = New XmlDocument()
        Dim node = GetInstance(doc, objectdata)
        Return node.OuterXml
    End Function

    Private Shared Function GetInstance(doc As XmlDocument, data As Object) As XmlNode

        Dim inst = doc.AppendChild(doc.CreateElement("Instance"))

        inst.Attributes.Append(doc.CreateAttribute("UniqueID")).Value = data.Id ' заполняем какие-то данные для примера
        inst.Attributes.Append(doc.CreateAttribute("Name")).Value = data.Name
        inst.Attributes.Append(doc.CreateAttribute("Type")).Value = data.Type

        inst.AppendChild(GetParams(doc, data.Parametres))

        inst.AppendChild(GetMaterials(doc, data.Materials))

        Return inst
    End Function

    Private Shared Function GetMaterials(doc As XmlDocument, data As Object()) As XmlNode
        Dim matsElement = doc.CreateElement("Materials")

        For Each item In data
            Dim inst = matsElement.AppendChild(doc.CreateElement("Material"))
            inst.Attributes.Append(doc.CreateAttribute("Name")).Value = item.MaterialName
            inst.AppendChild(GetParams(doc, item.Parametres))
        Next

        Return matsElement
    End Function

    Private Shared Function GetParams(doc As XmlDocument, data As Object()) As XmlNode
        Dim propertysElement = doc.CreateElement("Dimensions")

        For Each item In data
            propertysElement.AppendChild(GetProperty(doc, item))
        Next

        Return propertysElement
    End Function

    Private Shared Function GetProperty(ByRef doc As XmlDocument, item As Object) As XmlNode
        Dim xmlElement = doc.CreateElement("Property")
        xmlElement.Attributes.Append(doc.CreateAttribute("Name")).Value = item.Name
        xmlElement.Attributes.Append(doc.CreateAttribute("Value")).Value = item.Value
        If Not String.IsNullOrEmpty(item.MUnit) Then xmlElement.Attributes.Append(doc.CreateAttribute("Unit")).Value = item.MUnit

        Return xmlElement
    End Function

End Class
