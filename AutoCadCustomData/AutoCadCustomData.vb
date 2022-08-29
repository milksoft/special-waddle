
Imports Autodesk.AutoCAD.DatabaseServices

#Region "Информация для библиотеки DLL"
<Assembly: System.Reflection.AssemblyTitle("AutoCadCustomData")>
<Assembly: System.Reflection.AssemblyDescription("")>
<Assembly: System.Reflection.AssemblyCompany("")>
<Assembly: System.Reflection.AssemblyProduct("AutoCadCustomData")>
<Assembly: System.Reflection.AssemblyCopyright("Copyright © WizardSoft 2022")>
<Assembly: System.Reflection.AssemblyTrademark("")>
<Assembly: System.Reflection.AssemblyVersion("1.0.*")>
#End Region


Public Module AutoCadCustomData

    Private Const MyPropertyKey As String = "{90CE424B-0E50-454d-80D1-103A3D746B08}" 'Идентификатор данных

    ''' <summary>
    ''' Подготовить место для данных перед первым использованием
    ''' </summary>
    ''' <param name="obj">Идентификатор объекта</param>
    ''' <returns>Истина если операция прошла успешно, Лож если нет необходимости либо невозможно</returns>
    Public Function CreateFieldCustomData(ByVal obj As ObjectId) As Boolean

        Dim db As Database = obj.Database
        If db Is Nothing Then Return False

        Dim tr = db.TransactionManager.StartTransaction()
        Dim entity As DBObject = CType(tr.GetObject(obj, OpenMode.ForWrite), DBObject)
        Dim dictId As ObjectId = entity.ExtensionDictionary

        If dictId = ObjectId.Null Then
            entity.UpgradeOpen()
            entity.CreateExtensionDictionary()
            dictId = entity.ExtensionDictionary
        End If

        Dim xdict As DBDictionary = CType(tr.GetObject(dictId, OpenMode.ForRead), DBDictionary)
        Dim xrec As Xrecord

        If xdict.Contains(MyPropertyKey) Then
            xrec = CType(tr.GetObject(xdict(MyPropertyKey), OpenMode.ForRead), Xrecord)
            Return xrec IsNot Nothing
        Else
            xdict.UpgradeOpen()
            xrec = New Xrecord()
            xdict.SetAt(MyPropertyKey, xrec)
            tr.AddNewlyCreatedDBObject(xrec, True)
            Return True
        End If
    End Function

    ''' <summary>
    ''' Проверка есть ли данные у объекта
    ''' </summary>
    ''' <param name="obj">Идентификатор объекта</param>
    ''' <returns>Результат проверки</returns>
    Public Function IsExistsCustomData(ByVal obj As ObjectId) As Boolean
        Dim db As Database = obj.Database
        Dim tr = db.TransactionManager.StartOpenCloseTransaction()
        Dim entity As DBObject = CType(tr.GetObject(obj, OpenMode.ForWrite), DBObject)
        Dim dictId As ObjectId = entity.ExtensionDictionary
        If dictId = ObjectId.Null Then Return False
        Dim xdict As DBDictionary = CType(tr.GetObject(dictId, OpenMode.ForRead), DBDictionary)
        Return xdict.Contains(MyPropertyKey)
    End Function

    ''' <summary>
    ''' Сохранить данные
    ''' </summary>
    ''' <param name="obj">Идентификатор объекта</param>
    ''' <param name="data">строковые данные</param>
    ''' <returns>Истина если операция прошла успешно, Лож если требуется подготовка</returns>
    Public Function SetCustomData(ByVal obj As ObjectId, ByVal data As String) As Boolean
        Dim db As Database = obj.Database
        Dim tr = db.TransactionManager.StartTransaction()
        Dim entity As DBObject = CType(tr.GetObject(obj, OpenMode.ForWrite), DBObject)
        Dim dictId As ObjectId = entity.ExtensionDictionary
        If dictId = ObjectId.Null Then Return False
        Dim xdict As DBDictionary = CType(tr.GetObject(dictId, OpenMode.ForRead), DBDictionary)
        If Not xdict.Contains(MyPropertyKey) Then Return False
        Dim xrec = CType(tr.GetObject(xdict(MyPropertyKey), OpenMode.ForWrite), Xrecord)
        xrec.Data = New ResultBuffer(New TypedValue(CInt(DxfCode.Text), data))
        Return True
    End Function

    ''' <summary>
    ''' Получить данные
    ''' </summary>
    ''' <param name="obj">Идентификатор объекта</param>
    ''' <returns>Строковые данные </returns>
    Public Function GetCustomData(ByVal obj As ObjectId) As String
        Dim ret As String = String.Empty
        Dim db As Database = obj.Database

        Dim tr = db.TransactionManager.StartTransaction()
        Dim entity As DBObject = CType(tr.GetObject(obj, OpenMode.ForRead), DBObject)
        Dim dictId As ObjectId = entity.ExtensionDictionary
        If dictId = ObjectId.Null Then Return String.Empty
        Dim xdict = CType(tr.GetObject(dictId, OpenMode.ForRead), DBDictionary)
        If Not xdict.Contains(MyPropertyKey) Then Return String.Empty
        Dim xrec = CType(tr.GetObject(CType(xdict(MyPropertyKey), ObjectId), OpenMode.ForRead), Xrecord)
        If xrec Is Nothing OrElse xrec.Data Is Nothing Then Return String.Empty
        Dim data = xrec.Data.AsArray()
        If data.Length = 0 Then Return String.Empty
        ret = xrec.Data.AsArray()(0).Value.ToString()

        Return ret
    End Function

End Module