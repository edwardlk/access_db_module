' Code example
' to use:
'        - create code with correct fields
'        - open Access and create new database
'        - go to 'Create > Module'
'        - paste code into module window
'        - save & run
'        - to fix any mistakes, just fix errors in module window, delete tables, then rerun module

Public Sub daoCreateTables()
 ' Sketch of how to use this code to create some tables 
 ' Tables have 255 max columns
 Dim db As DAO.Database
 Dim tdf As DAO.TableDef
 Dim prp As DAO.Property
 
 Set db = CurrentDb
 On Error Resume Next
 
 ' Create the table definition
 Set tdf = db.CreateTableDef("table_name")
 
 ' Create the table primary key
 Dim fldID As DAO.Field
 Set fldID = tdf.CreateField("autoID", dbLong)
 fldID.Attributes = dbAutoIncrField
 fldID.Required = True
 tdf.Fields.Append fldID
 
 ' Add the table to the database
 db.TableDefs.Append tdf
 Set tdf = db.TableDefs("table_name")
 
 ' Add other fields
 ' Text field
 Dim fldA01 As DAO.Field
 Set fldA01 = tdf.CreateField("field_1", dbText)
 tdf.Fields.Append fldA01

 ' Integer field
 Dim fldA02 As DAO.Field
 Set fldA02 = tdf.CreateField("field_2", dbInteger)
 tdf.Fields.Append fldA02

 ' Float field
 Dim fldA03 As DAO.Field
 Set fldA03 = tdf.CreateField("field_3", dbDouble)
 tdf.Fields.Append fldA03

 ' Date (or time) field
 Dim fldI083 As DAO.Field
 Set fldI083 = tdf.CreateField("Date_dental", dbDate)
 tdf.Fields.Append fldI083

 ' Yes/No field with checkbox
 Dim fldA04 As DAO.Field
 Set fldA04 = tdf.CreateField("field_4", dbBoolean)
 tdf.Fields.Append fldA04
 Call SetPropertyDAO(fldA04, "DisplayControl", dbInteger, CInt(acCheckBox))
 
 ' Combo Box field
 Dim fldA05 As DAO.Field
 Set fldA05 = tdf.CreateField("field_5", dbText)
 tdf.Fields.Append fldA05
 Call setComboProperties(fldA05, "Option 1;Option 2; Option 3")

 ' Combo Box multi-select field
 Dim fldA06 As DAO.Field
 Set fldA06 = tdf.CreateField("field_6", dbText)
 tdf.Fields.Append fldA06
 Call setComboMultiProperties(fldA06, "Option 4; Option 5; Option 6")

 ' ...

 ' Create the table 2 definition
 Set tdf = db.CreateTableDef("table_2_name")
 
 ' Create the field definitions
 Dim fldID2 As DAO.Field
 Set fldID2 = tdf.CreateField("autoID", dbLong)
 fldID2.Attributes = dbAutoIncrField
 fldID2.Required = True
 tdf.Fields.Append fldID2
 
 ' Add the table to the database
 db.TableDefs.Append tdf
 Set tdf = db.TableDefs("table_2_name")
 
 ' Add the table 2 fields

 ' ...

 ' Refresh the tables and database
 db.TableDefs.Refresh
 Application.RefreshDatabaseWindow

 Debug.Print "Done"

End Sub

Function setComboProperties(obj As Object, strList As String)
 ' Purpose:   Set the properties of the single-select combo box.
 ' Arguments: obj = the object whose property should be set.
 '            strList = the list of potential answers.

 With obj
  .Properties.Append .CreateProperty("DisplayControl", dbInteger, AcControlType.acComboBox)
  .Properties.Append .CreateProperty("RowSourceType", dbText, "Value List")
  .Properties.Append .CreateProperty("RowSource", dbText, strList)
  .Properties.Append .CreateProperty("LimitToList", dbBoolean, True)
 End With

End Function

Function setComboMultiProperties(obj As Object, strList As String)
 ' Purpose:   Set the properties of the multi-select combo box.
 ' Arguments: obj = the object whose property should be set.
 '            strList = the list of potential answers.
    
 With obj
  .Properties.Append .CreateProperty("DisplayControl", dbInteger, AcControlType.acComboBox)
  .Properties.Append .CreateProperty("RowSourceType", dbText, "Value List")
  .Properties.Append .CreateProperty("RowSource", dbText, strList)
  .Properties.Append .CreateProperty("LimitToList", dbBoolean, True)
  .Properties.Append .CreateProperty("AllowMultipleValues", dbBoolean, True)
 End With

End Function

' Functions from http://allenbrowne.com/func-DAO.html

Function SetPropertyDAO(obj As Object, strPropertyName As String, intType As Integer, _
    varValue As Variant, Optional strErrMsg As String) As Boolean
 On Error GoTo ErrHandler
 ' Purpose:   Set a property for an object, creating if necessary.
 ' Arguments: obj = the object whose property should be set.
 '            strPropertyName = the name of the property to set.
 '            intType = the type of property (needed for creating)
 '            varValue = the value to set this property to.
 '            strErrMsg = string to append any error message to.
    
 If HasProperty(obj, strPropertyName) Then
  obj.Properties(strPropertyName) = varValue
 Else
  obj.Properties.Append obj.CreateProperty(strPropertyName, intType, varValue)
 End If
 SetPropertyDAO = True
 
ExitHandler:
 Exit Function

ErrHandler:
 strErrMsg = strErrMsg & obj.Name & "." & strPropertyName & " not set to " & varValue & _
  ". Error " & Err.Number & " - " & Err.Description & vbCrLf
 Resume ExitHandler
End Function

Public Function HasProperty(obj As Object, strPropName As String) As Boolean
 ' Purpose:   Return true if the object has the property.
 Dim varDummy As Variant
 
 On Error Resume Next
 varDummy = obj.Properties(strPropName)
 HasProperty = (Err.Number = 0)
End Function
