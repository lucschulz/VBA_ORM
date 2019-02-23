Attribute VB_Name = "m_SchemaBuilder"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' REQUIRED REFERENCES
'''
''' Microsoft ActiveX Data Objects 6.1 Library
''' Microsoft for Visual Basic for Applications Extensibility
''' Microsoft ADO Ext. 6.0 for DDL and Security
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private CodeMod As VBIDE.codeModule
Private LineNum As Long

'This prefix can be anything that's valid for a module/class name but must be different from
'what is used for your regular modules otherwise they will get deleted. Each table module
'which represents a database table will be name using this prefix followed by the table's actual name.
Private Const TABLE_PREFIX = "tbl_"

'The name used for the module that holds all the database table objects. After running CreateTableClasses
'use this class to reference the databse table and column names.
Private Const SCHEMA_MODULE_NAME = "db_Schema"




''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' Set this property to return the connection string for the database
''' you want to map.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Property Get CString() As String
    Main.SetConnectionStringAndVersion
    CString = Main.GetConnectionString
End Property


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' Run this subroutine to retrieve the database schema and populate
''' and create a new class for each table that contains the table schema.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub CreateTableClasses()

    Dim cn As New ADODB.Connection
    Dim moduleNames As New Dictionary
    
    Dim Catalog As New ADOX.Catalog
    Dim VBComp As VBIDE.vbComponent
    
    Dim VBProj As VBIDE.vbProject
    Set VBProj = ActiveWorkbook.vbProject
    
    RemoveExistingTables
    
    cn.Open CString
    Set Catalog.ActiveConnection = cn
    
    Dim table As ADOX.table, column As ADOX.column
    For Each table In Catalog.Tables
        If table.Type = "TABLE" Or table.Type = "VIEW" Then
        
            Set VBComp = VBProj.VBComponents.Add(vbext_ct_ClassModule)
            VBComp.Name = TABLE_PREFIX & table.Name
            
            PopulateTableClass table.Name, table, VBComp
            
            moduleNames.Add table.Name, VBComp.Name
            
        End If
    Next
    
    cn.Close
    Set cn = Nothing
    
    CreateDatabaseSchemaClass moduleNames, VBProj
            
End Sub



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' Creates the interfacing (loosely speaking) class that links all the tables into a single class.
''' Serves as a VBA representation of the database.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CreateDatabaseSchemaClass(ByVal moduleNames As Dictionary, project As VBIDE.vbProject)
    
    Dim dbSchemaClass As String
    dbSchemaClass = SCHEMA_MODULE_NAME
        
    If ModuleExist(dbSchemaClass, project) Then
        DeleteModule project, project.VBComponents(dbSchemaClass)
    End If
    
    Dim component As VBIDE.vbComponent
    Set component = project.VBComponents.Add(vbext_ct_ClassModule)
    component.Name = dbSchemaClass
    
    
    Dim table As String
    Dim module As String

        
    Dim i As Variant
    For Each i In moduleNames.Keys
        table = i
        module = moduleNames(i)

        With component.codeModule
            LineNum = .CountOfLines + 1
                .InsertLines LineNum, "Public Property Get " & table & "() As " & module
            IncrementLineNumber
                .InsertLines LineNum, vbTab & "Set " & table & " = New " & module
            IncrementLineNumber
                .InsertLines LineNum, "End Property" & vbNewLine & vbNewLine
        End With

    Next i
    
    
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' Creates the database objects within the class as string properties.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PopulateTableClass(tblName As String, table As ADOX.table, component As VBIDE.vbComponent)
        
    With component.codeModule
                      
      LineNum = .CountOfLines + 1
          .InsertLines LineNum, "Public Property Get TableName() As String"
      IncrementLineNumber
          .InsertLines LineNum, vbTab & "TableName" & " = " & Chr(34) & table.Name & Chr(34)
      IncrementLineNumber
          .InsertLines LineNum, "End Property" & vbNewLine
      IncrementLineNumber
    
      
      Dim column As Variant
      
      For Each column In table.Columns
          Dim propName As String
          propName = column.Name

          IncrementLineNumber
              .InsertLines LineNum, "Public Property Get " & propName & "() As String"
              
          IncrementLineNumber
              .InsertLines LineNum, vbTab & propName & " = " & Chr(34) & column.Name & Chr(34)
              
          IncrementLineNumber
              .InsertLines LineNum, "End Property" & vbNewLine & vbNewLine
              
          IncrementLineNumber
        Next
    End With
    
End Sub



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' Removes all the modules whose name begin with "tbl_". These represent the database
''' tables.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub RemoveExistingTables()
    
    Dim VBProj As VBIDE.vbProject
    Set VBProj = ActiveWorkbook.vbProject
    
    Dim mdl As Variant
    Dim i As Long
    Dim collClassNames As New Collection
    
    For i = 1 To VBProj.VBComponents.Count
        Dim cName As String
        cName = VBProj.VBComponents(i).Name
        
        If cName Like TABLE_PREFIX & "*" Then
            collClassNames.Add cName
        End If
    Next i
    
    For Each mdl In collClassNames
        Dim component As VBIDE.vbComponent
        Set component = VBProj.VBComponents(mdl)
        DeleteModule VBProj, component
    Next
    
End Sub



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' Deletes the specified module.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub DeleteModule(vbProject As VBIDE.vbProject, component As VBIDE.vbComponent)
    vbProject.VBComponents.Remove component
End Sub



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' Returns TRUE if a the a module with the specified name exists (module or class).
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ModuleExist(sModuleName As String, VBProj As VBIDE.vbProject) As Boolean

    Dim mdl As Variant
    ModuleExist = False
    
    For Each mdl In VBProj.VBComponents
        If mdl.Name = sModuleName Then
            ModuleExist = True
            Exit For
        End If
    Next
    
End Function


Private Sub IncrementLineNumber()
    LineNum = LineNum + 1
End Sub
