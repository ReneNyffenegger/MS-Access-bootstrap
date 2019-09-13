'
' Provide the functionality to create Access applications from
' the command line.
'
' The functions in this file should be called from a *.wsf file.
'

option explicit

dim fso
set fso = createObject("Scripting.FileSystemObject")

function createDB(accessFile) ' {

'   dim app
    set createDB = createOfficeApp("access", accessFile)

'   if fso.fileExists(accessFile) then
'      wscript.echo accessFile & " already exists. Deleting it."
'      fso.deleteFile(accessFile)
'   end if
' 
' 
'   set acc = createObject("Access.Application")
'   acc.newCurrentDatabase accessFile, 0 ' 0: acNewDatabaseFormatUserDefault

'   acc.visible     = true
'   acc.userControl = true ' http://stackoverflow.com/q/36282024/180275

'   set vb_editor = acc.vbe
'   set vb_proj   = vb_editor.activeVBProject
'   set vb_comps  = vb_proj.vbComponents

' '
' ' Add (type lib) reference to "Microsoft Visual Basic for Applications Extensibility 5.3"
' '
'   call addReference("{0002E157-0000-0000-C000-000000000046}", 5, 3)

end function ' }

function createOfficeApp(officeName, fileName) ' {
 '
 '  Note: when creating an Access database, this function
 '  returns an application object. When creating an excel
 '  Worksheet, it retuns an Excel Worksheet.
 '

    if fso.fileExists(fileName) then
       wscript.echo fileName & " already exists. Deleting it."
       fso.deleteFile(fileName)
    end if

    dim app
  
  
    if     officeName = "access" then

           set createOfficeApp = createObject("access.application")
           createOfficeApp.newCurrentDatabase fileName, 0 ' 0: acNewDatabaseFormatUserDefault

           set app = createOfficeApp

    elseIf officeName = "excel"  then

           set app             = createObject("excel.application")

         '
         ' createOfficeApp becomes a worksheet here, really...
         '
           set createOfficeApp = app.workBooks.add
           createOfficeApp.saveAs fileName, 52 ' 52 = xlOpenXMLWorkbookMacroEnabled

    elseIf officeName = "word"   then

           set app             = createObject("word.application")

           set createOfficeApp = app.documents.add
           createOfficeApp.saveAs fileName, 20 ' 20 = wdFormatFlatXMLMacroEnabled (Open XML file format with macros enabled saved as a single XML file.)

    end if

    app.visible     = true

    if officeName <> "word" then
  '
  ' Keep application opened after scripts terminates
  '   https://stackoverflow.com/q/36282024/180275
  '
  ' In Word, userControl is read only and set to true if
  ' the application was created with createObject(), getObject() or opened
  ' with open()
  '
      app.userControl = true
    end if

  '
  ' Add (type lib) reference to "Microsoft Visual Basic for Applications Extensibility 5.3"
  '
    call addReference(app, "{0002E157-0000-0000-C000-000000000046}", 5, 3)

end function ' }

sub insertModule(app, moduleFilePath, moduleName, moduleType) ' {
 '
 '  moduleType:
 '    1 = vbext_ct_StdModule
 '    2 = vbext_ct_ClassModule
 '
 '  Compare with https://renenyffenegger.ch/notes/development/languages/VBA/modules/Common/00_ModuleLoader
 '    

    if not fso.fileExists(moduleFilePath) then ' {
       wscript.echo moduleFilePath & " does not exist!"
       wscript.quit
    end if ' }


    dim vb_editor ' as vbe
    dim vb_proj   ' as VBProject
    dim vb_comps  ' as VBComponents
    dim mdl       ' as VBComponent

    set vb_editor = app.vbe
    set vb_proj   = vb_editor.activeVBProject
    set vb_comps  = vb_proj.vbComponents

  '
  ' Check if a module by the given name already exists.
  ' If so, remove it
  '
  ' 
  '

  '
  ' If no module with the name moduleName exists, by default
  ' vb_comps(moduleName) throws a 'VBAProject: Subscript out of range'
  ' error.
  ' We're going to let such an error escape by embedding the
  ' statement between the following two 'on error …' statements:
  '
    on error resume next
    set mdl = vb_comps(moduleName)
    on error goto 0

    if not isEmpty(mdl) then
'   if typeName(mdl) = "VBComponent" then
       vb_comps.remove mdl
    end if

    set mdl = vb_comps.add(moduleType)
   
    mdl.codeModule.addFromFile(ModuleFilePath)
   
    mdl.name = moduleName
   
    if app.name = "Microsoft Access" then
       app.doCmd.close 5, mdl.name, 1 ' 5=acModule, 1=acSaveYes
    end if


end sub ' }

sub addReference(app, guid, major, minor) ' {
  '
  ' guid identfies a type lib. Thus, the guid should be found in the
  ' Registry under HKEY_CLASSES_ROOT\TypeLib\
  '
  ' Note: guid probably needs the opening and closing curly paranthesis.
  '
    call app.VBE.activeVbProject.references.addFromGuid (guid, major, minor)
end sub ' }

function currentDir() ' {
     dim wshShell
     set wshShell = createObject("WScript.Shell")

     currentDir = wshShell.CurrentDirectory & "\"

end function ' }
