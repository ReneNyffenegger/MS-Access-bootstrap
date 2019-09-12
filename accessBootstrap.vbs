'
' Provide the functionality to create Access applications from
' the command line.
'
' The functions in this file should be called from a *.wsf file.
'

option explicit

' dim accessFile
' dim app       ' as access.application
' dim xls

dim fso
set fso = createObject("Scripting.FileSystemObject")
  
dim vb_editor ' as vbe

dim vb_proj   ' as VBProject
dim vb_comps  ' as VBComponents

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

    if fso.fileExists(fileName) then
       wscript.echo fileName & " already exists. Deleting it."
       fso.deleteFile(fileName)
    end if
  
  
    if     officeName = "access" then

           set createOfficeApp = createObject("access.application")
           createOfficeApp.newCurrentDatabase fileName, 0 ' 0: acNewDatabaseFormatUserDefault

    elseIf officeName = "excel"  then

           dim xls
           set createOfficeApp = createObject("excel.application")
           set xls = app.workBooks.add
           xls.saveAs fileName, 52 ' 52 = xlOpenXMLWorkbookMacroEnabled

    end if


    createOfficeApp.visible     = true
    createOfficeApp.userControl = true ' https://stackoverflow.com/q/36282024/180275

    set vb_editor = createOfficeApp.vbe
    set vb_proj   = vb_editor.activeVBProject
    set vb_comps  = vb_proj.vbComponents

  '
  ' Add (type lib) reference to "Microsoft Visual Basic for Applications Extensibility 5.3"
  '
    call addReference(createOfficeApp, "{0002E157-0000-0000-C000-000000000046}", 5, 3)

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

    dim mdl ' as VBComponent
    set mdl = vb_comps.add(1) ' 1 = vbext_ct_StdModule
   
    wscript.echo("adding scriptFile " & ModuleFilePath)
    mdl.codeModule.addFromFile (ModuleFilePath)
   
    mdl.name = moduleName
   
    if app.name = "MicrosoftAccess" then
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
