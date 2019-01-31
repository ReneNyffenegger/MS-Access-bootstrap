'
' Provide the functionality to create Access applications from
' the command line.
'
' The functions in this file should be called from a *.wsf file.
'

option explicit

dim accessFile
dim acc       ' as access.application

dim fso
set fso = createObject("Scripting.FileSystemObject")
  
dim vb_editor ' as vbe

dim vb_proj   ' as VBProject
dim vb_comps  ' as VBComponents

sub createDB(accessFile) ' {


    if fso.fileExists(accessFile) then
       wscript.echo accessFile & " already exists. Deleting it."
       fso.deleteFile(accessFile)
    end if
  
  
    set acc = createObject("Access.Application")
    acc.newCurrentDatabase accessFile, 0 ' 0: acNewDatabaseFormatUserDefault

    acc.visible     = true
    acc.userControl = true ' http://stackoverflow.com/q/36282024/180275

    set vb_editor = acc.vbe
    set vb_proj   = vb_editor.activeVBProject
    set vb_comps  = vb_proj.vbComponents

  '
  ' Add (type lib) reference to "Microsoft Visual Basic for Applications Extensibility 5.3"
  '
    call addReference("{0002E157-0000-0000-C000-000000000046}", 5, 3)


end sub ' }

sub insertModule(moduleFilePath, moduleName, moduleType) ' {
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
   
    acc.doCmd.close 5, mdl.name, 1 ' 5=acModule, 1=acSaveYes

end sub ' }

sub addReference(guid, major, minor) ' {
  '
  ' guid identfies a type lib. Thus, the guid should be found in the
  ' Registry under HKEY_CLASSES_ROOT\TypeLib\
  '
  ' Note: guid probably needs the opening and closing curly paranthesis.
  '
    call acc.VBE.activeVbProject.references.addFromGuid (guid, major, minor)
end sub ' }
