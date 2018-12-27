' OLD DOCUMENTATION, please remove me!!!!
'
' cscript accessBootstrap.vbs \path\to\accessFileToCreate.accdb \path\of\00_ModuleLoader.bas \path\to\initModule.bas
'
' Creates the access file specified with the first parameter (after
' deleting it if it exists)
'
' Then loads the module 00_ModuleLoader.bas (See https://renenyffenegger.ch/notes/development/languages/VBA/modules/Common/00_ModuleLoader) whose
' directory is specified with the second parameter.
' created access files,
'
' Then loads the initModule.bas module which must contain a createApp sub.
'


option explicit

dim accessFile
dim acc       ' as access.application

dim fso
set fso = createObject("Scripting.FileSystemObject")
  
dim vb_editor ' as vbe

dim vb_proj   ' as VBProject
dim vb_comps  ' as VBComponents

' dim args
' dim scriptFile

' dim initFunc
' initFunc = "createApp"

' set args = wscript.arguments

' if args.count < 3 then
'   wscript.echo args.count & " were given, but at least 3 are required:"
'   wscript.echo "  Name of access db to be created"
'   wscript.echo "  Path (without filename) to 00_ModuleLoader.bas"
'   wscript.echo "  Path (with filename) to module that creates the app."
'   wscript.quit
'end if
  
'accessFile = args(0)


' wscript.echo "accessFile = " & accessFile


' call insertModule(args(1) & "\00_ModuleLoader.bas", "00_ModuleLoader")
' call insertModule(args(2)                         , "createAppModule")
    
' wscript.echo "Calling " & initFunc  
' acc.run(initFunc)

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
    '    Add reference to "Microsoft Visual Basic for Applications Extensibility 5.3"
    '
    call addReference("{0002E157-0000-0000-C000-000000000046}", 5, 3)


end sub ' }

sub insertModule(moduleFilePath, moduleName, moduleType) ' {
 '
 '  moduleType:
 '    1 = vbext_ct_StdModule
 '    2 = vbext_ct_ClassModule
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
  ' Note: guid probably needs the opening and closing curly paranthesis.
  '
    call acc.VBE.activeVbProject.references.addFromGuid (guid, major, minor)
end sub ' }