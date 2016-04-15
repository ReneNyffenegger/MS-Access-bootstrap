'
' cscript accessBootstrap.vbs \path\to\accessFileToCreate.accdb \path\to\init.bas 
'
' Creates the access file specified with the first parameter (after
' deleting it if it exists)
'
' Then loads the module specified with the second parameter into the
' created access files, names it "_init" and runs the sub init
' that must exist in that module.
'


option explicit

dim accessFile
dim acc       ' as access.application
dim fso
  
dim vb_editor ' as vbe
dim mdl       ' as VBComponent

dim vb_proj   ' as VBProject
dim vb_comps  ' as VBComponents

dim args
dim scriptFile

set args = wscript.arguments

if args.count < 2 then
   wscript.echo "required: access file name and script file name"
   wscript.quit
end if
  
accessFile = args(0)
scriptFile = args(1)
  
set fso = createObject("Scripting.FileSystemObject")

if fso.fileExists(accessFile) then
   fso.deleteFile(accessFile)
end if
  
  
set acc = createObject("Access.Application")
  
acc.newCurrentDatabase accessFile, 0 ' 0: acNewDatabaseFormatUserDefault

acc.visible     = true
acc.userControl = true ' http://stackoverflow.com/q/36282024/180275

set vb_editor = acc.vbe
set vb_proj   = vb_editor.activeVBProject
set vb_comps  = vb_proj.VBComponents
    
  
set mdl = vb_comps.Add(1) ' 1 = vbext_ct_StdModule
   
wscript.echo("adding scriptFile " & scriptFile)
mdl.codeModule.addFromFile (scriptFile)
   
mdl.name = "_init"
   
acc.doCmd.close 5, "_init", 1 ' 5=acModule, 1=acSaveYes
  
acc.run("init")
