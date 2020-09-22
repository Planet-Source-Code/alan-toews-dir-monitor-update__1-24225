To use DirMonDll in your own projects, follow the steps below:

1)Compile the dll:
  In the VB IDE
  Click on AutoSaveDll in the project group window
  click File/Make DirMonDll.dll

2)Register the dll: 
  In Windows:
  Click on the start menu & select Run
  Type Regsvr32 {drive}:\{dll path}\DirMonDll.dll

3)Reference the dll in your own project:
  In the VB IDE click Project/References 
  Select AutoSaveDll from the list and click OK
  
4)In the General Declarations portion of a form or class module, add:
  Dim WithEvents dMon as DirMonDll.FormSettings

5)In the form_load procedure add
  Set dMon = New DirMonDll.FormSettings

=============
The dMon object is now ready to use. See the frmTest.frm for an example of how to use it.

