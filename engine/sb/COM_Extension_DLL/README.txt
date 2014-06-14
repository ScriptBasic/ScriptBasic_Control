
This is a script basic extension to add support for creating COM objects
and calling their methods.

The COM objects must support the IDispatch interface so they can be used from
scripting clients. (This is very common)

Visual Studio 2008 project files have been included. They are set to use
the script basic include directory as relative path ./../include/

The COM.dll will be compiled into the current directory.

Notes:

Script basic only supports several data types internally, COM supports
many. 

In testing with VB6 COM objects, the VB6 methods were forgiving
with data types. If the VB6 method expected a byte type, it would accept
a long (VT_I4) so long as the value was < 255. 

The two main types this extension handles are longs and strings which it will
proxy between script basic and COM types automatically.

The CallByName function accepts an arbitrary number of arguments. These are 
translated and copied into a DISPPARAMS array to be passed to the COM object.

The prototype of this function is:

callbyname object, "procname", [vbcalltype = VbMethod], [arg0], [arg1] ...

Where object is a long type returned from CreateObject() export.

If you are working with embedding ScriptBasic in your own application,
you can also use CallByName to operate on host created COM objects such as
VB6 Form, Class and GUI objects. 

All you need is for the host application to provide an ObjPtr() to the 
script to give it full access. 

More details on this are available here:
  http://sandsprite.com/blogs/index.php?uid=11&pid=310

The VB6_Example.dll is a sample COM object that COM_VB6_Example.sb script uses
to show the results of some tests with common data types, retrieving strings and
longs, displaying a UI, and manipulating VB6 Form COM objects

In order to use this ActiveX Dll on your system you will have to run regsvr32 on it
or compile it yourself. The easiest way to register it is 

start -> run -> type regsvr32 -> drag and drop the dll file into the run textbox to
have its path added as the argument.

remember com.dll has to be in the script basic modules directory and 
com.inc should be in the includes directory to use these samples.
