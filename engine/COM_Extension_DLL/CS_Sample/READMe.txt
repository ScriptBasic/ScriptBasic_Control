
this is a sample .NET dll which is COM visible.

1) double click the .reg file to register the COM component
on your system.

2) place sample.dll in your script basic /bin directory

Run sample cs_date.sb



Long story:

3 ways to register a .net dll for com.

(note you must set the ComVisible(true) in the properties file)

1) compile the dll on your system (do not have to copy to script basic /bin directory)

2) use regasm.exe [Sample.dll path]

regasm is included with the .NET runtime, but is not in the system
path, you will have to search for it. Example path

C:\WINDOWS>dir regasm* /s/b
C:\WINDOWS\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe

3) An alternative way to register your .NET dll for COM is to generate
a .reg file with

regasm sample.dll /regfile

Then users can just double click the .reg file to register it on their
system. 

A sample .reg file that works for this dll has been provided.

methods 2 and 3 require the dll to be in the exes home directory.
Not sure why.