The vb folder contains a .BAS module and set of classes that provide an interface to ENet.
The vc folder contains the source code and project files necessary to build a DLL version of ENet that can be used from Visual Basic.
The VB module and classes WILL NOT WORK without the compiled VC++ DLL.
The example folder contains a small example chat server/client that uses the ActiveX DLL. As noted, you will need both the VC++ DLL and the ActiveX DLL for it to work.
To use the chat example, simply start it up and connect to a server. If you don't know of any available servers, you can run your own by selecting 'Start Server'.
If you do not have VC++ to compile the DLL, you can download a compiled version from http://luminance.org/f/vbenet.zip.