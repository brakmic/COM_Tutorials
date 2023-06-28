cl /c /EHsc HelloWorldDll.cpp
cl /c /EHsc HelloWorldFactory.cpp
cl /c /EHsc HelloWorld.cpp
cl /c /EHsc ./midl/IHelloWorld_i.c

link /dll /def:HelloWorld.def /out:HelloWorld.dll HelloWorldDll.obj HelloWorldFactory.obj HelloWorld.obj IHelloWorld_i.obj Advapi32.lib Shlwapi.lib OleAut32.lib
