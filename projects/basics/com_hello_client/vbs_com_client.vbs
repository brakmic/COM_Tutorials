Set hw = CreateObject("HelloWorldLib.HelloWorld")

hw.SayHello

greeting = hw.SayHelloStr
WScript.Echo greeting

greeting = hw.SayHelloTo("John Doe") ' treated as method
WScript.Echo greeting

Set hw = Nothing
