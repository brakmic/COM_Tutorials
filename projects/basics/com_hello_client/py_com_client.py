import win32com.client as wincl

def main():
    try:
        hw = wincl.Dispatch("HelloWorldLib.HelloWorld")

        hw.SayHello

        greeting = hw.SayHelloStr
        print(greeting)

        greeting = hw.SayHelloTo("John Doe")
        print(greeting)

    except Exception as e:
        print("An error occurred: {0}".format(e))

if __name__ == "__main__":
    main()
