# RunDllMShim

_A bridge DLL to make calling a managed assembly/type/method from an unmanaged dll invoker rundll32.exe easier)_

## RunDll Pipelining

`Rundll32.exe => Unamanaged Shim => Managed Assembly -> (class,method)`


## Usage

```
rundll32.exe
	<path-to-shim>,rdl	// this shim, and the exported function (rdl)
	<CLR version>		// v2.0.50727, v4.0.30319
	<path-to-managed-dll>   // spaces in path are not handled. 
	<assembly class>	// namespace.class
	[<method>|RunCode]	// a. specify static method to invoke of signature:
				// public static int method(string arg) 
				//  -or-
				// b. if omitted it will call your default
				// public static int RunCode(string arg)

	[<string arg>]		// a. specify string argument 
				//  -or-
				// b. if omitted it will pass an empty string
```

## Example

```
C:\Windows\System32\rundll32.exe  .\RundllShim.dll,rdl v4.0.30319 C:\Temp\ManagedAssembly.dll KlassLibrary.Klass
C:\Windows\System32\rundll32.exe  .\RundllShim.dll,rdl v4.0.30319 C:\Temp\ManagedAssembly.dll KlassLibrary.Klass GetLength 
C:\Windows\System32\rundll32.exe  .\RundllShim.dll,rdl v4.0.30319 C:\Temp\ManagedAssembly.dll KlassLibrary.Klass GetLength "Hello"
```
