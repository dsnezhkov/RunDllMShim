/*

-= RundllShim =-
	
	A bridge DLL made to make calling a managed assembly/type/method from an unmanaged dll invoker rundll32.exe easier

	Usage: 
		rundll32.exe
				<path-to-shim>,rdl		// this shim, and the exported function (rdl)
				<CLR version>			// v2.0.50727, v4.0.30319
				<path-to-managed-dll>   // spaces in path are not handled. 
				<assembly class>		// namespace.class

				[<method>|RunCode]		// a. specify static method to invoke of signature:
										// public static int method(string arg) 
										//  -or-
										// b. if omitted it will call your default
										// public static int RunCode(string arg)

				[<string arg>]			// a. specify string argument 
										//  -or-
										// b. if omitted it will pass an empty string
	Example: 
		C:\Windows\System32\rundll32.exe  .\RundllShim.dll,rdl v4.0.30319 C:\Temp\ManagedAssembly.dll KlassLibrary.Klass
		C:\Windows\System32\rundll32.exe  .\RundllShim.dll,rdl v4.0.30319 C:\Temp\ManagedAssembly.dll KlassLibrary.Klass GetLength 
		C:\Windows\System32\rundll32.exe  .\RundllShim.dll,rdl v4.0.30319 C:\Temp\ManagedAssembly.dll KlassLibrary.Klass GetLength "Hello"
*/

#include "stdafx.h"
using namespace std;

HRESULT CallToManagedRT(
	PCWSTR pszVersion,
	PCWSTR pszAssemblyPath,
	PCWSTR pszClassName,
	PCWSTR pszClassMethodName,
	PCWSTR pszClassMethodArg
);

int argsize(const std::string *array);
bool CheckExistence(const char* filename);
std::wstring s2ws(const std::string& s);
void usage(const std::string err);


// our implementation of rdl (a function called by rundll.exe)
namespace workerCode {

	const string defaultMethod = "RunCode";
	const string defaultMethodArg = "";

	void rdl(const short argc, std::string *args)
	{

		cout << argc << endl;
		if (argc < 3) {
			usage("Argument error");
			return;
		}
		string assemblyRunTime = args[0];
		string assemblyPath = args[1];
		string assemblyClass = args[2];
		string assemblyClassMethod = defaultMethod;
		string assemblyClassMethodArg = defaultMethodArg;

		if (argc >= 4 ) {
			cout << "DEBUG: Method Name: " << args[3] << endl;
			assemblyClassMethod = args[3];
		}

		if (argc >= 5 ) {
			cout << "DEBUG Method Arg: " << args[4] << endl;
			assemblyClassMethodArg = args[4];
		}


		cout << "==- Start -==" << endl;
		cout << "DEBUG: Assembly Path: "<< assemblyPath << endl;
		cout << "DEBUG: Assembly Class: "<< assemblyClass << endl;
		cout << "DEBUG: Assembly Method: "<< assemblyClassMethod << endl;
		cout << "DEBUG: Assembly Method Arg: "<< assemblyClassMethodArg << endl;

		if (CheckExistence(assemblyPath.c_str())) {

			CallToManagedRT(
				s2ws(assemblyRunTime).c_str(),
				s2ws(assemblyPath).c_str(),
				s2ws(assemblyClass).c_str(),
				s2ws(assemblyClassMethod).c_str(),
				s2ws(assemblyClassMethodArg).c_str()
			);
		}
		else {
			usage("ERROR: Assembly inaccessible. Check path.");
		}
		cout << "==- End -==" << endl;
	}

} 


// function called by rundll.exe
// make sure to export the name from C++ via <project>.def file
extern "C" __declspec (dllexport) 
void  CALLBACK rdl(
	HWND hwnd,        // handle to owner window
	HINSTANCE hinst,  // instance handle for the DLL
	LPSTR  lpCmdLine, // string the DLL will parse
	int nCmdShow      // show state
)
{
	string lpCmdLineS;
	if (lpCmdLine == NULL) {
		exit(2);
	};
	lpCmdLineS = lpCmdLine;


	// Allocate new child Console so we can see output. rundll32 does not show output by default.
	// TODO: I have not found a good way to hijack parent console yet.
	AllocConsole();

	regex reg("\\s+");
	sregex_token_iterator iter(lpCmdLineS.begin(), lpCmdLineS.end(), reg, -1);
	sregex_token_iterator end;
	vector<string> avec(iter, end);

	string *args;
	args = new string[avec.size()];

	short argc = 0;
	for (int i = 0; i<avec.size(); i++) {
		args[i] = avec[i];//Copy the vector to the string
		argc++;
	}

	
	freopen("CONIN$", "r", stdin);
	freopen("CONOUT$", "w", stdout);
	freopen("CONOUT$", "w", stderr);
	
	// Call out implementation of rdl
	workerCode::rdl(argc, args);

	std::cout << "Press <any> key to close console" << std::endl;
	std::cin.get();

}


int argsize(const std::string *array)
{
	size_t i = 0;
	while (!array[i].empty())
		++i;
	return i;
}

bool CheckExistence(const char* filename)
{
	ifstream test(filename);
	return test.good();
}

std::wstring s2ws(const std::string& s)
{
	int len;
	int slength = (int)s.length() + 1;
	len = MultiByteToWideChar(CP_ACP, 0, s.c_str(), slength, 0, 0);
	wchar_t* buf = new wchar_t[len];
	MultiByteToWideChar(CP_ACP, 0, s.c_str(), slength, buf, len);
	std::wstring r(buf);
	delete[] buf;
	return r;
}

void usage(const std::string err) {
	std::cout << err << std::endl;
	std::cout << "Usage: rundll32.exe  <path-to-shim>,rdl <CLR version> <path-to-managed-dll> <assembly class> [<method>|RunCode] [<string arg>]" << std::endl;
	std::cout << std::endl;
	std::cout << "Example: C:\\Windows\\System32\\rundll32.exe  .\\CLRTry.dll, rdl v4.0.30319 C:\\Temp\\CSClassLibrary.dll CSClassLibrary.CSSimpleObject GetLength \"Hello\"" << std::endl;

}


// Entry point
BOOL APIENTRY DllMain(
	HMODULE hModule,
	DWORD  ul_reason_for_call,
	LPVOID lpReserved
)
{
	return TRUE;
}