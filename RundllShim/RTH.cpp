
/********************** ******************************************************
* 
* Some parts of this code is modified from MSFT by Op_Nomad and the license is:
* 
*
* THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, 
* EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED 
* WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
****************************** Module Header ******************************/

/****************************** MSFT Module Header ******************************\
 * Module Name:  RuntimeHostV4.cpp
 * Project:      CppHostCLR
 * Copyright (c) Microsoft Corporation.
 * 
 * The code in this file demonstrates using .NET Framework 4.0 Hosting 
 * Interfaces (http://msdn.microsoft.com/en-us/library/dd380851.aspx) to host 
 * .NET runtime 4.0, load a .NET assebmly, and invoke a type in the assembly.
 * 
 * This source is subject to the Microsoft Public License.
 * See http://www.microsoft.com/en-us/openness/licenses.aspx#MPL.
 * All other rights reserved.
 * 
 * THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, 
 * EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED 
 * WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
\***************************************************************************/


#pragma region Includes and Imports
#include <windows.h>

#include <metahost.h>
#pragma comment(lib, "mscoree.lib")

// Import mscorlib.tlb (Microsoft Common Language Runtime Class Library).
#import "mscorlib.tlb" raw_interfaces_only				\
    high_property_prefixes("_get","_put","_putref")		\
    rename("ReportEvent", "InteropServices_ReportEvent")
using namespace mscorlib;
#pragma endregion


//
//   FUNCTION: CallToManagedRT(PCWSTR, PCWSTR)
//
//   PURPOSE: The function demonstrates using .NET Framework 4.0 Hosting 
//   Interfaces to host a .NET runtime, and use the ICLRRuntimeHost interface
//   that was provided in .NET v2.0 to load a .NET assembly and invoke its 
//   type. Because ICLRRuntimeHost is not compatible with .NET runtime v1.x, 
//   the requested runtime must not be v1.x.
//   
//   If the .NET runtime specified by the pszVersion parameter cannot be 
//   loaded into the current process, the function prints ".NET runtime <the 
//   runtime version> cannot be loaded", and return.
//   
//   If the .NET runtime is successfully loaded, the function loads the 
//   assembly identified by the pszAssemblyName parameter. Next, the function 
//   invokes the public static function you pass of the class and print the result.
//
//   PARAMETERS:
//   * pszVersion - The desired DOTNETFX version, in the format ¡°vX.X.XXXXX¡±. 
//     The parameter must not be NULL. It¡¯s important to note that this 
//     parameter should match exactly the directory names for each version of
//     the framework, under C:\Windows\Microsoft.NET\Framework[64]. Because 
//     the ICLRRuntimeHost interface does not support the .NET v1.x runtimes, 
//     the current possible values of the parameter are "v2.0.50727" and 
//     "v4.0.30319". Also, note that the "v" prefix is mandatory.
//   * pszAssemblyPath - The path to the Assembly to be loaded.
//   * pszClassName - The name of the Type that defines the method to invoke.
//
//   RETURN VALUE: HRESULT of the run
//   Added dynamic method name and argument
//
HRESULT CallToManagedRT(
	PCWSTR pszVersion,
	PCWSTR pszAssemblyPath,
	PCWSTR pszClassName,
	PCWSTR pszClassStaticMethodName,
	PCWSTR pszClassMethodArg)
{
    HRESULT hr;
    DWORD dwRet;

    ICLRMetaHost *pMetaHost = NULL;
    ICLRRuntimeInfo *pRuntimeInfo = NULL;

    // ICorRuntimeHost and ICLRRuntimeHost are the two CLR hosting interfaces
    // supported by CLR 4.0. We will use the ICLRRuntimeHost interface that 
    // was provided in .NET v2.0 to support CLR 2.0 new features. 
    // ICLRRuntimeHost does not support loading the .NET v1.x runtimes.
    ICLRRuntimeHost *pClrRuntimeHost = NULL;

	wprintf(L"DEBUG: Passed for Execution: %s %s %s %s %s\n", 
			pszVersion, pszAssemblyPath, pszClassName, pszClassStaticMethodName, pszClassMethodArg);


    // 
    // Load and start the .NET runtime.
    // 

    wprintf(L"DEBUG: Load and start the .NET runtime %s \n", pszVersion);

    hr = CLRCreateInstance(CLSID_CLRMetaHost, IID_PPV_ARGS(&pMetaHost));
    if (FAILED(hr))
    {
        wprintf(L"ERROR: CLRCreateInstance failed w/hr 0x%08lx\n", hr);
        goto Cleanup;
    }


	// Display available CLR runtimes. This info may be useful for parameterizing
	// PCWSTR pszVersion
	// From: http://mode19.net/posts/clrhostingright/

	IEnumUnknown *installedRuntimes;
	hr = pMetaHost->EnumerateInstalledRuntimes(&installedRuntimes);
    if (FAILED(hr))
    {
        wprintf(L"ERROR: ICLRMetaHost::EnumerateInstalledRuntimes failed w/hr 0x%08lx\n", hr);
        goto Cleanup;
    }

	ICLRRuntimeInfo *runtimeInfo = NULL;
	ULONG fetched = 0;
	while ((hr = installedRuntimes->Next(1, (IUnknown **)&runtimeInfo, &fetched)) == S_OK && fetched > 0) {
		wchar_t versionString[20];
		DWORD versionStringSize = 20;
		hr = runtimeInfo->GetVersionString(versionString, &versionStringSize);

		if (versionStringSize >= 2){
			wprintf(L"DEBUG: Available CLR Version : %s\n", versionString);
		}
	}

    // Get the ICLRRuntimeInfo corresponding to a particular CLR version. It 
    // supersedes CorBindToRuntimeEx with STARTUP_LOADER_SAFEMODE.
    hr = pMetaHost->GetRuntime(pszVersion, IID_PPV_ARGS(&pRuntimeInfo));
    if (FAILED(hr))
    {
        wprintf(L"ERROR: ICLRMetaHost::GetRuntime failed w/hr 0x%08lx\n", hr);
        goto Cleanup;
    }

    // Check if the specified runtime can be loaded into the process. This 
    // method will take into account other runtimes that may already be 
    // loaded into the process and set pbLoadable to TRUE if this runtime can 
    // be loaded in an in-process side-by-side fashion. 
    BOOL fLoadable;
    hr = pRuntimeInfo->IsLoadable(&fLoadable);
    if (FAILED(hr))
    {
        wprintf(L"ERROR: ICLRRuntimeInfo::IsLoadable failed w/hr 0x%08lx\n", hr);
        goto Cleanup;
    }

    if (!fLoadable)
    {
        wprintf(L"ERROR: .NET runtime %s cannot be loaded\n", pszVersion);
        goto Cleanup;
    }

    // Load the CLR into the current process and return a runtime interface 
    // pointer. ICorRuntimeHost and ICLRRuntimeHost are the two CLR hosting  
    // interfaces supported by CLR 4.0. Here we use the ICLRRuntimeHost 
    // interface that was provided in .NET v2.0 to support CLR 2.0 new 
    // features. ICLRRuntimeHost does not support loading the .NET v1.x 
    // runtimes.
    hr = pRuntimeInfo->GetInterface(CLSID_CLRRuntimeHost, 
        IID_PPV_ARGS(&pClrRuntimeHost));
    if (FAILED(hr))
    {
        wprintf(L"ERROR: ICLRRuntimeInfo::GetInterface failed w/hr 0x%08lx\n", hr);
        goto Cleanup;
    }

    // Start the CLR.
    hr = pClrRuntimeHost->Start();
    if (FAILED(hr))
    {
        wprintf(L"ERROR: CLR failed to start w/hr 0x%08lx\n", hr);
        goto Cleanup;
    }

    // 
    // Load the NET assembly and call the static method arg: pszClassStaticMethodName of 
    // the type pszClassName in the assembly.
    // 

    wprintf(L"DEBUG: Load the assembly %s\n", pszAssemblyPath);

    // The invoked method of ExecuteInDefaultAppDomain must have the 
    // following signature: static int pszClassStaticMethodName (String pszClassMethodArg)
    // where pszClassStaticMethodName represents the name of the invoked method, and 
    // pszClassMethodArgrepresents the string value passed as a parameter to that 
    // method. 

	// If the HRESULT return value of ExecuteInDefaultAppDomain is 
    // set to S_OK, pReturnValue is set to the integer value returned by the 
    // invoked method. Otherwise, pReturnValue is not set.

    hr = pClrRuntimeHost->ExecuteInDefaultAppDomain(pszAssemblyPath, 
        pszClassName, pszClassStaticMethodName, pszClassMethodArg , &dwRet);

    if (FAILED(hr))
    {
        wprintf(L"ERROR: Failed to call %s(%s) w/hr 0x%08lx\n", pszClassStaticMethodName, pszClassMethodArg,  hr);
        goto Cleanup;
    }

    // Print the call result of the static method.
    wprintf(L"DEBUG: Call %s.%s(\"%s\") => %d\n", pszClassName, pszClassStaticMethodName, 
        pszClassMethodArg, dwRet);

Cleanup:

    if (pMetaHost)
    {
        pMetaHost->Release();
        pMetaHost = NULL;
    }
    if (pRuntimeInfo)
    {
        pRuntimeInfo->Release();
        pRuntimeInfo = NULL;
    }
    if (pClrRuntimeHost)
    {
        // Please note that after a call to Stop, the CLR cannot be 
        // reinitialized into the same process. This step is usually not 
        // necessary. You can leave the .NET runtime loaded in your process.
        //wprintf(L"Stop the .NET runtime\n");
        //pClrRuntimeHost->Stop();

        pClrRuntimeHost->Release();
        pClrRuntimeHost = NULL;
    }

    return hr;
}