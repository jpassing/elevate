#define STRICT
#define UNICODE
#define _UNICODE
#include <windows.h>
#include <shldisp.h>
#include <shlobj.h>
#include <exdisp.h>
#include <stdlib.h>
#include <shlwapi.h>
#include <atlbase.h>
#include <atlalloc.h>
#include "stdafx.h"

/*----------------------------------------------------------------------
 * Purpose:
 *		Execute a process on the command line with elevated rights on Vista
 *
 * Copyright:
 *		Johannes Passing (johannes.passing@googlemail.com)
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.
 *
 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
 * Lesser General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General Public
 * License along with this library; if not, write to the Free Software
 * Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307 USA
 */
#define BANNER L"(c) 2007 - Johannes Passing - http://int3.de/\n\n"

typedef struct _COMMAND_LINE_ARGS
{
  BOOL Debug;
  BOOL ShowHelp;
  BOOL Wait;
  BOOL StartComspec;
  BOOL Hide;
  BOOL Unelevated;
  PCWSTR Dir;
  PCWSTR ApplicationName;
  PCWSTR CommandLine;
} COMMAND_LINE_ARGS, *PCOMMAND_LINE_ARGS;


//from here:https://stackoverflow.com/a/8196291/813599
BOOL IsElevated() {
  BOOL fRet = FALSE;
  HANDLE hToken = NULL;
  if (OpenProcessToken(GetCurrentProcess(), TOKEN_QUERY, &hToken)) {
    TOKEN_ELEVATION Elevation;
    DWORD cbSize = sizeof(TOKEN_ELEVATION);
    if (GetTokenInformation(hToken, TokenElevation, &Elevation, sizeof(Elevation), &cbSize)) {
      fRet = Elevation.TokenIsElevated;
    }
  }
  if (hToken) {
    CloseHandle(hToken);
  }
  return fRet;
}

//following approach is from here: https://blogs.msdn.microsoft.com/oldnewthing/20131118-00/?p=2643
//goes to some effort to use explorer to "ShellExecute" so new process runs under logged in user's context
//this avoids needing to pass in login creds or otherwise prompt for them
//which is the case with using SysInternals psexec launching an un-elevated process for example

void FindDesktopFolderView(REFIID riid, void **ppv)
{
  CComPtr<IShellWindows> spShellWindows;
  spShellWindows.CoCreateInstance(CLSID_ShellWindows);

  CComVariant vtLoc(CSIDL_DESKTOP);
  CComVariant vtEmpty;
  long lhwnd;
  CComPtr<IDispatch> spdisp;
  spShellWindows->FindWindowSW(&vtLoc, &vtEmpty, SWC_DESKTOP, &lhwnd, SWFO_NEEDDISPATCH, &spdisp);

  CComPtr<IShellBrowser> spBrowser;
  CComQIPtr<IServiceProvider>(spdisp)->
    QueryService(SID_STopLevelBrowser,
      IID_PPV_ARGS(&spBrowser));

  CComPtr<IShellView> spView;
  spBrowser->QueryActiveShellView(&spView);

  spView->QueryInterface(riid, ppv);
}

void GetDesktopAutomationObject(REFIID riid, void **ppv)
{
  CComPtr<IShellView> spsv;
  FindDesktopFolderView(IID_PPV_ARGS(&spsv));
  CComPtr<IDispatch> spdispView;
  spsv->GetItemObject(SVGIO_BACKGROUND, IID_PPV_ARGS(&spdispView));
  spdispView->QueryInterface(riid, ppv);
}

class CCoInitialize {
public:
  CCoInitialize() : m_hr(CoInitialize(NULL)) { }
  ~CCoInitialize() { if (SUCCEEDED(m_hr)) CoUninitialize(); }
  operator HRESULT() const { return m_hr; }
  HRESULT m_hr;
};

VOID ShellExecuteFromExplorer(
  PCWSTR pszFile,
  PCWSTR pszParameters = nullptr,
  PCWSTR pszDirectory = nullptr,
  PCWSTR pszOperation = nullptr,
  int nShowCmd = SW_SHOWNORMAL)
{
  CCoInitialize init;

  CComPtr<IShellFolderViewDual> spFolderView;
  GetDesktopAutomationObject(IID_PPV_ARGS(&spFolderView));
  CComPtr<IDispatch> spdispShell;
  spFolderView->get_Application(&spdispShell);

  CComQIPtr<IShellDispatch2>(spdispShell)
    ->ShellExecute(CComBSTR(pszFile),
      CComVariant(pszParameters ? pszParameters : L""),
      CComVariant(pszDirectory ? pszDirectory : L""),
      CComVariant(pszOperation ? pszOperation : L""),
      CComVariant(nShowCmd));
}
/*****************************************************************/


INT Launch(
  __in PCWSTR ApplicationName,
  __in PCWSTR CommandLine,
  __in BOOL Wait,
  __in BOOL Hide,
  __in BOOL Unelevated,
  __in PCWSTR Dir
)
{
  SHELLEXECUTEINFO Shex;
  ZeroMemory(&Shex, sizeof(SHELLEXECUTEINFO));
  Shex.cbSize = sizeof(SHELLEXECUTEINFO);
  Shex.fMask = SEE_MASK_FLAG_NO_UI | SEE_MASK_NOCLOSEPROCESS;
  Shex.lpVerb = Unelevated ? nullptr : L"runas";
  Shex.lpDirectory = Dir;
  Shex.lpFile = ApplicationName;
  Shex.lpParameters = CommandLine;
  Shex.nShow = Hide ? SW_HIDE : SW_SHOWNORMAL;

  //if we're not currently elevated, trying to drop to unelevated, 
  //then we can use full fidelity ShellExecuteEx
  if (!(Unelevated && IsElevated())) {
    if (!ShellExecuteEx(&Shex)) {
      DWORD Err = GetLastError();
      fwprintf(stderr, L"%s could not be launched: %d\n", ApplicationName, Err);
      return EXIT_FAILURE;
    }
    _ASSERTE(Shex.hProcess);
    if (Wait) WaitForSingleObject(Shex.hProcess, INFINITE);
    CloseHandle(Shex.hProcess);
  }
  else {
    //the unelevated execute necessarily flows through Explorer...
    //which means the old ShellExecute API vs ShellExecuteEx...
    //so we don't get a process handle back that we can wait on
    //nor a resultCode to inspect
    ShellExecuteFromExplorer(Shex.lpFile, Shex.lpParameters, Shex.lpDirectory, nullptr/*operation: e.g. "open", "edit", etc*/, Shex.nShow);
  }

  return EXIT_SUCCESS;
}

INT DispatchCommand(__in PCOMMAND_LINE_ARGS Args)
{
  WCHAR AppNameBuffer[MAX_PATH];
  WCHAR CmdLineBuffer[MAX_PATH * 2];

  if (Args->ShowHelp)
  {
    wprintf(
      BANNER
      L"Execute a process on the command line with elevated rights on Vista\n"
      L"\n"
      L"Usage: Elevate [options] prog [args]\n"
      L"-?    - Shows this help\n"
      L"-v    - debug mode\n"
      L"-wait - Waits until prog terminates\n"
      L"-hide - Launches with hidden window\n"
      L"-unel - Will launch without elevation (from a currently elevated process)\n"
      L"          Precludes -wait option.\n"
      L"-dir  - working directory\n"
      L"-k    - Starts the the %%COMSPEC%% environment variable value and\n"
      L"          executes prog in it (CMD.EXE, 4NT.EXE, etc.)\n"
      L"prog  - The program to execute\n"
      L"args  - Optional command line arguments to prog\n");

    return EXIT_SUCCESS;
  }

  if (Args->StartComspec)
  {
    //
    // Resolve COMSPEC
    //
    if (0 == GetEnvironmentVariable(L"COMSPEC", AppNameBuffer, _countof(AppNameBuffer)))
    {
      fwprintf(stderr, L"%%COMSPEC%% is not defined\n");
      return EXIT_FAILURE;
    }
    Args->ApplicationName = AppNameBuffer;

    //
    // Prepend /K and quote arguments
    //
    if (FAILED(StringCchPrintf(
      CmdLineBuffer,
      _countof(CmdLineBuffer),
      L"/K \"%s\"",
      Args->CommandLine)))
    {
      fwprintf(stderr, L"Creating command line failed\n");
      return EXIT_FAILURE;
    }
    Args->CommandLine = CmdLineBuffer;
  }

  //wprintf( L"App: %s,\nCmd: %s\n", Args->ApplicationName, Args->CommandLine );
  return Launch(Args->ApplicationName, Args->CommandLine, Args->Wait, Args->Hide, Args->Unelevated, Args->Dir);
}

int __cdecl wmain(
  __in int Argc,
  __in WCHAR* Argv[]
)
{
  OSVERSIONINFO OsVer;
  COMMAND_LINE_ARGS Args;
  INT Index;
  BOOL FlagsRead = FALSE;
  WCHAR CommandLineBuffer[260] = { 0 };

  ZeroMemory(&OsVer, sizeof(OSVERSIONINFO));
  OsVer.dwOSVersionInfoSize = sizeof(OSVERSIONINFO);

  ZeroMemory(&Args, sizeof(COMMAND_LINE_ARGS));
  Args.CommandLine = CommandLineBuffer;

  //
  // Check OS version
  //
  if (GetVersionEx(&OsVer) &&
    OsVer.dwMajorVersion < 6)
  {
    fwprintf(stderr, L"This tool is for Windows Vista and above only.\n");
    return EXIT_FAILURE;
  }

  //
  // Parse command line
  //
  for (Index = 1; Index < Argc; Index++)
  {
    if (!FlagsRead &&
      (Argv[Index][0] == L'-' || Argv[Index][0] == L'/'))
    {
      PCWSTR FlagName = &Argv[Index][1];
      if (0 == _wcsicmp(FlagName, L"v"))
      {
        Args.Debug = TRUE;
      }
      else if (0 == _wcsicmp(FlagName, L"?"))
      {
        Args.ShowHelp = TRUE;
      }
      else if (0 == _wcsicmp(FlagName, L"wait"))
      {
        Args.Wait = TRUE;
      }
      else if (0 == _wcsicmp(FlagName, L"k"))
      {
        Args.StartComspec = TRUE;
      }
      else if (0 == _wcsicmp(FlagName, L"hide"))
      {
        Args.Hide = TRUE;
      }
      else if (0 == _wcsicmp(FlagName, L"unel"))
      {
        Args.Unelevated = TRUE;
      }
      else if (0 == _wcsicmp(FlagName, L"dir"))
      {
        /* cool, quoted args are were automatically resolved
        //keeping string conversion code for future reference
        #include <string>
        std::wstring s = Argv[++Index];
        if (s.front() == '"') {
          s.erase(0, 1); // erase the first character
          s.erase(s.size() - 1); // erase the last character
        }
        Args.Dir = s.c_str();
        */
        Args.Dir = Argv[++Index];
      }
      else
      {
        fwprintf(stderr, L"Unrecognized Flag %s\n", FlagName);
        return EXIT_FAILURE;
      }
    }
    else
    {
      FlagsRead = TRUE;
      if (Args.ApplicationName == NULL && !Args.StartComspec)
      {
        Args.ApplicationName = Argv[Index];
      }
      else
      {
        if (FAILED(StringCchCat(
          CommandLineBuffer,
          _countof(CommandLineBuffer),
          Argv[Index])) ||
          FAILED(StringCchCat(
            CommandLineBuffer,
            _countof(CommandLineBuffer),
            L" ")))
        {
          fwprintf(stderr, L"Command Line too long\n");
          return EXIT_FAILURE;
        }
      }
    }
  }

  if (Args.Debug)
    wprintf(
      L"ShowHelp:        %s\n"
      L"Wait:            %s\n"
      L"Hide:            %s\n"
      L"Unelevated:      %s\n"
      L"Dir:             %s\n"
      L"StartComspec:    %s\n"
      L"ApplicationName: %s\n"
      L"CommandLine:     %s\n",
      Args.ShowHelp ? L"Y" : L"N",
      Args.Wait ? L"Y" : L"N",
      Args.Hide ? L"Y" : L"N",
      Args.Unelevated ? L"Y" : L"N",
      Args.Dir,
      Args.StartComspec ? L"Y" : L"N",
      Args.ApplicationName,
      Args.CommandLine);

  //
  // Validate args
  //
  if (Argc <= 1)
  {
    Args.ShowHelp = TRUE;
  }

  if (!Args.ShowHelp &&
    ((Args.StartComspec && 0 == wcslen(Args.CommandLine)) ||
    (!Args.StartComspec && Args.ApplicationName == NULL)))
  {
    fwprintf(stderr, L"Invalid arguments\n");
    return EXIT_FAILURE;
  }

  return DispatchCommand(&Args);
}

