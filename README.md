This vbscript will run any command hidden.
Using Windows Scripting Host (WSH) is the only *native Windows* way to run a script completely hidden.
Alternatively, you could compile and EXE, but ... meh.
PowerShell doesn't run things hidden yet; since [`-WindowStyle Hidden` isn't sufficient](https://github.com/PowerShell/PowerShell/issues/3028).
Hopefully we'll have [a `pwshw.exe` soon](https://github.com/PowerShell/PowerShell/issues/3028#issuecomment-367169480) and this repo can be antiquated.

**If you're wanting to run PowerShell hidden, try [HiddenPowershell](https://github.com/UNT-CAS/HiddenPowershell) for more terse execution.**

# Usage

## Download HiddenRun.vbs

I suggest grabbing it at boot with a **startup script**, via GPO.
Users won't see the the powershell console of a startup script, so it's not invasive.
Be sure to adjust the `$path`:

> powershell.exe -ExecutionPolicy Bypass -NoProfile -NonInteractive -WindowStyle Hidden -Command "$path = 'C:\Temp'; [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; Invoke-WebRequest -Uri 'https://raw.githubusercontent.com/UNT-CAS/HiddenRun/v1.0/HiddenRun.vbs' -OutFile ('{0}\HiddenRun.vbs' -f $path) -UseBasicParsing"

:bangbang: Wherever you put it, be sure *users* can read, but not write to it.

Be sure you check for the latest release.
I don't expect a lot of changes to this script, but now that it's open source ... who knows?

I know this seems simple, but pratical implementation is usually a bit more complex.
[Here's how I've made it happen](https://github.com/UNT-CAS/HiddenPowershell/wiki/Practical-Download).

## Execute Command

**Do not use `cscript.exe`; it will cause a console window to appear.**

```
wscript.exe HiddenRun.vbs C:\Temp\CustomFoo.bat
```

This Will run your custom batch script in a completely hidden console by calling the batch script like this:

```
C:\Temp\CustomFoo.bat
```

If you have machines that have Windows Scripting Host (WSH) file extensions (like `.vbs`) disassociated from WSH; then you will need to add the `//E:vbscript` parameter:
```
wscript.exe //E:vbscript HiddenRun.vbs ...
```

# Logging

Logging is done to *Event Viewer*.
There will be two events for every run of the script. One at the start of the run, and the other at the completion/finish.
The details of the logs are:

- **Event Path:** `Windows Logs\Application`
- **Source:** `WSH`
- **Event ID:** *Depends on Status*
  - *Success*: `0` Script Finished; Command Exited with `0`.
  - *Error*: `1` Script Finished; Command Exited with something other than `0`.
  - *Information*: `4` Script Starting

## Start

The *Event ID* of the *starting* message will always be `4` (informational).
Here's an example of what that will look like:

```
HiddenRun Running: 
	C:\Temp\HiddenRun.vbs
	C:\Temp\CustomFoo.bat
```

##  Finish

The *Event ID* of the *finished* message will be `0` (success).
If your command exits with a non-zero exit code, the *Event ID* will be `1` (error).

Here's an example of what a *success* looks like; *Event ID* is `0`:
```
HiddenRun Exited: 
	C:\Temp\HiddenRun.vbs
	C:\Temp\CustomFoo.bat
	Exit Code: 0
```

Here's an example of what an *error* looks like; *Event ID* is `1`:
```
HiddenRun Exited: 
	C:\Temp\HiddenRun.vbs
	C:\Temp\CustomFoo.bat
	Exit Code: -1
```
