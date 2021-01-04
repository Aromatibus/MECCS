@if(0)==(0) rem /* Start of Batch
@echo off
setlocal
call :SetESC
title Mobile Environment Creation Support Script
echo.
echo %ESC%[102mMobile Environment Creation Support Script%ESC%[0m
echo.
echo We can help you create an environment to take advantage
echo of the portable version of the software in a portable environment
echo such as USB memory or virtual disk.
echo Copyright(c) 2021 Aromatibus (https://github.com/Aromatibus)
echo Translated with www.DeepL.com/Translator (free version)
echo.
:: Switch to JScript mode
cscript.exe //nologo //E:JScript "%~f0" %*
echo.
:: Make ConfigFile
if %errorlevel% equ 10002 (
  goto :eof
)
:: Reboot Administrator
if %errorlevel% equ 10001 (
  echo.
  echo %ESC%[105m
::  timeout /t 2 /nobreak
  goto :eof
)
:: Normal End
if %errorlevel% equ 10000 (
  echo %ESC%[102mThe process is finished.%ESC%[0m
  echo.
  timeout /t 5 /nobreak
  goto :eof
)
:: Error
echo %ESC%[41;97mThere was an error in the script.%ESC%[0m
echo ErrorLevel=%errorlevel%
echo.
pause
goto :eof
:SetESC
for /F "tokens=1,2 delims=#" %%a ^
in ('"prompt #$H#$E# & echo on & for %%b in (1) do rem"') ^
do (set ESC=%%b)
@exit /b
gap
rem End of Batch */
@end
/*** Start of JScript <JavaScript ECMAScript3?> ***/

// Flag to display the execution status in the command prompt window.
var DebugFlg = true;
// Flags that work with internal information
//  without creating an external configuration file.
var OneFileMode = false;

// Polyfill trim
if (!String.prototype.trim) {
  String.prototype.trim = function () {
    return this.replace(/^[\s　\uFEFF\xA0]+|[\s　\uFEFF\xA0]+$/g, '');
  };
}
// Polyfill indexOf
var NotFound = -1;
if (!Array.prototype.indexOf) {
  Array.prototype.indexOf = function(obj, start){
    for(var i = (start || 0), j = this.length; i < j; i++){
      if(this[i] === obj){ return i }
    }
    return NotFound;
  };
}
// Polyfill repeat
if (!String.prototype.repeat) {
  String.prototype.repeat = function(count) {
    return Array(count * 1 + 1).join(this);
  };
}
// Polyfill isArray
if (!Array.isArray) {
  Array.isArray = function(arg) {
    return Object.prototype.toString.call(arg) === '[object Array]';
  };
}

// Main
function main() {
  // Batch %errorlevel%
  var ErrorlevelABNormal = 0; // Script terminated abnormally.(not use)
  var ErrorlevelABNormal = 1; // Script terminated abnormally.(not use)
  var ErrorlevelNormal = 10000;
  var ErrorlevelRebootAdmin = 10001;
  var ErrorlevelMakeConfig = 10002;
  // object variable
  var ws = new ActiveXObject("WScript.Shell");
  var wseProcess = ws.Environment("Process");
  var wseSystem = ws.Environment("System");
  var wseUser = ws.Environment("User");
  var wseVolatile = ws.Environment("Volatile");
  var wsn = new ActiveXObject("WScript.Network");
  var sfo = new ActiveXObject("Scripting.FileSystemObject");
  // environment reference variable
  var RunningDrv = sfo.GetDriveName(WScript.ScriptFullName);
  var ReferenceDrv = RunningDrv;
  var RunningFolder = sfo.GetParentFolderName(WScript.ScriptFullName);
  var RunningFileName = sfo.GetBaseName(WScript.ScriptFullName);
  var ConfigFileName = RunningFileName + ".txt";
  var ScriptArgArray = GetArrayArgs();
  var AdminFlg = isAdmin();
  // constant
  var RegFound = 0; // When discovered, it will be greater than or '0'.
  var RegNotFound = -1;
  var MODE = {
    READING: 1,
    WRITING: 2,
    APPENDING: 8,
    CREATE: true,
    NOT_CREATE: false,
    WAIT: true,
    NOT_WAIT: false
  }
  var WINSTYLE = {
    HIDDEN: 0,
    NORMAL: 1,
    MINIMUM: 2,
    MAXIMUM: 3,
    INACTIVE_NORMAL: 4,
    LAST_TIME: 5,
    INACTIVE_MINIMUM: 7
  }
  // Variable
  var ReadData;
  var CommandName;
  var ArgArray = [];
  var ArgLength;
  var RegEx;
  var TmpPATH = "";
  var ExeArgFlg = false;
  // Set the current folder to the script folder.
  ws.CurrentDirectory = RunningFolder;
  // Creating and reading configuration files.
  if (OneFileMode) {
    ConfigText = CreateConfigFile("");
  } else {
    if (!sfo.FileExists(ConfigFileName)) {
      CreateConfigFile(ConfigFileName);
      return ErrorlevelMakeConfig;
    }
    var otf = sfo.OpenTextFile(ConfigFileName, MODE.READING);
    var ConfigText = otf.ReadAll();
    otf.Close();
  }
  // Format the configuration file.
  // Append CRLF to the end of the text.
  ConfigText = ConfigText + "\r\n";
  // Unified to CRLF.
  ConfigText = ConfigText.replace(/(\r|\n|\r\n)/g, "\r\n");
  // Remove white space at the beginning and end of lines.(TRIM)
  ConfigText = ConfigText.replace(/(\r\n)[\s　\uFEFF\xA0]+/g, "\r\n");
  ConfigText = ConfigText.replace(/[\s　\uFEFF\xA0]+(\r\n)/g, "\r\n");
  // Delete comment line.
  ConfigText = ConfigText.replace(/^#.*$(\r\n)/gm, "");
  // blank line deletion.
  ConfigText = ConfigText.replace(/^$(\r\n)/g, "");
  ConfigText = ConfigText.replace(/(\r\n)^$/m, "");
  // Process prioritized commands.
  RegEx = /^RunAdministrator\s*=\s*true$/gmi;
  if (ConfigText.search(RegEx) !== RegNotFound) {
    if (!(AdminFlg & 2)) {
      DebugMsg("Re-execute with administrative privileges.\n");
      ExecuteCommand(WScript.ScriptFullName, ScriptArgArray, "RunAs");
      return ErrorlevelRebootAdmin;
    }
  }
  // Decompose the configuration information into an array by row
  ConfigText = ConfigText.split("\r\n");
  // If "PATH" in the process environment variable
  // does not end with ";", ";" will be added.
  if (wseProcess.item("PATH").slice(-1) !== "\;") { TmpPATH = "\;"; }
  // Display information on user and execution process privileges.
  DebugMsg("User Name    = " + wsn.UserName);
  if (AdminFlg & 1) {
    DebugMsg("User Account = Administrator");
  } else {
    DebugMsg("User Account = User");
  }
  DebugMsg("Computer     = " + wsn.ComputerName);
  if (AdminFlg & 2) {
    DebugMsg("Process      = Administrator\n");
  } else {
    DebugMsg("Process      = User\n");
  }
  // Argument information display.
  if (ScriptArgArray.length === 0) {
    DebugMsg("Arguments nothing.\n");
  } else {
    DebugMsg("Arguments = [" + ScriptArgArray + "]\n");
  }
  // Environment variable information display.
  DebugMsg("Reference Drive = \"" + ReferenceDrv + "\"\n");
  DebugMsg("Process PATH =\n[" + wseProcess.item("PATH") + "]\n");
  DebugMsg("System PATH =\n[" + wseSystem.item("PATH") + "]\n");
  DebugMsg("User PATH =\n[" + wseUser.item("PATH") + "]\n");
  DebugMsg("Volatile PATH =\n[" + wseVolatile.item("PATH") + "]\n");
  // Process configuration information array by array.
  for (var i=0; i < ConfigText.length; i++){
    ReadData = ConfigText[i];
//    DebugMsg("Command [" + NumberFormat(i, 2) + "] : " + ReadData);
    // Pre-processing to add variables
    // to the environment variable "PATH".
    if (sfo.FolderExists(SameDrv(ReadData, ReferenceDrv))) {
      ReadData = SameDrv(ReadData, ReferenceDrv);
      RegEx = new RegExp(ReadData.replace(/\\/g, "\\\\"), "gi");
      if (wseProcess.item("PATH").search(RegEx) === RegNotFound) {
        TmpPATH = TmpPATH + ReadData + ";";
        DebugMsg("PATH Attached.   = [" + ReadData + "]");
      } else {
        DebugMsg("PATH Unattached. = [" + ReadData + "]");
      }
    }
    // Remove Volatile environment variables.
    RegEx = /^RemoveVolatile\s*=\s*true$/i;
    if (ReadData.search(RegEx) === RegFound) {
      ws.Exec("cmd /c reg delete "
          + "\"HKEY_CURRENT_USER\\Volatile Environment\"");
      DebugMsg("Registry \"HKEY_CURRENT_USER\\Volatile Environment\""
          + " removed.");
    }
    // Remove named environment variables.
    RegEx = /^(.*)\s*:=:$/gi;
    if (ReadData.search(RegEx) === RegFound) {
      ws.Exec("cmd /c reg delete "
          + "\"HKEY_CURRENT_USER\\Volatile Environment\" /v \""
          + ReadData.replace(RegEx, "$1")
          + "\" /f");
      wseVolatile.item(ReadData.replace(RegEx, "$1")) = "";
      wseProcess.item(ReadData.replace(RegEx, "$1")) = "";
      DebugMsg(ReadData.replace(RegEx, "$1") + "=(Removed)");
    }
    // Add a named environment variable.
    RegEx = /^(.*)\s*:=:\s*(.*)$/gi;
    if (ReadData.search(RegEx) === RegFound) {
      ReadData = SameDrv(ReadData, ReferenceDrv);
      wseVolatile.item(ReadData.replace(RegEx, "$1"))
          = ReadData.replace(RegEx, "$2");
      wseProcess.item(ReadData.replace(RegEx, "$1"))
          = ReadData.replace(RegEx, "$2");
      DebugMsg(ReadData.replace(RegEx, "$1") + "="
             + ReadData.replace(RegEx, "$2") + "\n");
    }
    // Change the drive to reference.
    RegEx = /^ReferenceDrive\s*=$/gi;
    if (ReadData.search(RegEx) === RegFound) {
      ReferenceDrv = RunningDrv;
      DebugMsg("Set the referenced drive to \"" + ReferenceDrv + "\".\n");
    }
    RegEx = /^ReferenceDrive\s*=\s*(\"|\')?([A-Z]):?\\?(\"|\')?$/gi;
    if (ReadData.search(RegEx) === RegFound) {
      RegEx = ReadData.replace(RegEx, "$2").toUpperCase() + ":";
      if (!sfo.FolderExists(RegEx)) {
        DebugMsg("The drive specified in "
            + "\"ReferenceDrive\" is not found."
            + " [" + RegEx + "]\n");
      } else {
        ReferenceDrv = RegEx;
        DebugMsg("Set the referenced drive to \"" + ReferenceDrv + "\"\n");
      }
    }
    // Standby processing for specified number of seconds.
    RegEx = /^WaitSec\((\d+)\)$/gi;
    if (ReadData.search(RegEx) === RegFound) {
      DebugMsg("");
      WaitSecond(parseInt(ReadData.replace(RegEx, "$1")))
    }
    // Create a shortcut to the desktop.
    CommandName = "DesktopShortcut";
    ArgLength = 2;
    if (SplitCommand(CommandName, ReadData, ArgLength) !== false) {
      ArgArray = SplitCommand(CommandName, ReadData, ArgLength);
      ArgArray = SameDrv(ArgArray, ReferenceDrv);
      MakeDesktopShortcut(ArgArray[0], ArgArray[1]);
    }
    // Create a symbolic link. Requires administrator privileges.
    CommandName = "RelativeMKLink";
    ArgLength = 2;
    if (SplitCommand(CommandName, ReadData, ArgLength) !== false) {
      ArgArray = SplitCommand(CommandName, ReadData, ArgLength);
      ArgArray = SameDrv(ArgArray, ReferenceDrv);
      RelativeMKLink(ArgArray[0], ArgArray[1]);
    }
    // Mount the virtual disk. Requires administrator privileges.
    CommandName = "VDiskAttach";
    ArgLength = 2;
    if (SplitCommand(CommandName, ReadData, ArgLength) !== false) {
      ArgArray = SplitCommand(CommandName, ReadData, ArgLength);
      VDiskAttach(SameDrv(ArgArray[0], ReferenceDrv), ArgArray[1]);
    }
    // UnMount the virtual disk. Requires administrator privileges.
    CommandName = "VDiskDetach";
    ArgLength = 1;
    if (SplitCommand(CommandName, ReadData, ArgLength) !== false) {
      ArgArray = SplitCommand(CommandName, ReadData, ArgLength);
      ArgArray = SameDrv(ArgArray[0], ReferenceDrv);
      VDiskDetach(ArgArray[0]);
    }
    // Execute the specified string.
    CommandName = "Execute";
    ArgLength = 1;
    if (SplitCommand(CommandName, ReadData, ArgLength) !== false) {
      ArgArray = SplitCommand(CommandName, ReadData, ArgLength);
      ArgArray = SameDrv(ArgArray, ReferenceDrv);
      DebugMsg("\n" + CommandName + " : [" + ArgArray + "]");
      // Add to environment variable "PATH" with volatile attribute.
      if (TmpPATH !== "") {
        wseVolatile.item("PATH") = wseVolatile.item("PATH") + TmpPATH;
        wseProcess.item("PATH") = wseProcess.item("PATH") + TmpPATH;
        DebugMsg("\nAttach Volatile PATH = \n["
            + wseVolatile.item("PATH") + "]\n");
        TmpPATH = "";
      }
      // Confirm that the file is there, and then run it.
      if (sfo.FileExists(ArgArray)) {
        ExecuteCommand(ArgArray);
      } else {
        DebugMsg("Execute : Error Execute File Not Found.\n");
      }
    }
    // Execute a string with specified arguments.
    CommandName = "ExecuteArg";
    ArgLength = 1;
    if (SplitCommand(CommandName, ReadData, ArgLength) !== false) {
      ArgArray = SplitCommand(CommandName, ReadData, ArgLength);
      ArgArray = SameDrv(ArgArray, ReferenceDrv);
      ExeArgFlg = true;
      DebugMsg("\n" + CommandName + " : [" + ArgArray + "]");
      // Add to environment variable "PATH" with volatile attribute.
      if (TmpPATH !== "") {
        wseVolatile.item("PATH") = wseVolatile.item("PATH") + TmpPATH;
        wseProcess.item("PATH") = wseProcess.item("PATH") + TmpPATH;
        DebugMsg("\nAttach Volatile PATH = \n["
            + wseVolatile.item("PATH") + "]\n");
        TmpPATH = "";
      }
      // Confirm that the file is there, and then run it.
      if (sfo.FileExists(ArgArray)) {
        if (ScriptArgArray.length != 0) {
          ExecuteCommand(ArgArray, ScriptArgArray);
        } else {
          DebugMsg("No Arguments.\n");
          ExecuteCommand(ArgArray);
        }
      } else {
        DebugMsg("Error Execute File Not Found.\n");
      }
    }
  }
  // Add to environment variable "PATH" with volatile attribute.
  if (TmpPATH !== "") {
    wseVolatile.item("PATH") = wseVolatile.item("PATH") + ";" + TmpPATH;
    wseProcess.item("PATH") = wseProcess.item("PATH") + ";" + TmpPATH;
    DebugMsg("\nAttach Volatile PATH = \n["
        + wseVolatile.item("PATH") + "]\n");
  }
  // If there is an argument and "ExecuteArg"
  // is not in the configuration information,
  // the argument will be executed.
  if ((ScriptArgArray.length > 0) && (!ExeArgFlg)) {
    if (ScriptArgArray.length > 1) {
      ExecuteCommand(ScriptArgArray[0], ScriptArgArray.slice(1));
    } else {
      ExecuteCommand(ScriptArgArray);
    }
  }
  return ErrorlevelNormal;
}

// Display information for debugging.
function DebugMsg(TextMessage, opt1, opt2) {
  DebugFlg = (typeof DebugFlg !== "undefined") ? DebugFlg : false;
  if (DebugFlg !== true) { return; }
  var ws = new ActiveXObject("WScript.Shell");
  var sfo = new ActiveXObject("Scripting.FileSystemObject");
  var FILE = {
    READING: 1,
    WRITING: 2,
    APPENDING: 8,
    CREATE: true,
    NOT_CREATE: false
  }
  ws.CurrentDirectory = sfo.GetParentFolderName(WScript.ScriptFullName);
  var opt = [];
  opt.push((typeof opt1 !== "undefined")
      ? opt1.toString().toUpperCase() : "");
  opt.push((typeof opt2 !== "undefined")
      ? opt2.toString().toUpperCase() : "");
  if(opt[0] === "FILE") {
    if (!sfo.FileExists(opt[1])) {
      wsOut = sfo.OpenTextFile(opt[1], FILE.WRITING, FILE.CREATE);
    } else {
      wsOut = sfo.OpenTextFile(opt[1], FILE.APPENDING, FILE.NOT_CREATE);
    }
    wsOut.WriteLine(TextMessage);
    wsOut.Close();
    return;
  }
  var wsOut = WScript.StdOut;
  if(opt.indexOf("STDERR") > NotFound) { wsOut = WScript.StdErr; }
  if(opt.indexOf("STDOUT") > NotFound) { wsOut = WScript.StdOut; }
  if(opt.indexOf("WRITE") > NotFound) {
    wsOut.Write(TextMessage);
  } else {
    wsOut.WriteLine(TextMessage);
  }
}

// Integers are separated by three digits,
// decimal places are truncated,
// and justified to the specified number of characters.
// Non-numeric characters will be aligned
// to the specified number of characters.
function NumberFormat(FloatNum, Digits, DecPoint) {
  DecPoint =
      (typeof DecPoint !== "undefined") ? DecPoint : 0;
  if (DecPoint > 0) { Digits = Digits - DecPoint - 1; }
      var RegEx = new RegExp(/^[+,-]?([1-9]\d*|0)(\.\d+)?$/);
  if (RegEx.test(FloatNum)) {
    var IntNum = Math.floor(parseFloat(FloatNum, 10));
    var IntNumStr = IntNum.toString()
        .replace( /(\d)(?=(\d\d\d)+(?!\d))/g, '$1,');
    var SmallNum = (parseFloat(FloatNum, 10) - IntNum); // Bugがある?
    var SmallNumStr = "";
    if (DecPoint !== 0) {
      SmallNumStr = ("." + SmallNum.toString().slice(2, DecPoint + 2)
                  + "0".repeat(DecPoint)).slice(0, DecPoint + 1);
    }
    var SameDigitChar = "";
    if ((Digits - IntNumStr.length) > 0) {
      SameDigitChar = " ".repeat(Digits - IntNumStr.length);
    }
    return (SameDigitChar + IntNumStr + SmallNumStr);
  } else {
    var SameDigitChar = " ".repeat(Digits - FloatNum.toString().length);
    return (SameDigitChar + FloatNum.toString());
  }
}

// Convert byte strings to numbers, optional unit conversion.
function ConvByteStr(ByteStr, DataType) {
  DataType = (typeof DataType !== "undefined") ? DataType : "Byte";
  var RegEx = new RegExp("(bit|Byte|kB|KB|MB|GB|TB|KiB|MiB|GiB|TiB)", "i");
  if (!RegEx.test(DataType.toString())) { DataType = "Byte"; }
  var ByteUnits = {
    bit: 8,
    kB: 1000,
    MB: Math.pow(1000, 2),
    GB: Math.pow(1000, 3),
    TB: Math.pow(1000, 4),
    KB: Math.pow(2, 10),
    KiB: Math.pow(2, 10),
    MiB: Math.pow(2, 20),
    GiB: Math.pow(2, 30),
    TiB: Math.pow(2, 40)
  }
  var RegEx = new RegExp("^(\\d{1,3}(,?\\d{3})*(\\.\\d+)?)(\\s+)?"
            + "(kB|KB|MB|GB|TB|KiB|MiB|GiB|TiB)?$", "g");
  var ResultRegEx = RegEx.exec(ByteStr);
  try {
      var ByteValue = parseFloat(ResultRegEx[1], 10);
  } catch (error) {
    return false;
  }
  with(ByteUnits) {
    switch (ResultRegEx[ResultRegEx.length - 1]) {
      case "kB":
        ByteValue *= kB;
        break;
      case "KB":
        ByteValue *= KB;
        break;
      case "KiB":
        ByteValue *= KiB;
        break;
      case "MB":
        ByteValue *= MB;
        break;
      case "MiB":
        ByteValue *= MiB;
        break;
      case "GB":
        ByteValue *= GB;
        break;
      case "GiB":
        ByteValue *= GiB;
        break;
      case "TB":
        ByteValue *= TB;
        break;
      case "TiB":
        ByteValue *= TiB;
        break;
      case "":
        ByteValue = ByteValue;
        break;
      Default:
        return false;
    }
    // DataTypeに応じて戻り値を返します。
    switch (DataType) {
      case "bit":
        ByteValue *= bit;
        break;
      case "Byte":
        ByteValue = ByteValue;
        break;
      case "kB":
        ByteValue /= kB;
        break;
      case "KB":
        ByteValue /= KB;
        break;
      case "KiB":
        ByteValue /= KiB;
        break;
      case "MB":
        ByteValue /= MB;
        break;
      case "MiB":
        ByteValue /= MiB;
        break;
      case "GB":
        ByteValue /= GB;
        break;
      case "GiB":
        ByteValue /= GiB;
        break;
      case "TB":
        ByteValue /= TB;
        break;
      case "TiB":
        ByteValue /= TiB;
        break;
      Default:
        return false;
    }
  }
  return ByteValue;
}

// Wait for specified number of seconds.
function WaitSecond(WaitSec, TextMsg) {
  TextMsg = (typeof TextMsg !== "undefined") ? TextMsg : " Wait ";
  var wsStdOut = WScript.StdOut;
  var StartMsec = new Date();
  WaitMsec = parseInt(WaitSec) * 1000;
  do {
    WaitSec
        = parseInt((WaitMsec - (new Date() - StartMsec)) / 1000 + 0.9);
    if (TextMsg !== "") {
      DebugMsg("\r" + TextMsg + " " + WaitSec + "seconds.",
          "StdOut", "Write");
    }
  } while (WaitSec);
}

// Array the arguments.
function GetArrayArgs(){
  var ScriptArgArray = [];
  for(var i = 0; i < WScript.Arguments.Count(); i++){
    ScriptArgArray.push(WScript.Arguments(i));
  }
  return ScriptArgArray;
}

// Array the command arguments.
function SplitCommand(CommandName, CommandText, ArgLength) {
  ArgLength = (typeof ArgLength !== "undefined") ? ArgLength : "";
  var RegNotFound = -1;
  var RegEx = new RegExp("^" + CommandName + "\\(\\s*?(.*)\\s*?\\)$", "i");
  if (CommandText.search(RegEx) === RegNotFound) { return false; }
  var ArgText = CommandText.replace(RegEx, "$1");
  if (ArgLength === 0) { return ArgText; }
  var ArgArray = ArgText.split(",");
  var RegEx = new RegExp("((\"|\')?(\\w+)(\"|\')?)", "gi");
  for (var i=0; i < ArgArray.length; i++){
    ArgArray[i] = ArgArray[i].replace(RegEx, "$3").trim();
  }
  if (ArgLength !== "") { ArgArray.length = ArgLength; }
  return ArgArray;
}

// Change the drive in the string to the specified drive.
function SameDrv(TextStr, ReferenceDrv) {
  var sfo = new ActiveXObject("Scripting.FileSystemObject");
  ReferenceDrv = (typeof ReferenceDrv !== "undefined")
      ? ReferenceDrv
      : sfo.GetDriveName(WScript.ScriptFullName);
  var RegEx = /^[a-z]/i;
  if (RegEx.test(ReferenceDrv)) {
    if (ReferenceDrv.length === 1) {
      ReferenceDrv = ReferenceDrv + ":";
    }
    ReferenceDrv = sfo.GetDriveName(ReferenceDrv) + "\\";
    if (Array.isArray(TextStr)) {
      for (var i = 0; i < TextStr.length; i++) {
        TextStr[i] = TextStr[i].replace(/^\\/g, ReferenceDrv);
        TextStr[i] = TextStr[i].replace(/[a-z]:\\/gi, ReferenceDrv);
      }
    } else {
        TextStr = TextStr.replace(/^\\/g, ReferenceDrv);
        TextStr = TextStr.replace(/[a-z]:\\/gi, ReferenceDrv);
    }
  }
  return TextStr;
}

// Obtaining administrator authority information.
function isAdmin() {
  var wsn = new ActiveXObject("WScript.Network");
  var ResultCode = 0;
  if (isAdminUser(wsn.UserName)) { ResultCode += 1; }
  if (isAdminProcess()) { ResultCode += 2; };
  return ResultCode;
}

// Obtain administrative privilege information for the logged-in user.
function isAdminUser(UserName) {
  var ws = new ActiveXObject("WScript.Shell");
  var ResultString;
  var ResultArray = [];
  var wslocal = ws.Exec("net localgroup Administrators");
  var wsdomain = ws.Exec("net localgroup Administrators /domain");
  ResultString = wslocal.StdOut.ReadAll();
  ResultString += wsdomain.StdOut.ReadAll();
  ResultArray = ResultString.toString().split("\r\n");
  if(ResultArray.indexOf(UserName) === -1) {
    return false;
  }
  return true;
}

// Obtain administrative privileges for the running process.
function isAdminProcess() {
  var ws = new ActiveXObject("WScript.Shell");
  var AppName = "openfiles";
  var ResultCode = false;
  var MODE = {
    WAIT: true,
    NOT_WAIT: false
  }
  var WINSTYLE = {
    HIDDEN: 0,
    NORMAL: 1,
    MINIMUM: 2,
    MAXIMUM: 3,
    INACTIVE_NORMAL: 4,
    LAST_TIME: 5,
    INACTIVE_MINIMUM: 7
  }
  if(ws.Run(AppName, WINSTYLE.HIDDEN, MODE.WAIT) === 0) {
    ResultCode = true;
  }
  return(ResultCode);
}

// Run the specified file in a new process.
function ExecuteCommand(ExeFName, AryArgs, Opt0, Opt1, Opt2) {
/*
** This may be a bug, but the ShellExecute function and the Powershell
**  start-process command handle arguments differently
**  when run as "RunAs" with administrative privileges than
**  when run normally.
** Several patterns have been reported to Microsoft.
** This script uses the ShellExecute function to run Powershell
**  with administrative privileges, and then executes Powershell
**  with the desired command and arguments to avoid this.
*/
  var sa = new ActiveXObject("Shell.Application");
  var sfo = new ActiveXObject("Scripting.FileSystemObject");
  var vOperation = {
    EDIT: "edit", // not use
    FIND: "find", // not use
    OPEN: "open", // Run or Open
    PRINT: "print", // not use
    PROPERTIES: "properties", // not use
    RUNAS: "runas" // Aministrator
  }
  var vShow = {
    HIDDEN: 0,
    NORMAL: 1,
    MINIMUM: 2,
    MAXIMUM: 3,
    INACTIVE_NORMAL: 4,
    LAST_TIME: 5,
    INACTIVE_MINIMUM: 7,
    APPLICATION: 10
  }
  var pwsWin = {
    NoWin: "-NoNewWindow",
    Normal: "-WindowStyle Normal",
    Hidden: "-WindowStyle Hidden",
    Minimized: "-WindowStyle Minimized",
    Maximized: "-WindowStyle Maximized"
  }
  if (!sfo.FileExists(ExeFName)) { return false; }
  var sFile = "powershell.exe";
  AryArgs = (typeof AryArgs !== "undefined") ? AryArgs : "";
  AryArgs = (AryArgs instanceof Array)
           ? AryArgs.join("\\\" \\\"") : AryArgs;
  AryArgs = "\\\"" + AryArgs + "\\\"";
  var vDirectory = sfo.GetParentFolderName(WScript.ScriptFullName);
  Opt0 = (typeof Opt0 !== "undefined") ? Opt0 : "";
  Opt1 = (typeof Opt1 !== "undefined") ? Opt1 : "";
  Opt2 = (typeof Opt2 !== "undefined") ? Opt2 : "";
  var Opt_Ary = [Opt0.trim(), Opt1.trim(), Opt2.trim()];
  var OptArgs = "";
  var ExeFlg = vOperation.OPEN;
  var WinFlg = pwsWin.Normal;
  for(var i = 0; i < Opt_Ary.length; i++) {
    switch (Opt_Ary[i].toString().toUpperCase()) {
      case "RUNAS":
        ExeFlg = vOperation.RUNAS;
        break;
      case "NOWIN":
      case "NOWINDOW":
      case "NONEWWINDOW":
        WinFlg = pwsWin.NoWin;
        break;
      case "NORMAL":
        WinFlg = pwsWin.Normal;
        break;
      case "HIDDEN":
        WinFlg = pwsWin.Hidden;
        break;
      case "MINIMIZED":
      case "MIN":
        WinFlg = pwsWin.Minimized;
        break;
      case "MAXIMIZED":
      case "MAX":
        WinFlg = pwsWin.Maximized;
        break;
      default:
        OptArgs = Opt_Ary[i].trim();
    }
  }
  var vArgs = "start-process "
            + "-FilePath '" + ExeFName + "' "
            + "-WorkingDirectory '" + vDirectory + "' "
            + WinFlg;
  OptArgs = ("\\\"" + OptArgs + "\\\" "
           + AryArgs).replace(/\\\"\\\"/gm, "").trim();
  if (OptArgs !== "") {
    vArgs = vArgs + " -ArgumentList '" + OptArgs + "'";
  }
  DebugMsg("Operation = " + ExeFlg);
  DebugMsg("Command : " + sFile + " " + vArgs);
  DebugMsg("");
  return sa.ShellExecute(sFile, vArgs, vDirectory, ExeFlg, vShow.HIDDEN);
}

// Get the title from the destination html of the specified url.
function GetHPTitle(url){
  var RegEx = /^https?:\/\/[\w!\?/\+\-_~=;\.,\*&@#\$%\(\)'\[\]]+$/g;
  if (url.search(RegEx) === 0) {
    var http = WScript.CreateObject("Msxml2.ServerXMLHTTP");
    http.open("GET", url, false);
    try {
      http.send();
    } catch (e) {
      return false;
    }
    var htmlfile = WScript.CreateObject("htmlfile");
    htmlfile.write("<meta http-equiv='x-ua-compatible' content='IE=11'>"
        + http.responseText);
    var gethp = htmlfile.parentWindow;
    var hptitle = gethp.document.querySelector("title");
    return hptitle.textContent;
  } else {
    return false;
  }
}

// Create a shortcut on the desktop.
function MakeDesktopShortcut(ShortcutName, TargetFile) {
  var ws = new ActiveXObject("WScript.Shell");
  var sfo = new ActiveXObject("Scripting.FileSystemObject");
  var WINSTYLE = {
    HIDDEN: 0,
    NORMAL: 1,
    MINIMUM: 2,
    MAXIMUM: 3,
    INACTIVE_NORMAL: 4,
    LAST_TIME: 5,
    INACTIVE_MINIMUM: 7
  }
  var Shortcut_Full_Name = false;
  var RegEx = /^["']?(.*?)["']?$/g;
  ShortcutName = ShortcutName.replace(RegEx, "$1");
  var RegEx = /^https?:\/\/[\w!\?/\+\-_~=;\.,\*&@#\$%\(\)'\[\]]+$/g;
  if (TargetFile.search(RegEx) === 0) {
    if (ShortcutName === "") {
      ShortcutName = GetHPTitle(TargetFile);
    }
    Shortcut_Full_Name = ws.SpecialFolders("Desktop") + "\\"
        + ShortcutName + ".url";
  } else if (sfo.FileExists(TargetFile) || sfo.FolderExists(TargetFile)) {
    if (ShortcutName === "") {
      ShortcutName = sfo.GetBaseName(TargetFile);
    }
    Shortcut_Full_Name = ws.SpecialFolders("Desktop") + "\\"
        + ShortcutName + ".lnk";
  } else {
    DebugMsg("Shortcut Target Not Found.");
  }
  if (sfo.FileExists(Shortcut_Full_Name)) {
    sfo.DeleteFile(Shortcut_Full_Name, true);
  }
  if (Shortcut_Full_Name !== false) {
    DebugMsg("Shortcut Created.");
    var wscs = ws.CreateShortcut(Shortcut_Full_Name);
    wscs.TargetPath = TargetFile;
    if (sfo.GetExtensionName(Shortcut_Full_Name) === "lnk") {
      wscs.WindowStyle = WINSTYLE.MAXIMUM;
      wscs.IconLocation = TargetFile + ", 0";
      wscs.WorkingDirectory = ws.SpecialFolders("Desktop");
      wscs.Description = "Created by "
          + sfo.GetBaseName(WScript.ScriptFullName);
    }
    try {
      wscs.Save();
    } catch(e) {
      DebugMsg("Error(" + e + ")");
    }
  }
  DebugMsg("  Shortcut File Name : " + Shortcut_Full_Name);
  DebugMsg("  Target File Name   : " + TargetFile);
  DebugMsg("");
}

// Create symbolic links with relative paths.
function RelativeMKLink(LinkFile, TargetFile) {
  var ws = new ActiveXObject("WScript.Shell");
  var sfo = new ActiveXObject("Scripting.FileSystemObject");
  var RunningDrv = sfo.GetDriveName(WScript.ScriptFullName) + "\\";
  var MKLink_Cmd = "cmd /c mklink";
  var ErrFlg = 0;
  var MODE = {
    WAIT: true,
    NOT_WAIT: false
  }
  var WINSTYLE = {
    HIDDEN: 0,
    NORMAL: 1,
    MINIMUM: 2,
    MAXIMUM: 3,
    INACTIVE_NORMAL: 4,
    LAST_TIME: 5,
    INACTIVE_MINIMUM: 7
  }
  LinkFile = LinkFile.replace(/^\\/, "");
  LinkFile = LinkFile.replace(/^[a-z]:\\/gi, "");
  TargetFile = TargetFile.replace(/^\\/, "");
  TargetFile = TargetFile.replace(/^[a-z]:\\/gi, "");
  DebugMsg("\nRelative MKLink");
  DebugMsg("  Link   File : " + RunningDrv + LinkFile);
  DebugMsg("  Target File : " + RunningDrv + TargetFile);
  if(!isAdminProcess()) {
    DebugMsg("\nError : Process Not Administrator.");
    ErrFlg += 1;
  }
  if (sfo.FileExists(RunningDrv + LinkFile)) {
    DebugMsg("\nError : Exist Link File.");
    ErrFlg += 2;
  }
  if (!sfo.FolderExists(sfo.GetParentFolderName(RunningDrv + LinkFile))) {
    DebugMsg("\nError : Link Folder Not Found."
        + sfo.GetParentFolderName(RunningDrv + LinkFile));
    ErrFlg += 4;
  }
  if (!sfo.FileExists(RunningDrv + TargetFile)) {
    DebugMsg("\nError : Target File Not Found.");
    ErrFlg += 8;
  }
  if(ErrFlg === 0) {
    MKLink_Cmd = MKLink_Cmd + " \""
        + RunningDrv + LinkFile + "\" \""
        + Array(LinkFile.split("\\").length).join("..\\")
        + TargetFile + "\"";
    DebugMsg("  Command     : " + MKLink_Cmd);
    ws.Run(MKLink_Cmd, WINSTYLE.HIDDEN, MODE.NOT_WAIT);
  }
  DebugMsg("");
}

// Mount or optimize (compact) the virtual file.
function VDiskAttach(VDiskName, VDOpt) {
  var ws = new ActiveXObject("WScript.Shell");
  var sfo = new ActiveXObject("Scripting.FileSystemObject");
  var VDiskCMD = {
    CMD: "cmd /c (",
    SelectFile: "echo select vdisk file=",
    SelectVolume: "echo select volume=",
    Attach: "&&echo attach vdisk",
    AttachReadOnly: "&&echo attach vdisk readonly",
    Detach: "&&echo detach vdisk",
    Assign: "&&echo assign letter=",
    VolumeList: "&&echo list volume",
    DetailDisk: "&&echo detail disk",
    DetailVDISK: "&&echo detail vdisk",
    Compact: "&&echo compact vdisk",
    Exit: "&&echo exit)|Diskpart.exe"
  }
  var VDError = {
    NotAdministrator: -1,
    FileNotFound: -2,
    UseDriveLetter: -8
  }
  var NotFound = -1;
  var ResultStr;
  var RegEx;
  var ResultRegEx;
  if (isAdminProcess() !== true) {
    DebugMsg("\nError : Process Not Administrator.");
    return VDError.NotAdministrator;
  }
  // Check the options, even if there is an error, go through them.
  VDOpt = (typeof VDOpt !== "undefined")
         ? VDOpt.toString().toUpperCase() : "";
  if (VDOpt !== "COMPACT") {
    if (VDOpt.length === 1) { VDOpt = VDOpt + ":"; }
    VDOpt = sfo.GetDriveName(VDOpt);
  }
  if (sfo.GetParentFolderName(VDiskName) === "") {
    VDiskName = sfo.GetParentFolderName(WScript.ScriptFullName)
        + "\\" + VDiskName;
  }
  if (!sfo.FileExists(VDiskName)) {
    DebugMsg("\nError : VDISK Not Found. [" + VDiskName + "]\n");
    return VDError.FileNotFound;
  }
  with(VDiskCMD) {
    var VolumeCheck
      = CMD
      + SelectFile + "\"" + VDiskName + "\""
      + Detach
      + VolumeList
      + Exit;
    var AttachVDisk
      = CMD
      + SelectFile + "\"" + VDiskName + "\""
      + Detach
      + Attach
      + VolumeList
      + DetailDisk
      + Exit;
    var AttachROVDisk
      = CMD
      + SelectFile + "\"" + VDiskName + "\""
      + Detach
      + AttachReadOnly
      + VolumeList
      + DetailDisk
      + Exit;
    var DetachVDisk
      = CMD
      + SelectFile + "\"" + VDiskName + "\""
      + Detach
      + Exit;
  }
  var VolumeNums = [];
  var ConnectedDrvs = [];
  RegEx = new RegExp("\\s+Volume\\s+(\\d+)\\s+([A-Z])?\\s+", "gi");
  ResultStr = ws.Exec(VolumeCheck).StdOut.ReadAll();
  while ((ResultRegEx = RegEx.exec(ResultStr)) !== null) {
    VolumeNums.push(ResultRegEx[1]);
    ConnectedDrvs.push(ResultRegEx[2]);
  }
  if (ConnectedDrvs.indexOf(VDOpt) > NotFound) {
    DebugMsg("The specified drive \"" + VDOpt + "\" is in use.");
    VDError.UseDriveLetter;
  }
  if (VDOpt !== "COMPACT") {
    ResultStr = ws.Exec(AttachVDisk).StdOut.ReadAll();
  } else {
    ResultStr = ws.Exec(AttachROVDisk).StdOut.ReadAll();
  }
  var i = 0;
  while ((ResultRegEx = RegEx.exec(ResultStr)) !== null) {
    if (VolumeNums[i] !== ResultRegEx[1]) { break; }
    i++;
  }
  with(VDiskCMD) {
    var CompactVDisk
      = CMD
      + SelectVolume + ResultRegEx[1]
      + CompactVDisk
      + DetailDisk
      + Exit;
    var AssignVDisk
      = CMD
      + SelectVolume + ResultRegEx[1]
      + Assign + VDOpt
      + DetailDisk
      + Exit;
  }
  if (VDOpt !== "") {
    if (VDOpt === "COMPACT") {
      ResultStr = ws.Exec(CompactVDisk).StdOut.ReadAll();
      VDOpt = ResultRegEx[2];
    } else {
      ResultStr = ws.Exec(AssignVDisk).StdOut.ReadAll();
    }
  } else {
    VDOpt = ResultRegEx[2];
  }
  var FileInfo = sfo.GetFile(VDiskName);
  DebugMsg("Attach VDisk.");
  DebugMsg("VDisk File Name             : " + VDiskName);
  DebugMsg(" File Type                  : " + FileInfo.Type);
  DebugMsg(" File Size                  : "
      + NumberFormat(ConvByteStr(FileInfo.Size, "GiB"), 9, 3)
      + " GiB ("
      + NumberFormat(FileInfo.Size, 16) + " Byte)");
  DebugMsg("");
  var DriveInfo = sfo.GetDrive(VDOpt);
  DebugMsg("Mounted on Drive            : " + DriveInfo.DriveLetter);
  DebugMsg(" File System                : " + DriveInfo.FileSystem);
  DebugMsg(" Drive Total Size           : "
      + NumberFormat(ConvByteStr(DriveInfo.TotalSize, "GiB"), 9, 3)
      + " GiB ("
      + NumberFormat(DriveInfo.TotalSize, 16) + " Byte)");
  DebugMsg(" Drive Free Space           : "
      + NumberFormat(ConvByteStr(DriveInfo.FreeSpace, "GiB"), 9, 3)
      + " GiB ("
      + NumberFormat(DriveInfo.TotalSize, 16) + " Byte)");
  DebugMsg("");
  return VDOpt;
}

// Unmount the virtual files.
function VDiskDetach(VDiskName) {
  var ws = new ActiveXObject("WScript.Shell");
  var sfo = new ActiveXObject("Scripting.FileSystemObject");
  var ResultStr;
  var RegEx;
  var ResultRegEx;
  var VolumeNo;
  var VDiskCMD = {
    CMD: "cmd /c (",
    SelectFile: "echo select vdisk file=",
    SelectVolume: "echo select volume=",
    Attach: "&&echo attach vdisk",
    AttachReadOnly: "&&echo attach vdisk readonly",
    Detach: "&&echo detach vdisk",
    Assign: "&&echo assign letter=",
    VolumeList: "&&echo list volume",
    DetailDisk: "&&echo detail disk",
    DetailVDISK: "&&echo detail vdisk",
    Compact: "&&echo compact vdisk",
    Exit: "&&echo exit)|Diskpart.exe"
  }
  var VDError = {
    NotAdministrator: -1,
    FileNotFound: -2,
    OptStrError: -4,
    UseDriveLetter: -8
  }
  if (isAdminProcess() !== true) {
    return VDError.NotAdministrator;
  }
  if (sfo.GetParentFolderName(VDiskName) === "") {
    VDiskName = sfo.GetParentFolderName(WScript.ScriptFullName)
        + "\\" + VDiskName;
  }
  if (!sfo.FileExists(VDiskName)) {
    DebugMsg("\nError : VDisk Not Found. [" + VDiskName + "]\n");
    return VDError.FileNotFound;
  }
  var DetachVDisk
      = VDiskCMD.CMD
      + VDiskCMD.SelectFile + "\"" + VDiskName + "\""
      + VDiskCMD.Detach
      + VDiskCMD.Exit;
  ResultStr = ws.Exec(DetachVDisk).StdOut.ReadAll();
  DebugMsg("Detach VDisk.");
  DebugMsg("VDisk File Name             : " + VDiskName);
  DebugMsg("");
}

// Create a BASE64 string
function MakeBase64Str(SizeByte, bitStr) {
  bitStr = (typeof bitStr !== "undefined") ? bitStr : "000000";
  var RegEx = new RegExp("^[01]{6}$", "");
  if (!RegEx.test(bitStr.toString())) { return false; }
  // Some expressions in Base64 are omitted because the goal is to fill
  // them with a sequence of consecutive binary numbers.
  var Base64CharTBL = ("ABCDEFGHIJKLMNOPQRSTUVWXYZ"
                    + "abcdefghijklmnopqrstuvwxyz"
                    + "0123456789+/").split("");
  // Binary number sequence to be written (one character per 6-bit unit)
  var Base64Char = Base64CharTBL[parseInt(bitStr, 2)];
  var Base64PaddingChar = "=";
  // Convert the size specified in bytes to bits
  var SizeBit = ConvByteStr(SizeByte, "bit");
  // Add the missing bits so that the bit data can be divided
  // into 6-bit units.
  var Base64CalcStep1  = SizeBit + (6 - (SizeBit % 6));
  // Calculates the number of characters divided by one digit (6 bits)
  // of Base64 characters.
  var Base64CalcStep2 = (Base64CalcStep1 / 6);
  // Calculate the number of missing characters to divide
  // a Base64 string by 4 digits.
  var Base64CalcStep3 = (4 - (Base64CalcStep2 % 4));
  // Create a Base64 string
  var Base64Str
      = Base64Char.repeat(Base64CalcStep2)
      + Base64PaddingChar.repeat(Base64CalcStep3);
  return Base64Str;
}

// Convert Base64 string to binary data.
function Base64ToBin(Base64Str) {
var xml = WScript.CreateObject("Microsoft.XMLDOM");
  var node = xml.createElement("base64-node");
  node.dataType = "bin.base64";
  node.text = Base64Str;
  return node.nodeTypedValue;
}

// Creates a binary zero file in the free space of the specified drive
// and then deletes it.
function MakeZeroFillFile(DriveLetter, FileSize, bitStr) {
// Maximum number of characters that can be stored in a
// JScript variable is 16MiB - 1Byte
// Maximum size of one file that can be handled by
// "ADODB.Stream" is 2GiB - 1bit.
// Limit the number of characters to 8MiB and the file size to 1GiB.
  var fso = new ActiveXObject("Scripting.FileSystemObject");
  if (DriveLetter.length === 1) { DriveLetter = DriveLetter + ":"; }
  DriveLetter = fso.GetDriveName(DriveLetter);
  if (!fso.FolderExists(DriveLetter)) { return false; }
  bitStr = (typeof bitStr !== "undefined") ? bitStr : "000000";
  var RegEx = new RegExp("^[01]{6}$", "");
  if (!RegEx.test(bitStr.toString())) { return false; }
  var ByteUnits = {
    bit: 8,
    MiB: Math.pow(2, 20),
    GiB: Math.pow(2, 30)
  }
  var StreamFile = {
    TypeBinary: 1,
    TypeText: 2,
    Create: 1,
    OverWrite: 2,
    ReadAll: -1,
    ReadLine: -2,
    WriteChar: 0,
    WriteLine: 1
  }
  var i = 0;
  // Convert the file size to be created
  var FileSizeByte = ConvByteStr(FileSize);
  // Amount of data to be written to a single file at a time
  var OneTimeMaxByte = 8 * ByteUnits.MiB;
  // Maximum amount of data per file
  var OneFileMaxByte = 1 * ByteUnits.GiB;
  // Number of times to write to the maximum amount of data per file.
  // If the number of times is not divisible, an error will occur
  // because of the error in the file size.
  var OneFileWriteCount = OneFileMaxByte / OneTimeMaxByte;
  if ((OneFileMaxByte % OneTimeMaxByte) !== 0) { return false; }
  // Number of times to write in the maximum amount of data per file
  // for the file size to be created.
  var FileWriteCount = Math.floor(FileSizeByte / OneFileMaxByte);
  // The maximum amount of data per file
  // that remains after multiple writes.
  var RemainingByte = FileSizeByte % OneFileMaxByte;
  // Number of times to write to the remaining data
  var RemainingWriteCount = Math.floor(RemainingByte / OneTimeMaxByte);
  // Final amount of data remaining (recalculation)
  var RemainingByte = RemainingByte % OneTimeMaxByte;

/*
  DebugMsg("FileSize            = " + NumberFormat(FileSize, 15, 0));
  DebugMsg("FileSizeByte        = " + NumberFormat(FileSizeByte, 15, 0));
  DebugMsg();
  DebugMsg("OneTimeMaxByte      = " + NumberFormat(OneTimeMaxByte, 15, 0));
  DebugMsg("OneFileMaxByte      = " + NumberFormat(OneFileMaxByte, 15, 0));
  DebugMsg();
  DebugMsg("FileWriteCount      = " + NumberFormat(FileWriteCount, 15, 0));
  DebugMsg("RemainingByte       = " + NumberFormat(FileSizeByte % OneFileMaxByte, 15, 0));
  DebugMsg();
  DebugMsg("RemainingWriteCount = " + NumberFormat(RemainingWriteCount, 15, 0));
  DebugMsg("RemainingByte       = " + NumberFormat(RemainingByte, 15, 0));
  DebugMsg();
  DebugMsg("Base64Str = " + bitStr + " : " + MakeBase64Str(1, bitStr));
  DebugMsg();
*/

  // Convert Base64 string to binary data.
  var OneTimeStr = MakeBase64Str(OneTimeMaxByte, bitStr);
  var OneTimeBin = Base64ToBin(OneTimeStr);
  var RemainingStr = MakeBase64Str(RemainingByte, bitStr);
  var RemainingBin = Base64ToBin(RemainingStr);
  // Create a folder and files for ZeroFill.
  var FolderName = DriveLetter + "\\@ZEROFILL";
  // Folder creation
  i = "";
  while (fso.FolderExists(FolderName + i)) { i++; }
  FolderName = FolderName + i;
  fso.CreateFolder(FolderName);
  // Binary file creation
  var ads = new ActiveXObject("ADODB.Stream");
  ads.Type = StreamFile.TypeBinary;
  var FileNo = 0;
  // Create a 1GiB file
  ads.Open();
  if (FileWriteCount > 0) {
    i = OneFileWriteCount;
    while (i > 0) {
      ads.Write(OneTimeBin);
      ads.Position = ads.Size;
      i--;
    }
    i = FileWriteCount;
     while (i > 0) {
      FileNo++;
      FileName = FolderName + "\\"
               + "0".repeat(8 - FileNo.toString().length)
               + FileNo.toString() + ".bin";
      DebugMsg("Write File : " + FileName
          + " (" + NumberFormat(ads.Size, 13, 0) + " Bytes)");
      ads.SaveToFile(FileName, StreamFile.OverWrite);
      i--;
    }
  }
  ads.Close();
  // Create files of less than 1 GiB
  ads.Open();
  i = RemainingWriteCount;
  while (i > 0) {
    ads.Write(OneTimeBin);
    ads.Position = ads.Size;
    i--;
  }
  if (RemainingByte > 0) { ads.Write(RemainingBin); }
  if (ads.Size > 0) {
    FileNo++;
    FileName = FolderName + "\\"
             + "0".repeat(8 - FileNo.toString().length)
             + FileNo.toString() + ".bin";
    DebugMsg("Write File : " + FileName
        + " (" + NumberFormat(ads.Size, 13, 0) + " Bytes)");
    ads.SaveToFile(FileName, StreamFile.OverWrite);
  }
  ads.Close();
  fso.DeleteFolder(FolderName, true);
  return true;
}

// Create a file of configuration information.
function CreateConfigFile(ConfigFileName) {
  var ConfigText = function() {/*
#
# Mobile Environment Creation Support Script
# Copyright(c) 2021 Aromatibus (https://github.com/Aromatibus)
#
# We can help you create an environment to take advantage
# of the portable version of the software in a portable environment
# such as USB memory or virtual disk.
#
# 1.What I can do.
#  1.1 You can set the environment based on the drive where the script
#      was executed or the drive you specify.
#  1.2 The environment variables you set will be restored to their
#      previous state after shutdown.
#  1.3 Allows you to create shortcuts on the desktop.
#  1.4 You can easily create symbolic links with relative paths.
#  1.5 Virtual disk files (VHD, VHDX) can be mounted and unmounted
#      on the specified drive.
#  1.6 Execute programs.
#  1.7 Drag and drop arguments or files to execute them,
#      or specify arguments to execute a script.
#
# 2.The file name of the script can be changed.
#  2.1 You can use a name that matches your environment.
#  2.2 The configuration file "script-name.txt" will be loaded.
#  2.3 The extensions "CMD" and "BAT" for batch files
#      and "JS" for JScript can be used.
#  2.4 If you change the extension to "JS", change "true" in
#      "var Debug_Flg = true;" in the script to "false"
#      because the message cannot be displayed.
#
# 3.About configuration files
#  If you change "false" in "var OneFileMode = false;" to "true"
#  in the script, the configuration information in the script
#  will be read instead of creating an external configuration file.
#
# 4.Please refer to this configuration file
#       for how to write the configuration file.
#  4.1 Lines beginning with "#" are comment lines.
#
# 5.About the "PATHEXT" environment variable
#  5.1 At the command prompt, if there are files with the same name,
#      the order of priority of the extensions to be executed
#      is generally "EXE" > "BAT" > "CMD" > "VBS" > "JS".
#      Check the environment variable "PATHEXT" for details.
#  5.2 In the sample configuration file, the environment variable
#      "PATHEXT" is changed so that the file extension "CMD"
#      is executed first.
#  5.3 The extension "PY" is added so that it can be treated as an
#      executable file.
#  5.4 It will be easier to distinguish between them if you decide
#      that "CMD" is a JScript type batch file.
#
# 6."RunAdministrator=true" means
#  If the executed process does not have administrative privileges,
#  it will be rerun with administrative privileges.
#  Default setting is off (false).
#  Specify this when executing a command or program that requires
#  administrative privileges.
#  This command will be executed with the highest priority.
#
# 7."ReferenceDrive=DriveLetter" means
#  Change the base reference drive.
#  Default is the drive where you ran the script.
#
# 8.Setting the environment variable "PATH" (if you set a folder)
#  Drive name is changed and added to the environment variable "PATH".
#  Environment variable "PATH" will be set last unless "Execute"
#  and "ExecuteArg" are executed.
#  Drive name will be rewritten to the name of the drive
#  where the script was executed by default.
#  If the drive where the script was executed is drive D, then
#    a:\abc -> D:\abc
#    \abc   -> D:\abc
#  Drive name will be replaced as shown above.
#  If the same folder exists in the environment variable "PATH",
#  or if the folder does not exist, it will be ignored.
#  If you have a folder with the same name, it will have the same
#  priority.
#  If a file with the same name is in a different folder with a higher
#  priority, that one will be given priority.
#  When the application execution alias is enabled.
#  For some software such as Python.exe, the "PATH" set in
#  the application execution alias takes priority.
#  If the application execution alias is enabled, some software,
#  such as Python.exe, will prioritize the "PATH"set in the application
#  execution alias, and may not be executed in the specified environment
#  due to the later reference order.
#  "PATH" set in the application execution alias takes precedence.
#  In the first configuration file created, a symbolic link is created
#  with the name "PYT".
#  This can be avoided by running "PYT" instead of "Python.exe".
#
# 9.":=:" means
#  Set the environment variable with ":=:" as the border,
#  with the left side as the environment variable name
#  and the right side as the variable.
#  If the environment variable name is already in use,
#  it will be overwritten, but will return after shutdown.
#  If the right side is not specified, the environment variable
#  is deleted.
#  Drive name will be overwritten in the same way as the "PATH" setting.
#  If you do not specify the right side, the environment variable
#  will be overwritten, but it will return after shutdown.
#
# 10."RemoveVolatile=true" means
#  Remove all Volatile environment variables from the registry.
#  The Volatile environment variables are supposed to be reset
#  after reboot, but sometimes they are not, so we prepared this command.
#  Normally, use ":=:" to delete them.
#  This is a command that should reset the environment variables
#  to their default values.
#  The results of executing the command are at your own risk.
#
# 11."DesktopShortcut(StringA, StringB)" means
#  Create a shortcut on the user's desktop.
#  Drive name will be rewritten in the same way as the "PATH" setting.
#  String A becomes the shortcut name and will be overwritten
#  if there is another shortcut with the same name.
#  If you omit the shortcut name by enclosing it in "",
#  it becomes a file name, a folder name, or a URL title.
#  String B specifies the file name, folder name, or URL starting with
#  "http" of the link source.
#  Drive name can be rewritten in the same way as the "PATH" setting.
#
# 12."RelativeMKLink(Link file name, Link source file)" means
#  Create a symbolic link with the link file name.
#  File must be specified as an absolute path.
#  Link file will be created with a PATH relative to the source file,
#  so it will function properly even if the drive changes.
#  Drive name will be rewritten in the same way as the "PATH" setting.
#  You need to run the program with administrative privileges.
#
# 13. "VDiskAttch(virtual disk file name, drive letter)" means
#  Mount the virtual disk file name to the drive with the drive name
#  to be connected.
#  If you omit the path to the virtual disk file name,
#  it will be loaded from the same folder as the script.
#  If "Compact" is specified as the drive letter,
#  the virtual disk will be optimized.
#  It has been tested with virtual disk files
#  created by Diskpart command, VHD and VHDX.
#  Drive name can be rewritten in the same way as the "PATH" setting.
#  It must be run by a process with administrative privileges.
#
# 14."VDiskDetch(virtual disk file)" means.
#  Unmount the virtual disk file.
#  If you omit the path to the virtual disk file name,
#  it will be read from the same folder as the script.
#  If the path is omitted from the virtual disk file name,
#  it will be read from the same folder as the script.
#  Drive name will be rewritten in the same way as the "PATH" setting.
#  It must be run by a process with administrative privileges.
#
# 15."WaitSec(number of seconds to wait)" means.
#  Stops the process for the number of seconds to wait.
#
# 16."Execute(String)" means
#  Environment variables set immediately before are taken over
#  and the string is executed.
#  Environment variables set thereafter have no effect.
#  Other than applications, strings will be executed according
#  to the associated extensions.
#  Drive name will be rewritten in the same way as the "PATH" setting.
#
# 17."ExecuteArg(String)" means
#  Environment variables set immediately before are taken over
#  and the string is executed.
#  Environment variables set thereafter have no effect.
#  Other than applications, strings will be executed according
#  to the associated extensions.
#  If no argument is specified, the program will be executed
#  with no argument.
#  Drive name will be rewritten in the same way as the "PATH" setting.
#
# 18.Argument, or when executed by dragging and dropping a file
#  Drive name will not be rewritten.
#  It is an argument of "ExecuteArg(string)".
#  If "ExecuteArg(string)" is not used, the string of the
#  first argument is executed at the end of the process.
#  Second and subsequent arguments become the arguments
#  of the first argument.
#
# 19.License
#  Copyright(c) 2021 Aromatibus (https://github.com/Aromatibus)
#  Released under the MIT license.
#  See https://opensource.org/licenses/MIT
#  Translated with www.DeepL.com/Translator (free version)
#

# Run as administrator
RunAdministrator=true

# Create an alias for "Python.exe" to avoid app execution aliases
#RelativeMKLink(a:\Python\PYT, a:\Python\Python.exe)

# Add the folder to the "PATH" environment variable
\VSCode
\VSCode\MinGW\bin
\VSCode\node.js
\VSCode\Git\bin
\VSCode\Git\cmd
\Python
\Python\Scripts

# Add environment variables.
NODE_PATH:=:a:\VSCode\node.js\node_modules\npm\node_modules;a:\VSCode\node.js\node_modules\npm
GIT.PATH:=:a:\VSCode\git\bin\git.exe

# Change the extensions and priorities that can be executed.
# Priority up (CMD), Add (PY)
PATHEXT:=:.CMD;.COM;.EXE;.BAT;.VBS;.VBE;.JS;.JSE;.WSF;.WSH;.MSC;.PY

# Create a shortcut to the desktop
DesktopShortcut(VSCode Portable, a:\VSCode\Code.exe)
DesktopShortcut("", a:\Source)
DesktopShortcut("", https://github.com/Aromatibus)

# Connect a virtual disk and start a file from the virtual disk.
ReferenceDrive = c
vdiskattach(C:\Users\Public\Documents\test.vhd, z)
ReferenceDrive = z
execute(z:\TestEnv.cmd)

# Reset the destination drive to default.
ReferenceDrive =

# Launch NexusFont, no arguments
Execute(a:\Fonts\NexusFont\NexusFont.exe)

# Wait for NexusFont to start.
WaitSec(5)

# Launch VSCode, with arguments
ExecuteArg(a:\VSCode\Code.exe)
*/}.toString().split("\r\n").slice(1,-1).join("\r\n");
  var ws = new ActiveXObject("WScript.Shell");
  var sfo = new ActiveXObject("Scripting.FileSystemObject");
  var MODE = {
    READING: 1,
    WRITING: 2,
    APPENDING: 8,
    CREATE: true,
    NOT_CREATE: false,
    WAIT: true,
    NOT_WAIT: false
  }
  var WINSTYLE = {
    HIDDEN: 0,
    NORMAL: 1,
    MINIMUM: 2,
    MAXIMUM: 3,
    INACTIVE_NORMAL: 4,
    LAST_TIME: 5,
    INACTIVE_MINIMUM: 7
  }
  if (ConfigFileName === "") { return ConfigText; }
  var otf = sfo.OpenTextFile(ConfigFileName, MODE.WRITING, MODE.CREATE);
  otf.Write(ConfigText);
  otf.Close();
  DebugMsg("Created a configuration file. : " + ConfigFileName);
  ws.Run(ConfigFileName, WINSTYLE.MAXIMUM, MODE.NOT_WAIT);
}

WScript.Quit(main());
