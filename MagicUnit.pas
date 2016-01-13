unit MagicUnit;

interface

uses
  SysUtils,dialogs, Classes, Shellapi, Windows, Messages, Graphics, Forms,
  UrlMon, Registry, WinInet, JPEG, ShlObj, ClipBrd, TLHelp32, WinSock, ComObj,
  ActiveX, DateUtils, MMSystem;

Const
  SyTile = 0; { Tile Wallpaper }
  SyCenter = 1; { Center wallpaper }
  SyFit = 2; { Resize the wallpaper to match current resolution }
  SyXY = 3;

{Computer}
  Procedure RunScreenSaver;
  Procedure LogOff;
  Procedure Standby;
  Procedure Restart;
  Procedure ShutDown;
  Function GetComputerName: String;
  function GetCurrentUser : string;
  Procedure MyComputer;
  Procedure Controlpanel;
  Procedure WindowsNoClose(Bool: Boolean);
  function GetDriveSerialNumber(Drive: Char) :String;

{Process and Applications}
  Procedure GetProcessList;
  Function GetAppRun: TStringList;
  Function IsAppRun(FileName: String): Boolean;
  Procedure Execute(FileName: String);
  Procedure KillApp(FileName: String);
  Procedure RunOnStartup(Title, FileName: String; Bool: Boolean);

{Recycle bin}
  Function SendToRecycleBin(Item: String): Boolean;
  Procedure EmbtyRecycleBin;

{Date and Hour}
  Procedure ChangeDate(dt: TDateTime);
  Procedure ChangeHour(dt: TDateTime);
  Function GetTime: String;

{Clipboard}
  Procedure TextToClipbrd(Value: String);
  Function TextFromClipbrd: String;
  procedure BmpToClipboard(Filename: String);
  procedure JpgToClipboard(Filename: String);
  procedure GetClipboardFileNames(FileNames: TStrings);

{Desktop}
  Procedure ShowHideDesktop(Bool: Boolean);
  Procedure ShowHideTaskbar(Bool: Boolean);
  procedure RefreshDesktop;
  procedure ChangerWallpaper(Filename : String; Style, X, Y : Integer);

{Bitmap and Jpg files}
  Procedure JpgToBmp(FileSource, FileDist: String);
  Procedure BmpToJpg(FileSource, FileDist: String; Value: Integer);

{Special directories}
  Function GetCurDir: String;
  Function GetWinDir: String;
  Function GetSysDir: String;
  Function GetTempDir: String;
  function GetTempInternetDir: string;
  function GetCookiesDir: string;
  function GetHistoryDir: string;
  function GetDesktopDir: string;
  function GetProgFilesDir: string;
  function GetMyDocDir: string;
  function GetFavorisDir: string;
  function GetStartUpDir: string;
  function GetRecentDir: string;
  function GetSendToDir: string;
  function GetStartMenuDir: string;
  function GetFontsDir: string;
  function GetMusicDir: string;

{CPU and Memory}
  Function GetCPUUsage: String;
  Function GetFreeMemory: String;

{Files and Folders}
  Function AddSlash(Filename: String): String;
  Function GetDrives: TStringList;
  Function GetFileSize(FileName: String): LongInt;
  Procedure CopyFile(FileSource, FileDist: String);
  function CopyAllFiles(ItemSource, ItemDest: string; Rename: boolean): boolean;
  Function FilesInFolder(Folder: String): TStringList;
  Procedure DeleteFolder(Folder: String);

{Association}
  Procedure AddAssociate(Ext, FileName, Icone: String);
  Procedure DeleteAssociate(Ext: String);
  function GetAssociate(const Ext: string): string;

{Internet}
  Procedure InternetConnect;
  Procedure InternetAutodial;
  Procedure InternetDisconnect;
  function IsConnected: boolean;
  function DownloadFile(FileSource, FileDest: string): Boolean;
  Function IsValidEmail(Email: String): Boolean;
  procedure SendMail(MailTo, Subject, Cc, Bcc: string; Body, Attachment: String);

{Network}
  Function IsValidIP(Ip: string): Boolean;
  Function GetLocalIP: String;
  Function HostToIP(HostName: String): String;
  Function GetHosts(Hostlist: TStrings): Boolean;
  Function GetIpAddresses(IPlist: TStrings): Boolean;

{Millsance}
  Function Crypt(Value:String): String;
  Procedure CaptureScreen(Bmp: TBitmap);
  Procedure EnabledInput(Bool: Boolean);
  procedure CreateShortcut(FileName, Dest, Caption: string);
  
{CD Rom}
  procedure OpenCdRomDrive(Drive: Char; Open: Boolean);
  function IsCdRomDrive(Drive: Char): boolean;

var
  AppList: TStringList;

implementation

Const
  CSIDL_MUSIC = $000d;

var
  FDrvList, FDirCont, DateiList, HandleList, PIDList: TStringList;

Function GetFileNameFromHandle(Handle: hwnd): String;
var
  PID: DWord;
  aSnapShotHandle: THandle;
  ContinueLoop: Boolean;
  aProcessEntry32: TProcessEntry32;
begin
  GetWindowThreadProcessID(Handle, @PID);
  aSnapShotHandle := CreateToolHelp32SnapShot(TH32CS_SNAPPROCESS, 0);
  aProcessEntry32.dwSize := SizeOf(aProcessEntry32);
  ContinueLoop := Process32First(aSnapShotHandle, aProcessEntry32);
  while Integer(ContinueLoop) <> 0 do
    begin
      if aProcessEntry32.th32ProcessID = PID then
        begin
          result := aProcessEntry32.szExeFile;
          break;
        end;
      ContinueLoop := Process32Next(aSnapShotHandle, aProcessEntry32);
    end;
  CloseHandle(aSnapShotHandle);
end;

Function EnumWindowsProc(hWnd: HWND; lParam: LPARAM): Bool;
var
  Capt, Cla: array[0..255] of char;
  Datei: string;
  ident: dword;
  bool: boolean;
  i, j: integer;
begin
  GetWindowText(hWnd, Capt, 255);
  GetClassName(hwnd, Cla, 255);
  Datei := Uppercase(ExtractFileName(GetFileNameFromhandle(hwnd)));
  DateiList.Add(Datei);
  HandleList.Add(IntToStr(HWnd));
  GetWindowThreadProcessId(StrToInt(HandleList[HandleList.Count - 1]), @ident);
  PIDList.Add(IntToStr(ident));
  for i := 0 to dateilist.Count - 1 do
    begin
      bool := false;
      for j := 0 to AppList.Count - 1 do
        if dateilist.Strings[i] = AppList.Strings[j] then
          bool := true;
      if bool = false then
        if Uppercase(ExtractFileName(application.ExeName)) <> dateilist.Strings[i] then
          AppList.Add(dateilist.Strings[i]);
    end;
  Result:=true;
end;

Procedure GetProcessList;
begin
  PIDList.Clear;
  DateiList.Clear;
  HandleList.Clear;
  AppList.Clear;
  EnumWindows(@EnumWindowsProc, 0);
end;

Function GetAppRun: TStringList;
begin
  GetprocessList;
  result := AppList;
end;

Function IsAppRun(FileName: String): Boolean;
begin
  GetProcessList;
  FileName := UpperCase(FileName);
  result := AppList.IndexOf(FileName) <> -1
end;

function KillProcessByPID(PID : DWord): boolean;
var
	myhandle: THandle;
	i: integer;
begin
  myhandle := OpenProcess(PROCESS_TERMINATE, False, PID);
  TerminateProcess(myhandle, 0);
  for i:=0 to 5000 do
    Application.ProcessMessages;
  GetProcessList;
  Result := PIDList.IndexOf(IntToStr(PID)) = -1;
end;

Procedure KillApp(FileName: String);
var
  i: integer;
begin
  for i := 0 to DateiList.Count - 1 do
    if Pos(Filename, DateiList[i]) > 0 then
      break;
  if i < DateiList.Count then
  	if not KillProcessByPID(StrToInt64(PIDList[i])) then
	  	exit;
end;

Function AddSlash(Filename: String): String;
begin
  if length(Filename) > 1 then
  if FileName[length(FileName)] <> '\' then
    FileName := FileName + '\';
  result := FileName;
end;

Function GetComputerName: String;
var
  N: Cardinal;
  Buf: array[0..16] of AnsiChar;
begin
  N := SizeOf(Buf) - 1;
  windows.GetComputerName(Buf, N);
  Result := PChar(@Buf[0]);
end;

Function HostToIP(HostName: String): String;
type
  TaPInAddr = array[0..10] of PInAddr;
  PaPInAddr = ^TaPInAddr;
var
  phe: PHostEnt;
  pptr: PaPInAddr;
  i: Integer;
  GInitData: TWSAData;
begin
  WSAStartup($101, GInitData);
  Result := '';
  phe := GetHostByName(PChar(HostName));
  if phe = nil then Exit;
  pPtr := PaPInAddr(phe^.h_addr_list);
  i := 0;
  while pPtr^[i] <> nil do
    begin
      Result := inet_ntoa(pptr^[i]^);
      Inc(i);
    end;
  WSACleanup;
end;

Function GetLocalIP: String;
type
  TaPInAddr = array[0..10] of PInAddr;
  PaPInAddr = ^TaPInAddr;
var
  phe: PHostEnt;
  pptr: PaPInAddr;
  Buffer: array[0..63] of Char;
  I: Integer;
  GInitData: TWSAData;
begin
  WSAStartup($101, GInitData);
  Result := '';
  GetHostName(Buffer, SizeOf(Buffer));
  phe := GetHostByName(buffer);
  if phe = nil then Exit;
  pPtr := PaPInAddr(phe^.h_addr_list);
  I := 0;
  while pPtr^[I] <> nil do
    begin
      Result := inet_ntoa(pptr^[I]^);
      Inc(I);
    end;
  WSACleanup;
end;

function IsValidIP(Ip: String): Boolean;
var
  z, i: integer;
  st: array[1..3] of byte;
const
  ziff = ['0'..'9'];
begin
  st[1] := 0;
  st[2] := 0;
  st[3] := 0;
  z := 0;
  for i := 1 to Length(ip) do
    if not(ip[i] in ziff) then
      if ip[i] = '.' then
        begin
          Inc(z);
          if z < 4 then
            st[z] := i
          else
            begin
              result := false;
              Exit;
            end;
        end
      else
        begin
          result := false;
          Exit;
        end;
  if (z <> 3) or (st[1] < 2) or (st[3 ] = Length(ip)) or (st[1] + 2 > st[2]) or
    (st[2] + 2 > st[3]) or (st[1] > 4) or (st[2] > st[1] + 4) or
      (st[3] > st[2] + 4) then
        begin
          result := false;
          Exit;
        end;
  z := StrToInt(Copy(ip, 1,st[1] - 1));
  if (z > 255) or (ip[1] = '0') then
    begin
      result := false;
      Exit;
    end;
  z := StrToInt(Copy(ip, st[1] + 1, st[2] - st[1] - 1));
  if(z > 255) or ((z <> 0) and (ip[st[1] + 1] = '0'))then
    begin
      result := false;
      Exit;
    end;
  z := StrToInt(Copy(ip, st[2] + 1, st[3] - st[2] - 1));
  if (z > 255) or ((z <> 0) and (ip[st[2] + 1] = '0'))then
    begin
      result := false;
      Exit;
    end;
  z := StrToInt(Copy(ip, st[3] + 1,Length(ip) - st[3]));
  if(z > 255)or((z <> 0)and(ip[st[3] + 1] = '0'))then
    begin
      result := false;
      Exit;
    end;
  Result := true;
end;

function IsValidEmail(Email: String): boolean;
  function CheckAlloweEd(const s: String): boolean;
  var
    i: Integer;
  begin
  Result := False;
  for i := 1 to Length(s) do
    if not (s[i] in ['a'..'z', 'A'..'Z', '0'..'9', '_', '-', '.']) then Exit;
  Result := true;
  end;
var
  i, len: Integer;
  namePart, serverPart: String;
begin
  Result := False;
  i := Pos('@', Email);
  if (i = 0) or (Pos('..', Email) > 0) then Exit;
  namePart := Copy(Email, 1, i - 1);
  serverPart := Copy(Email, i + 1, Length(Email));
  len := Length(serverPart);
  if (len < 4) or (Pos('.', serverPart) = 0) or (serverPart[1] = '.') or
    (serverPart[len] = '.') or (serverPart[len - 1] = '.') then Exit;
  Result := CheckAlloweEd(namePart) and CheckAlloweEd(serverPart);
end;

Function GetHosts(Hostlist: TStrings): Boolean;
var
  WorkGroupHandle, ComputerHandle: THandle;
  Network: TNetResource;
  Error, BufferLength, I,J , WorkGroupEntries, ComputerEntries: DWORD;
  WorkGroupBuffer, ComputerBuffer: array[1..250] of TNetResource;
begin
  result := false;
  HostList.Clear;
  FillChar(Network, SizeOf(Network), 0);
  with Network do
    begin
      dwScope := RESOURCE_GLOBALNET;
      dwType := RESOURCETYPE_ANY;
      dwUsage := RESOURCEUSAGE_CONTAINER
    end;
  Error := WNetOpenEnum(RESOURCE_GLOBALNET, RESOURCETYPE_ANY, 0, @Network,
    WorkGroupHandle);
  if Error = NO_ERROR then
    begin
      WorkGroupEntries := 250;
      BufferLength := SizeOf(WorkGroupBuffer);
      Error := WNetEnumResource(WorkGroupHandle, WorkGroupEntries,
        @WorkGroupBuffer, BufferLength);
      if Error = NO_ERROR then
        begin
          for I := 1 to WorkGroupEntries do
            begin
              Error := WNetOpenEnum(RESOURCE_GLOBALNET, RESOURCETYPE_ANY, 0,
                @WorkGroupBuffer[I], ComputerHandle);
              if Error = NO_ERROR then
                begin
                  ComputerEntries := 250;
                  BufferLength := SizeOf(ComputerBuffer);
                  Error := WNetEnumResource(ComputerHandle, ComputerEntries,
                  @ComputerBuffer, BufferLength);
                  if Error = NO_ERROR then
                    for J := 1 to ComputerEntries do
                      HostList.Add(Copy(ComputerBuffer[J].lpRemoteName, 3,
                        Length(ComputerBuffer[J].lpRemoteName) - 2));
                  WNetCloseEnum(ComputerHandle);
                end;
            end;
        end;
      WNetCloseEnum(WorkGroupHandle);
    end;
  if HostList.Count > 0 then
    Result := true;
end;

Function GetIpAddresses(IPlist: TStrings): Boolean;
var
  i: byte;
begin
  result := false;
  if GetHosts(IPlist)then
    begin
      for i := 0 to IPlist.Count - 1 do
        IPlist.Strings[i] := HostToIP(IPlist.Strings[i]);
      result := true;
    end;
end;

Procedure RunScreenSaver;
begin
  PostMessage(GetDesktopWindow, WM_SYSCOMMAND, SC_SCREENSAVE, 0);
end;

Procedure Standby;
begin
  SendMessage(0, WM_SYSCOMMAND, SC_MONITORPOWER, 1);
end;

Procedure LogOff;
begin
  ExitWindows(ewx_force, 0);
end;

Procedure Restart;
begin
  ExitWindowsEx(ewx_reboot, 0);
end;

Procedure ShutDown;
begin
  ExitWindowsEx(ewx_shutdown, 0);
end;

Function SendToRecycleBin(Item: String): Boolean;
var
  Sh: TSHFileOpStructA;
  P1: array[byte] of char;
begin
  FillChar(P1, sizeof(P1), 0);
  StrPcopy(P1, ExpandFileName(Item) + #0#0);
  with SH do
    begin
      wnd := 0;
      wFunc := FO_DELETE;
      pFrom := P1;
      pTo := nil;
      fFlags := FOF_ALLOWUNDO or FOF_NOCONFIRMATION;
      fAnyOperationsAborted := false;
      hNameMappings := nil;
    end;
  Result := (ShFileOperation(Sh) = 0);
end;

Procedure EmbtyRecycleBin;
type
  TBin = function(Wnd: HWND; LPCTSTR: PChar; DWORD: Word): Integer; stdcall;
var
  Bin: TBin; LibH: THandle;
begin
  LibH := LoadLibrary(PChar('Shell32.dll'));
  if LibH <> 0 then
    begin
      @Bin := GetProcAddress(LibH, 'SHEmptyRecycleBinA');
      if @Bin <> nil then
      Bin(Application.Handle, '', $00000001 or $00000002 or $00000004);
      FreeLibrary(LibH);
      @Bin := nil;
    end
end;

Procedure DeleteFolder(Folder: String);
var
  i: Integer; Rec: TSearchRec;
begin
  if Folder[Length(Folder)] <> '\'then
    Folder := Folder + '\';
  i := FindFirst(Folder + '*.*', faAnyFile, Rec);
  while i = 0 do
    begin
      if((Rec.Attr and faDirectory) <= 0)then
        SysUtils.DeleteFile(Folder + Rec.Name)
      else if Rec.Name[1] <> '.' then
        begin
          DeleteFolder(Folder + Rec.Name);
          removeDir(Folder + Rec.Name)
        end;
      i := FindNext(Rec);
    end;
  SysUtils.FindClose(Rec);
end;

Procedure Controlpanel;
begin
  execute('control')
end;

Function FilesInFolder(Folder: String): TStringList;
var
  i: Integer; Rec: TSearchRec;
begin
  Folder := AddSlash(Folder);
  FDirCont.Clear;
  i := FindFirst(Folder + '*.*', faAnyFile, Rec);
  while i = 0 do
    begin
      if((Rec.Attr and faDirectory) > 0)then
        FDirCont.Add(Rec.Name);
      i := FindNext(Rec);
  end;
  sysutils.FindClose(Rec);
  i := FindFirst(Folder + '*.*', faAnyFile, Rec);
  while i=0 do
    begin
      if((Rec.Attr and faDirectory) <= 0)then
      FDirCont.Add(Rec.Name);
      i := FindNext(Rec);
    end;
  sysutils.FindClose(Rec);
  result := FDirCont;
end;

Procedure Execute(FileName: String);
begin
  shellexecute(0, 'open', pchar(filename), nil, nil, sw_show);
end;

Function GetDrives: TStringList;
var
  i: integer;
  Filename: String;
begin
  FDrvlist.Clear;
  for i := 0 to 25 do
    begin
      Filename := (Char(i + ord('A')) + ':\');
      if GetDriveType(pChar(Filename)) <> 1 then
        FDrvList.Add(Filename)
    end;
  Result := FDrvList;
end;

Function GetCurDir: String;
begin
  Result := AddSlash(GetCurrentDir);
end;

Function GetWinDir: String;
var
  Dir: array[0..MAX_PATH - 1] of char;
begin
  SetString(Result, Dir, GetWindowsDirectory(Dir, MAX_PATH));
  Result := AddSlash(Result);
end;

Function GetSysDir: String;
var
  Dir: array[0..MAX_PATH - 1] of char;
begin
  SetString(Result, Dir, GetSystemDirectory(Dir, MAX_PATH));
  Result := AddSlash(Result);
end;

Function GetTempDir: String;
var
  Dir: array[0..MAX_PATH] of Char;
begin
  if GetTempPath(SizeOf(Dir), Dir) <> 0 then
    Result := StrPas(Dir);
  Result := AddSlash(Result);
end;

function GetTime: String;
Begin
  Result := TimeToStr(GetTickCount / 86400000);
End;

Procedure CaptureScreen(Bmp: TBitmap);
begin
  with bmp do
    begin
      Width := screen.Width;
      Height := screen.Height;
      BitBlt(Canvas.Handle, 0, 0, Width, Height, GetDC(GetDesktopWindow),
        0, 0, SrcCopy);
    end;
end;

procedure ChangerWallpaper(Filename : String; Style, X, Y : Integer);
var
 reg : TRegIniFile;
 Tile : String;
 PosX, PosY : Integer;
begin
  Tile := '0';
  PosX := 0;
  PosY := 0;
  case Style of
    SyTile : Tile := '1';
    SyCenter : Style := SyTile;
    SyXY :
      Begin
        Style := SyTile;
        PosX := X;
        PosY := Y;
      end;
   end;
  reg := TRegIniFile.Create('Control Panel\Desktop');
  with reg do
    try
      WriteString ('', 'Wallpaper', Filename);
      WriteString ('', 'TileWallpaper', Tile);
      WriteString ('', 'WallpaperStyle', IntToStr (Style));
      WriteString ('', 'WallpaperOriginX', IntToStr (PosX));
      WriteString ('', 'WallpaperOriginY', IntToStr (PosY));
    finally
      Free;
    end;
 SystemParametersInfo(SPI_SETDESKWALLPAPER, 0, PChar(Filename),
  SPIF_UPDATEINIFILE Or SPIF_SENDWININICHANGE);
end;

Function GetFileSize(FileName: String): LongInt;
var
  Search: TSearchRec;
begin
  If FindFirst(ExpandFileName(FileName), faAnyFile, Search) = 0 Then
  Result := Search.Size else result := 0
end;

Procedure ShowHideDesktop(Bool: Boolean);
begin
  if bool then
    showwindow(findwindow('progman', nil), sw_show)
  else
    showwindow(findwindow('progman', nil), sw_hide)
end;

Procedure ShowHideTaskbar(Bool: Boolean);
begin
  if bool then
    showwindow(findwindow('shell_traywnd', nil), sw_show)
  else
    showwindow(findwindow('shell_traywnd', nil), sw_hide)
end;

Procedure AddAssociate(Ext, FileName, Icone: String);
  procedure Ecrit(clef, nom, valeur: string);
  var
    Reg: TRegistry;
  begin
    with reg do
      try
        Reg := TRegistry.Create;
        RootKey := HKEY_CLASSES_ROOT;
        OpenKey(clef, True);
        WriteString(nom, valeur);
        CloseKey;
        Free;
      except
      end;
  end;
begin
  Ecrit('.' + ext, '', ext + '_file');
  Ecrit(ext + '_file\shell\open\command', '', FileName + ' %1');
  Ecrit(ext + '_file\DefaultIcon', '', icone);
end;

Procedure DeleteAssociate(Ext: String);
var
  Reg: TRegistry;
begin
  with reg do
    try
      Reg := TRegistry.Create;
      RootKey := HKEY_CLASSES_ROOT;
      DeleteKey('.' + ext);
      Free;
      Reg := TRegistry.Create;
      RootKey := HKEY_CLASSES_ROOT;
      DeleteKey(ext + '_file');
      CloseKey;
      Free;
    except
    end;
end;

function GetAssociate(const Ext: string): string;
var
  FileClass: string;
  Reg: TRegistry;
begin
  Result := '';
  Reg := TRegistry.Create(KEY_EXECUTE);
  Reg.RootKey := HKEY_CLASSES_ROOT;
  FileClass := '';
  if Reg.OpenKeyReadOnly(Ext) then
  begin
    FileClass := Reg.ReadString('');
    Reg.CloseKey;
  end;
  if FileClass <> '' then begin
    if Reg.OpenKeyReadOnly(FileClass + '\Shell\Open\Command') then
    begin
      Result := Reg.ReadString('');
      Reg.CloseKey;
    end;
  end;
  Reg.Free;
end;

Procedure InternetAutodial;
begin
  WinInet.InternetAutodial(2, 0);
end;

Procedure InternetConnect;
begin
  WinInet.InternetAutodial(0, 0);
end;

Procedure InternetDisconnect;
begin
  InternetAutodialHangup(0);
end;

Procedure EnabledInput(Bool: Boolean);
  function FunctionDetect (var LibPointer: Pointer): boolean;
  var
    LibHandle: tHandle;
  begin
   Result := false;
   LibPointer := NIL;
    if LoadLibrary(PChar('USER32.DLL')) = 0 then exit;
    LibHandle := GetModuleHandle(PChar('USER32.DLL'));
    if LibHandle <> 0 then
     begin
      LibPointer := GetProcAddress(LibHandle, PChar('BlockInput'));
      if LibPointer <> NIL then Result := true;
     end;
  end;
 var xBlockInput : function (Block: Boolean): Boolean; stdcall;
begin
 if FunctionDetect (@xBlockInput) then
  xBlockInput (not bool);
end;

Procedure BmpToJpg(FileSource, FileDist: String; Value: Integer);
var
  JPEG: TJPEGImage;
  BMP: TBitmap;
begin
  if(value < 0) or(value > 100) then value := 50;
  BMP := TBitmap.Create;
  try
    BMP.LoadFromFile(FileSource);
    JPEG := TJPEGImage.Create;
    try
      JPEG.CompressionQuality := Value;
      JPEG.Assign(BMP);
      JPEG.SaveToFile(FileDist);
    finally
      JPEG.Free
    end;
  finally
    BMP.Free
  end;
end;

Procedure JpgToBmp(FileSource, FileDist: String);
var
  JPEG: TJPEGImage;
  BMP: TBitmap;
begin
  JPEG := TJPEGImage.Create;
  try
    JPEG.LoadFromFile(FileSource);
    BMP := TBitmap.Create;
    try
      BMP.Width := JPEG.Width;
      BMP.Height := JPEG.Height;
      BMP.Canvas.Draw(0, 0, JPEG);
      BMP.SaveToFile(FileDist);
    finally
      BMP.Free
    end;
  finally
    JPEG.Free
  end;
end;

Procedure CopyFile(FileSource, FileDist: String);
var
  S, T : TFileStream;
begin
  S := TFileStream.Create(FileSource, fmOpenRead);
  try
    T := TFileStream.Create(FileDist, fmOpenWrite or fmCreate);
    try
      T.CopyFrom(S, S.Size);
    finally
      T.Free
    end;
  finally
    S.Free
  end;
end;

function CopyAllFiles(ItemSource, ItemDest: string; Rename: boolean): boolean;
var
  F: TShFileOpStruct;
begin
  FillChar(F, SizeOf(F), #0);
  with F do
  try
    Wnd := 0;
    wFunc := FO_COPY;
    pFrom := PChar(ItemSource + #0);
    pTo := PChar(ItemDest + #0);

    if Rename then
      F.fFlags := FOF_RENAMEONCOLLISION or FOF_SIMPLEPROGRESS
    else
      F.fFlags := FOF_SIMPLEPROGRESS;

    fAnyOperationsAborted := False;
    hNameMappings := nil;
    Result := (ShFileOperation(F) = 0);
  finally
  end;
end;

Procedure ChangeHour(dt: TDateTime);
var
  SysTm : TSystemTime;
begin
  GetSystemTime(SysTm);
  with SysTm do
    begin
      wHour := HourOf(dt);
      wMinute := MinuteOf(dt);
      wSecond := SecondOF(dt);
    end;
  SetSystemTime(SysTm);
  SendMessage(FindWindow('Shell_TrayWnd', nil), WM_TIMECHANGE, 0, 0);
end;

Procedure ChangeDate(dt: TDateTime);
var
  SysTm : TSystemTime;
begin
  GetSystemTime(SysTm);
  with SysTm do
    begin
      wYear := YearOf(dt);
      wMonth := MonthOf(dt);
      wDay := DayOf(dt);
    end;
  SetSystemTime(SysTm);
  SendMessage(FindWindow('Shell_TrayWnd', nil), WM_TIMECHANGE, 0, 0);
end;

Procedure RunOnStartup(Title, FileName: String; Bool: Boolean);
var
  Reg: TRegistry;
begin
  Reg := TRegistry.Create;
  try
    Reg.RootKey := HKEY_LOCAL_MACHINE;
    if Reg.OpenKey('\Software\Microsoft\Windows\CurrentVersion\Run', FALSE) then
      case bool of
        FALSE: Reg.DeleteValue(Title);
        TRUE: Reg.WriteString(Title, FileName)
      end;
  finally
    Reg.Free
  end;
end;

Procedure WindowsNoClose(Bool: Boolean);
var
  Reg: TRegistry;
  c: byte;
begin
  Reg := TRegistry.Create;
  c := 01;
  try
    Reg.RootKey := HKEY_CURRENT_USER;
    if Reg.OpenKey('\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer',FALSE) then
      case bool of
        FALSE: Reg.DeleteValue('NoClose');
        TRUE: Reg.WriteBinaryData('NoClose', c, c);
      end;
  finally
    Reg.Free
  end;
end;

function GetDriveSerialNumber(Drive: Char) :String;
var
  Serial, DirLen, Flags: DWord;
  DLabel : Array[0..11] of Char;
begin
  Try
    GetVolumeInformation(PChar(Drive), dLabel, 12, @Serial, DirLen, Flags, nil, 0);
    Result := IntToHex(Serial, 8);
  Except
    Result :='';
  end;
end;

Procedure MyComputer;
var
  MyItemIDList: PItemIDList;
  MyShellEx: TShellExecuteInfo;
begin
  SHGetSpecialFolderLocation(0, CSIDL_DRIVES, MyItemIDList);
  with MyShellEx do
    begin
      cbSize := Sizeof(MyShellEx);
      fMask := SEE_MASK_IDLIST;
      Wnd := 0;
      lpVerb := nil;
      lpFile := nil;
      lpParameters := nil;
      lpDirectory := nil;
      nShow := SW_SHOW;
      hInstApp := 0;
      lpIDList := MyItemIDList;
    end;
  ShellExecuteEx(@MyShellEx);
end;

Function GetFreeMemory: String;
var
  MemSta:TMemoryStatus;
begin
  GlobalMemoryStatus(MemSta);
  result := format('%%%.1f',[100. * MemSta.dwAvailPhys / MemSta.dwTotalPhys]);
end;

Function GetCPUUsage: String;
var
  data: longint;
begin
  with TRegistry.Create do
    try
      RootKey := HKEY_DYN_DATA;
      OpenKey('PerfStats\StatData', false);
      ReadBinaryData('KERNEL\CPUUsage', data, GetDataSize('KERNEL\CPUUsage'));
    finally
      free
    end;
  result := '%' + inttostr(data)
end;

Function Crypt(Value: String): String;
var
  Index: integer;
begin
  Result := Value;
  for Index := 1 to Length(Value) do
    Result[Index] := chr(not(ord(Value[Index])));
end;

Procedure TextToClipbrd(Value: String);
var
  A: array[0..255]of char;
begin
  StrPCopy(A, Value);
  Clipboard.SetTextBuf(A);
end;

Function TextFromClipbrd: String;
begin
  if Clipboard.HasFormat(CF_TEXT)then
    result := Clipboard.AsText
end;

procedure BmpToClipboard(Filename: String);
var
  MyFormat : Word;
  Bitmap : TBitMap;
  AData: Cardinal;
  APalette : HPALETTE;
begin
  Bitmap := TBitmap.Create;
  with Bitmap do
  try
    LoadFromFile(Filename);
    SaveToClipBoardFormat(MyFormat, AData, APalette);
    ClipBoard.SetAsHandle(MyFormat, AData);
  finally
    Free;
  end;
end;

procedure JpgToClipboard(Filename: String);
var
  MyFormat : Word;
  JPEG : TJPEGImage;
  AData: Cardinal;
  APalette : HPALETTE;
begin
  JPEG := TJPEGImage.Create;
  with JPEG do
  try
    LoadFromFile(Filename);
    SaveToClipBoardFormat(MyFormat, AData, APalette);
    ClipBoard.SetAsHandle(MyFormat, AData);
  finally
    Free;
  end;
end;

procedure GetClipboardFileNames(FileNames: TStrings);
var
  f: THandle;
  buffer: array [0..MAX_PATH] of Char;
  i, numFiles: Integer;
begin
  with Clipboard do
    begin
      if not HasFormat(CF_HDROP) then Exit;
      Open;
      try
        f := GetAsHandle(CF_HDROP);
        if f <> 0 then
          begin
            FileNames.Clear;
            numFiles := DragQueryFile(f, $FFFFFFFF, nil, 0);
            for i := 0 to numfiles - 1 do
              begin
                buffer[0] := #0;
                DragQueryFile(f, i, buffer, SizeOf(buffer));
                FileNames.Add(buffer);
              end;
          end;
      finally
        Close;
      end;
    end;
end;

function GetSpecialFolder(FolderID : longint) : string;
var
  Path : pchar;
  idList : PItemIDList;
begin
  GetMem(Path, MAX_PATH);
  SHGetSpecialFolderLocation(0, FolderID, idList);
  SHGetPathFromIDList(idList, Path);
  Result := string(Path);
  FreeMem(Path);
end;

function GetTempInternetDir: string;
begin
  Result := AddSlash(GetSpecialFolder(CSIDL_INTERNET_CACHE));
end;

function GetCookiesDir: string;
begin
  Result := AddSlash(GetSpecialFolder(CSIDL_COOKIES));
end;

function GetHistoryDir: string;
begin
  Result := AddSlash(GetSpecialFolder(CSIDL_HISTORY));
end;

function GetDesktop: string;
begin
  Result := AddSlash(GetSpecialFolder(CSIDL_DESKTOP));
end;

function GetDesktopDir: string;
begin
  Result := AddSlash(GetSpecialFolder(CSIDL_DESKTOPDIRECTORY));
end;

function GetProgFilesDir: string;
begin
  Result := AddSlash(GetSpecialFolder(CSIDL_PROGRAMS));
end;

function GetMyDocDir: string;
begin
  Result := AddSlash(GetSpecialFolder(CSIDL_PERSONAL));
end;

function GetFavorisDir: string;
begin
  Result := AddSlash(GetSpecialFolder(CSIDL_FAVORITES));
end;

function GetStartUpDir: string;
begin
  Result := AddSlash(GetSpecialFolder(CSIDL_STARTUP));
end;

function GetRecentDir: string;
begin
  Result := AddSlash(GetSpecialFolder(CSIDL_RECENT));
end;

function GetSendToDir: string;
begin
  Result := AddSlash(GetSpecialFolder(CSIDL_SENDTO));
end;

function GetStartMenuDir: string;
begin
  Result := AddSlash(GetSpecialFolder(CSIDL_STARTMENU));
end;

function GetFontsDir: string;
begin
  Result := AddSlash(GetSpecialFolder(CSIDL_FONTS));
end;

function GetMusicDir: string;
begin
  Result := AddSlash(GetSpecialFolder(CSIDL_MUSIC));
end;

procedure RefreshDesktop;
begin
  SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, nil, nil);
end;

procedure CreateShortcut(Filename, Dest, Caption: string);
var
  IObject: IUnknown;
  ISLink: IShellLink;
  IPFile: IPersistFile;
begin
  IObject := CreateComObject(CLSID_ShellLink);
  ISLink  := IObject as IShellLink;
  IPFile  := IObject as IPersistFile;
  with ISLink do begin
    SetPath(pChar(Dest));
    SetWorkingDirectory(pChar(ExtractFilePath(Dest)));
  end;
  IPFile.Save(PWChar(AddSlash(Dest) + Caption + '.lnk'), false);
end;

procedure CreFicLien (FicSource, FicRaccourci :string) ;
var
  ShellLink : IShellLink ;
begin
  if UpperCase (extractFileExt (FicRaccourci)) <> '.LNK' then FicRaccourci := FicRaccourci + '.lnk' ;

  ShellLink := CreateComObject(CLSID_ShellLink) as IShellLink ;
  ShellLink.SetDescription ('') ;
  ShellLink.SetPath (PChar (FicSource)) ;
  ShellLink.SetShowCmd (SW_SHOW) ;
  (ShellLink as IpersistFile).Save (StringToOleStr (FicRaccourci), true) ;

end;

function IsConnected: boolean;
var
  reg: TRegistry;
  buff: dword;
begin
  result := False;
  reg := TRegistry.Create;
  with reg do
    try
      RootKey := HKEY_LOCAL_MACHINE;
      if OpenKey('\System\CurrentControlSet\Services\RemoteAccess', false) then
        begin
          ReadBinaryData('Remote Connection', buff, sizeof(buff));
          result := buff = 1;
          CloseKey;
        end;
    finally
      Free;
    end;
end;

function DownloadFile(FileSource, FileDest: string): Boolean;
begin
  try
    Result :=  UrlDownloadToFile(nil, PChar(FileSource), PChar(FileDest), 0, nil) = 0;
  except
    Result := False;
  end;
end;

procedure OpenCdRomDrive(Drive: Char; Open: Boolean);
begin
  case Open of
    True: mciSendString(Pchar('set cdaudio!' + Drive + ' door open wait'), nil, 0, 0);
    False: mciSendString(Pchar('set cdaudio!' + Drive + ' door closed wait'), nil, 0, 0);
  end;
end;

function IsCdRomDrive(Drive: Char): boolean;
begin
  Result := GetDriveType(PChar(Drive + ':\')) = DRIVE_CDROM;
end;

procedure SendMail(MailTo, Subject, Cc, Bcc: string; Body, Attachment: String);
var
  msg: string;
begin
 msg:=
 'mailto:' + mailTo +
 '?Subject=' + Subject +
 {'&cc=' + CC +
 '&bcc=' + BCC +  }
 '&body=' + body ;
 //'&attachment=' + attachment ;
 ShellExecute(0, nil, PChar(msg), nil, nil, SW_SHOW);
end;

function GetCurrentUser : string;
var
  buffer : String;
  buffsize : DWORD;
begin
  buffsize := 128;
  Setlength(buffer,buffsize);
  Getusername(Pchar(buffer),buffsize);
  result := buffer;
end;

initialization
  AppList := TStringList.Create;
  FDrvlist := TStringList.Create;
  FDirCont := TStringList.Create;
  DateiList := TStringList.Create;
  HandleList := TStringList.Create;
  PIDList := TStringList.Create;

finalization
  AppList.Free;
  FDrvList.Free;
  FDirCont.Free;
  DateiList.Free;
  HandleList.Free;
  PIDList.Free;
end.
