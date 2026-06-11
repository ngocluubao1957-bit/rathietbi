@echo off
chcp 65001 >nul 2>&1
setlocal EnableDelayedExpansion

:: ============================================================
::  LAY_THONG_TIN.BAT — Thu thập thông tin phần cứng máy tính
::  Tương thích: Windows 7 / 8 / 10 / 11
::  Kết quả: Ghi ra file JSON rồi mở trang web tự động điền
:: ============================================================

title Lay thong tin may tinh...
color 0A

echo.
echo  =====================================================
echo   CONG CU LAY THONG TIN MAY TINH - CNTT
echo  =====================================================
echo.
echo  Dang thu thap thong tin, vui long cho...
echo.

:: --- Xác định thư mục chạy file ---
set "OUTDIR=%~dp0"
set "OUTFILE=%OUTDIR%thong_tin_may.json"

:: ============================================================
:: 1. HOSTNAME
:: ============================================================
set "HOSTNAME_VAL=%COMPUTERNAME%"

:: ============================================================
:: 2. HỌ TÊN NGƯỜI DÙNG ĐĂNG NHẬP
:: ============================================================
set "USERNAME_VAL=%USERNAME%"

:: ============================================================
:: 3. MANUFACTURER & MODEL
:: ============================================================
for /f "skip=1 tokens=*" %%i in ('wmic computersystem get manufacturer 2^>nul') do (
    if not defined MFR set "MFR=%%i"
)
for /f "skip=1 tokens=*" %%i in ('wmic computersystem get model 2^>nul') do (
    if not defined MDL set "MDL=%%i"
)
:: Fallback dùng csproduct
if not defined MFR (
    for /f "skip=1 tokens=*" %%i in ('wmic csproduct get vendor 2^>nul') do (
        if not defined MFR set "MFR=%%i"
    )
)
if not defined MDL (
    for /f "skip=1 tokens=*" %%i in ('wmic csproduct get name 2^>nul') do (
        if not defined MDL set "MDL=%%i"
    )
)
:: Trim whitespace
for /f "tokens=* delims= " %%a in ("!MFR!") do set "MFR=%%a"
for /f "tokens=* delims= " %%a in ("!MDL!") do set "MDL=%%a"

:: ============================================================
:: 4. SERIAL NUMBER (BIOS)
:: ============================================================
for /f "skip=1 tokens=*" %%i in ('wmic bios get serialnumber 2^>nul') do (
    if not defined SN set "SN=%%i"
)
for /f "tokens=* delims= " %%a in ("!SN!") do set "SN=%%a"
:: Nếu là OEM / trống
if /i "!SN!"=="To be filled by O.E.M." set "SN="
if /i "!SN!"=="Default string" set "SN="
if "!SN!"=="0" set "SN="

:: ============================================================
:: 5. CPU
:: ============================================================
for /f "skip=1 tokens=*" %%i in ('wmic cpu get name 2^>nul') do (
    if not defined CPU set "CPU=%%i"
)
for /f "tokens=* delims= " %%a in ("!CPU!") do set "CPU=%%a"

:: ============================================================
:: 6. RAM (tổng, tính ra GB)
:: ============================================================
set "RAM_MB=0"
for /f "skip=1 tokens=*" %%i in ('wmic computersystem get TotalPhysicalMemory 2^>nul') do (
    if not defined RAM_BYTES set "RAM_BYTES=%%i"
)
:: Chuyển bytes -> GB (chia 1073741824) bằng PowerShell
for /f %%r in ('powershell -nologo -command "try{[math]::Round(%RAM_BYTES%/1GB,1)}catch{0}" 2^>nul') do set "RAM_GB=%%r"
if not defined RAM_GB set "RAM_GB=?"

:: ============================================================
:: 7. Ổ CỨNG — liệt kê tất cả disk
:: ============================================================
set "HDD_LIST="
for /f "skip=1 tokens=1,2,3 delims=," %%a in ('wmic diskdrive get model^,size^,mediatype /format:csv 2^>nul ^| findstr /v "^$"') do (
    set "DM=%%b"
    set "DS=%%c"
    if defined DM (
        for /f %%s in ('powershell -nologo -command "try{[math]::Round(%%DS%%/1GB,0)}catch{0}" 2^>nul') do set "DSG=%%s"
        if defined HDD_LIST (
            set "HDD_LIST=!HDD_LIST! / !DM! (!DSG!GB)"
        ) else (
            set "HDD_LIST=!DM! (!DSG!GB)"
        )
    )
)
:: Fallback đơn giản nếu csv không chạy được
if not defined HDD_LIST (
    for /f "skip=1 tokens=*" %%i in ('wmic diskdrive get model 2^>nul') do (
        if not defined HDD_LIST set "HDD_LIST=%%i"
    )
)
for /f "tokens=* delims= " %%a in ("!HDD_LIST!") do set "HDD_LIST=%%a"

:: ============================================================
:: 8. ĐỊA CHỈ IP & MAC (lấy interface đang active)
:: ============================================================
set "IP_ADDR="
set "MAC_ADDR="
for /f "tokens=2 delims=:" %%i in ('ipconfig 2^>nul ^| findstr /i "IPv4"') do (
    if not defined IP_ADDR (
        for /f "tokens=* delims= " %%a in ("%%i") do set "IP_ADDR=%%a"
    )
)
:: Lấy MAC từ interface có IP khớp
for /f "tokens=1-4" %%a in ('arp -a 2^>nul') do (
    if "%%a"=="!IP_ADDR!" if not defined MAC_ADDR set "MAC_ADDR=%%c"
)
:: Fallback: lấy MAC từ wmic
if not defined MAC_ADDR (
    for /f "skip=1 tokens=*" %%i in ('wmic nic where "NetEnabled=True" get MACAddress 2^>nul') do (
        if not defined MAC_ADDR set "MAC_ADDR=%%i"
    )
    for /f "tokens=* delims= " %%a in ("!MAC_ADDR!") do set "MAC_ADDR=%%a"
)

:: ============================================================
:: 9. HỆ ĐIỀU HÀNH
:: ============================================================
for /f "skip=1 tokens=*" %%i in ('wmic os get caption 2^>nul') do (
    if not defined OS_NAME set "OS_NAME=%%i"
)
for /f "skip=1 tokens=*" %%i in ('wmic os get osarchitecture 2^>nul') do (
    if not defined OS_ARCH set "OS_ARCH=%%i"
)
for /f "tokens=* delims= " %%a in ("!OS_NAME!") do set "OS_NAME=%%a"
for /f "tokens=* delims= " %%a in ("!OS_ARCH!") do set "OS_ARCH=%%a"
set "OS_FULL=!OS_NAME! !OS_ARCH!"

:: ============================================================
:: 10. OFFICE / PHẦN MỀM (kiểm tra registry)
:: ============================================================
set "OFFICE_VER="
:: Office 365 / 2019 / 2016
for /f "tokens=2*" %%a in ('reg query "HKLM\SOFTWARE\Microsoft\Office" /v "" 2^>nul ^| findstr /i "REG_SZ"') do (
    if not defined OFFICE_VER set "OFFICE_VER=%%b"
)
:: Kiểm tra thư mục Office phổ biến
if not defined OFFICE_VER (
    if exist "%ProgramFiles%\Microsoft Office\root\Office16\WINWORD.EXE" set "OFFICE_VER=Microsoft Office 2016/2019/365"
    if exist "%ProgramFiles(x86)%\Microsoft Office\root\Office16\WINWORD.EXE" set "OFFICE_VER=Microsoft Office 2016/2019/365"
    if exist "%ProgramFiles%\Microsoft Office\Office15\WINWORD.EXE" set "OFFICE_VER=Microsoft Office 2013"
    if exist "%ProgramFiles(x86)%\Microsoft Office\Office15\WINWORD.EXE" set "OFFICE_VER=Microsoft Office 2013"
    if exist "%ProgramFiles%\Microsoft Office\Office14\WINWORD.EXE" set "OFFICE_VER=Microsoft Office 2010"
)
if not defined OFFICE_VER set "OFFICE_VER=Khong tim thay"

:: ============================================================
:: 11. LOẠI MÁY: Laptop hay Desktop
:: ============================================================
set "PC_TYPE=Desktop"
for /f "skip=1 tokens=*" %%i in ('wmic systemenclosure get chassistypes 2^>nul') do (
    if not defined CHASSIS set "CHASSIS=%%i"
)
:: 8=Laptop, 9=Laptop, 10=Notebook, 14=Notebook
echo !CHASSIS! | findstr /r "{8}" >nul 2>&1 && set "PC_TYPE=Laptop"
echo !CHASSIS! | findstr /r "{9}" >nul 2>&1 && set "PC_TYPE=Laptop"
echo !CHASSIS! | findstr /r "{10}" >nul 2>&1 && set "PC_TYPE=Laptop"
echo !CHASSIS! | findstr /r "{14}" >nul 2>&1 && set "PC_TYPE=Laptop"

:: ============================================================
:: 12. THỜI GIAN THU THẬP
:: ============================================================
set "TIMESTAMP=%date% %time%"

:: ============================================================
:: ESCAPE FUNCTION (thay nháy đôi)
:: ============================================================
:: Hàm đơn giản: thay " bằng '
call :escape MFR
call :escape MDL
call :escape CPU
call :escape HDD_LIST
call :escape OS_FULL
call :escape OFFICE_VER

:: ============================================================
:: GHI FILE JSON
:: ============================================================
echo { > "!OUTFILE!"
echo   "hostname":    "!HOSTNAME_VAL!", >> "!OUTFILE!"
echo   "username":    "!USERNAME_VAL!", >> "!OUTFILE!"
echo   "hang":        "!MFR!", >> "!OUTFILE!"
echo   "model":       "!MDL!", >> "!OUTFILE!"
echo   "serial":      "!SN!", >> "!OUTFILE!"
echo   "cpu":         "!CPU!", >> "!OUTFILE!"
echo   "ram":         "!RAM_GB!", >> "!OUTFILE!"
echo   "hdd":         "!HDD_LIST!", >> "!OUTFILE!"
echo   "ip":          "!IP_ADDR!", >> "!OUTFILE!"
echo   "mac":         "!MAC_ADDR!", >> "!OUTFILE!"
echo   "os":          "!OS_FULL!", >> "!OUTFILE!"
echo   "office":      "!OFFICE_VER!", >> "!OUTFILE!"
echo   "loai":        "!PC_TYPE!", >> "!OUTFILE!"
echo   "timestamp":   "!TIMESTAMP!" >> "!OUTFILE!"
echo } >> "!OUTFILE!"

:: ============================================================
:: HIỂN THỊ KẾT QUẢ
:: ============================================================
echo.
echo  -------------------------------------------------
echo   DA THU THAP XONG THONG TIN MAY:
echo  -------------------------------------------------
echo   Hostname  : !HOSTNAME_VAL!
echo   User      : !USERNAME_VAL!
echo   Hang/Model: !MFR! !MDL!
echo   Serial    : !SN!
echo   CPU       : !CPU!
echo   RAM       : !RAM_GB! GB
echo   O cung    : !HDD_LIST!
echo   IP        : !IP_ADDR!
echo   MAC       : !MAC_ADDR!
echo   OS        : !OS_FULL!
echo   Office    : !OFFICE_VER!
echo   Loai may  : !PC_TYPE!
echo  -------------------------------------------------
echo   >> Da ghi: !OUTFILE!
echo  -------------------------------------------------
echo.

:: ============================================================
:: MỞ TRANG WEB (truyền URL trang khảo sát của bạn)
:: ============================================================
:: Thay URL bên dưới bằng địa chỉ GitHub Pages của bạn
set "WEB_URL=index.html"

:: Nếu file index.html nằm cùng thư mục → mở local
if exist "!OUTDIR!index.html" (
    echo  Dang mo trang web local...
    start "" "!OUTDIR!index.html"
) else (
    echo  Dang mo trang web tren mang...
    start "" "!WEB_URL!"
)

echo.
echo  Trang web se tu dong doc file: thong_tin_may.json
echo  Kiem tra thong tin roi nhan [Gui phieu].
echo.
echo  Nhan phim bat ky de dong cua so nay...
pause >nul
goto :eof

:: ============================================================
:: HÀM ESCAPE: thay dấu " bằng '
:: ============================================================
:escape
set "val=!%~1!"
set "val=!val:"='!"
set "%~1=!val!"
goto :eof
