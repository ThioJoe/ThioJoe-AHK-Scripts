#Requires AutoHotkey v2.0.19

; Various utility functions. This script doesn't run anything by itself, but is meant to be included in other scripts

class ThioUtils {
; -------------------------------------------------------------------------------

; Gets all the controls of a window as objects from the Windows API, given the window's HWND
; Can be used as a replacement for WinGetControls() which only returns control names, this way you can get the names and HWNDs in one go
;    Return Type: Array of control objects with properties: Class (String), ClassNN (String), ControlID (Int), Hwnd (Int / Pointer)
;    Optional Parameter: getText (Bool) - If true, will also get the text of each control and put it into a Text property
static GetAllControlsAsObjects_ViaWindowsAPI(windowHwnd, getText:=unset) {
    ; ---------------- Local Functions ----------------
    EnumChildProc(hwnd, lParam) {
        controlsArray := ObjFromPtrAddRef(lParam)
        
        ; Get control class
        classNameBuffer := Buffer(256)
        DllCall("GetClassName",
            "Ptr", hwnd,
            "Ptr", classNameBuffer,
            "Int", 256
        )
        className := StrGet(classNameBuffer) ; Convert buffer to string

        classNN := ControlGetClassNN(hwnd)
        
        ; Get control ID
        id := DllCall("GetDlgCtrlID", "Ptr", hwnd)

        controlObject := { Hwnd: hwnd, Class: className, ClassNN: classNN, ControlID: id }

        if (IsSet(getText) and getText) {
            text := ControlGetText(hwnd)
            controlObject.Text := text
        }
        
        ; Add control info to the array
        controlsArray.Push(controlObject)
        
        return true  ; Continue enumeration
    }
    ; ------------------------------------------------

    controlsArray := []

    ; Enumerate child windows
    DllCall("EnumChildWindows",
        "Ptr", windowHwnd,
        "Ptr", CallbackCreate(EnumChildProc, "F", 2),
        "Ptr", ObjPtr(controlsArray)
    )
    
    return controlsArray
}

; Checks if a window has a control with a specific class name. Allows wildcards in the form of an asterisk (*), otherwise exact matching is done
;    Return Type: Bool (True/False)
static CheckWindowHasControlName(hwnd, pattern) {
    try {
        controlsObj := WinGetControls("ahk_id " hwnd)
        
        ; Wildcard pattern matching
        if InStr(pattern, "*") {
            pattern := "^" StrReplace(pattern, "*", ".*") "$"
            for ctrlName in controlsObj {
                if RegExMatch(ctrlName, pattern)
                    return true
            }
        }
        ; Exact matching
        else {
            for ctrlName in controlsObj {
                if InStr(ctrlName, pattern)
                    return true
            }
        }
    }
    return false
}

; Shows a dialog box to enter a directory path and validates it
;    Return Type: String (Path) or empty string if cancelled or invalid
static ShowDirectoryPathEntryBox() {
    path := InputBox("Enter a path to navigate to", "Path", "w300 h100")
    
    ; Check if user cancelled the InputBox
    if (path.Result = "Cancel")
        return ""

    ; Trim whitespace
    trimmedPath := Trim(path.Value)
        
    ; Check if the input is empty
    if (trimmedPath = "")
        return ""

    ; Use Windows API to check if the directory exists. Also works for UNC paths
    callResult := DllCall("Shlwapi\PathIsDirectoryW", "Str", trimmedPath)
    if callResult = 0 {
        ;MsgBox("Invalid path format. Please enter a valid path.")
        return ""
    } else {
        return trimmedPath
    }
}

; Get the handle of the window under the mouse cursor
static GetWindowHwndUnderMouse() {
    MouseGetPos(unset, unset, &WindowhwndOut)
    ;MsgBox("Window Hwnd: " Windowhwnd)
    return WindowhwndOut
}

; Get the class name of the control under the mouse cursor
static GetControlClassUnderMouse() {
    MouseGetPos(unset, unset, unset, &classNN)
    return classNN
}

; Get the Handle/Hwnd of the specific control under the mouse cursor (not the window's handle)
static GetControlUnderMouseHandleID() {
    MouseGetPos(unset, unset, unset, &controlHandleID, 2) ; If controlHandleID is not provided, it will be an empty string
    return controlHandleID
}

; Sets the theme of menus by the process - Adapted from https://www.autohotkey.com/boards/viewtopic.php?style=19&f=82&t=133886#p588184
; Usage: Put this before creating any menus. AllowDark will folow system theme. Seems that once set, the app must restart to change it.
;SetMenuTheme("AllowDark")
static SetContextMenuTheme(appMode:=0) {
    static preferredAppMode       :=  {Default:0, AllowDark:1, ForceDark:2, ForceLight:3, Max:4}
    static uxtheme                :=  dllCall("Kernel32.dll\GetModuleHandle", "Str", "uxtheme", "Ptr")

    if (uxtheme) {
        fnSetPreferredAppMode := dllCall("Kernel32.dll\GetProcAddress", "Ptr", uxtheme, "Ptr", 135, "Ptr")
        fnFlushMenuThemes := dllCall("Kernel32.dll\GetProcAddress", "Ptr", uxtheme, "Ptr", 136, "Ptr")
    } else {
        return -1
    }

    if (preferredAppMode.hasProp(appMode))
        appMode:=preferredAppMode.%appMode%

    if (fnSetPreferredAppMode && fnFlushMenuThemes) { ; Ensure the functions were found
        prev := dllCall(fnSetPreferredAppMode, "Int", appMode)
        dllCall(fnFlushMenuThemes)
        return prev
    } else {
        return -1
    }
}

; Uses Windows API SendMessage to directly send a mouse wheel movement message to a window, instead of using multiple wheel scroll events
;     > This is useful for apps that ignore the Windows scroll speed / scroll line amount settings
; Use positive multiplier for scrolling up and negative for scrolling down. The handle ID can be either for the window or a control inside it
static MouseScrollMultiplied(multiplier, forceWindowHandle:=false, targetHandleID:=unset, mousePosX:=unset, mousePosY:=unset) {
    ; Gets the mouse position and handles under the mouse if not provided
    if !IsSet(targetHandleID) or !IsSet(mousePosX) or !IsSet(mousePosY) {
        MouseGetPos(&mousePosX, &mousePosY, &windowHandleID, &controlHandleID, 2) ; If controlHandleID is not provided, it will be an empty string
        if (forceWindowHandle)
            controlHandleID := ""
        ; Use the control handle if available, since most apps seem to work with that than if only the window handle is used
        targetHandleID := controlHandleID ? controlHandleID : windowHandleID
    } 

    ; 120 is the default delta for one scroll notch in Windows, regardless of mouse setting for number of lines to scroll
    delta := Round(120 * multiplier) 
    ; Construct wParam: shift delta to high-order word (left 16 bits)
    wParam := delta << 16
    ; Construct lParam: combine x and y coordinates: x goes in low word, y in high word
    lParam := (mousePosY << 16) | (mousePosX & 0xFFFF)
    ; WM_MOUSEWHEEL = 0x020A
    result := SendMessage(0x020A, wParam, lParam, unset, "ahk_id " targetHandleID)

    ; Uncomment below for debugging. Sometimes apps return a failed result even if it works, so leaving this commented out since it's not reliable
    ; ---------------------------------------------
    ; using := (IsSet(controlHandleID) && controlHandleID) ? "Control" : ((IsSet(windowHandleID) && windowHandleID) ? "Window" : "Given Parameter")
    ; resultStr := result > 0 ? "Failed (" result ")" : "Success (" result ")"
    ; ToolTip("SendMessage WM_MOUSEWHEEL returned: " resultStr "`n wParam: " wParam "`n Delta: " delta "`n" using " Handle: " targetHandleID)
    ; ---------------------------------------------

    return result
}

; Check if mouse is over a specific window and control. Allows for wildcards in the control name
; Example:          #HotIf mouseOver("ahk_exe dopus.exe", "dopus.tabctrl1")
static CheckMouseOverControlAndWindow(winTitle, ctl := '') {
    MouseGetPos(unset, unset, &hWnd, &classNN)
    if classNN = "" {
        return false
    }

    ; Checks if any window exists with the desired checked title and also the hWnd of that under the mouse
    ; Effectively checking if the window under the mouse matches title of the one passed in
    if WinExist(winTitle ' ahk_id' hWnd) {
        ; Sets up the ctl regex pattern by replacing standard wildcard with regex wildcard
        if (ctl = '')
            ctl := '.*'
        else if InStr(ctl, '*')
            ctl := StrReplace(ctl, '*', '.*')

        ; Match classNN using the wildcard. Will also match exact matches if no wildcard
        if RegExMatch(classNN, '^' ctl '$')
            return true
        else 
            return false
    }
}

; Optimized version for exact control matches:
static CheckMouseOverControlAndWindowExact(winTitle, ctl) {
    MouseGetPos(unset, unset, &hWnd, &classNN)
    if classNN = "" {
        return false
    }

    return WinExist(winTitle " ahk_id" hWnd) && (ctl = classNN)
}

static CheckMouseOverControl(ctl){
    MouseGetPos(unset, unset, unset, &classNN)
    if classNN = "" {
        return false
    }

    if InStr(ctl, '*')
        ctl := StrReplace(ctl, '*', '.*')
    
    ; Match classNN using the wildcard. Will also match exact matches if no wildcard
    if RegExMatch(classNN, '^' ctl '$')
        return true
    else 
        return false
}

static CheckMouseOverControlExact(ctl){
    MouseGetPos(unset, unset, unset, &classNN)
    return (ctl = classNN)
}

; Check if mouse is over a specific window and control (allows wildcards), with additional parameters for various properties of the control
static CheckMouseOverControlAdvanced(winTitle, ctl := '', ctlMinWidth := 0) {
    ; ------ Local Functions ------
    checkWidth(ctrlToCheck, winHwnd, ctlMinWidth) {
        if ctlMinWidth = 0
            return true
        ControlGetPos(&OutX, &OutY, &OutWidth, &OutHeight, ctrlToCheck, winHwnd)
        return OutWidth > ctlMinWidth
    }
    ; ----------------------------
    MouseGetPos(unset, unset, &hWnd, &classNN)
    if classNN = "" {
        return false
    }

    ; Checks if any window exists with the desired checked title and also the hWnd of that under the mouse
    ; Effectively checking if the window under the mouse matches title of the one passed in
    if WinExist(winTitle ' ahk_id' hWnd) {
        ; Sets up the ctl regex pattern by replacing standard wildcard with regex wildcard
        if (ctl = '')
            ctl := '.*'
        else if InStr(ctl, '*')
            ctl := StrReplace(ctl, '*', '.*')

        ; Match classNN using the wildcard
        if RegExMatch(classNN, '^' ctl '$')
            matched := true
        else 
            matched := false
    } else {
        return false
    }

    ; If window and control match, check further criteria
    if (matched){
        if (ctlMinWidth > 0) {
            ctrlHwnd := ControlGetHwnd(classNN, "ahk_id " hWnd) ; Get the control's handle ID
            return (checkWidth(ctrlHwnd, hWnd, ctlMinWidth) = true)
        }
    }
    ; If nothing returned true, return false
    return false
}

; Check if mouse is over a specific window by program name (even if not focused)
; Example:          #HotIf mouseOverProgram("ahk_exe notepad.exe")
static CheckMouseOverProgram(programTitle) {
    MouseGetPos(unset, unset, &hWnd)
    Return WinExist(programTitle " ahk_id" hWnd)
}

; Example: CheckMouseOverSpecificWindowClass("#32770")
static CheckMouseOverSpecificWindowClass(classNNCheck) {
    MouseGetPos(unset, unset, &hWnd)
    ; Get the classNN of the window
    windowClass := WinGetClass("ahk_id " hWnd)
    ; Check if the classNN matches the one provided
    return (windowClass == classNNCheck)
}

; Launch any program and move it to the mouse position, with parameters for relative offset vs mouse position
; Optionally, you can provide the path to the executable to launch (which may be faster, and should be more reliable), otherwise it will use the program title
static LaunchProgramAtMouse(programTitle, xOffset := 0, yOffset := 0, exePath := "", forceWinActivate := false, sizeX:=0, sizeY:=0) {
    ; Store original settings to restore later
    originalMouseMode := A_CoordModeMouse
    originalWinDelay := A_WinDelay
    
    ; Optimize window operations delay
    SetWinDelay(0)
    
    ; Set coordinate mode for consistent positioning
    CoordMode("Mouse", "Screen")
    
    ; Get mouse position once and calculate new position
    MouseGetPos(&mouseX, &mouseY)
    newX := mouseX + xOffset
    newY := mouseY + yOffset
    
    ; Launch program. If path is provided, use it. Otherwise, use the program title
    If (exePath != "")
        Run(exePath)
    Else    
        Run(programTitle)
    
    timeoutAt := A_TickCount + 3000
    
    ; Instead of using WinWait, this uses a timer to check if the window exists which is considerably faster (at least 100ms)
    CheckWindow() {
        if (A_TickCount > timeoutAt) {
            SetTimer(CheckWindow, 0)
            MsgBox("Timed out after 3 seconds")
            return
        }
        
        ; Use direct window title for faster checking
        if WinExist("ahk_exe " programTitle) {
            SetTimer(CheckWindow, 0)
            
            ; Combine move and activate into one operation if possible
            if (forceWinActivate) {
                if (sizeX > 0)
                    WinMove(newX, newY, sizeX, sizeY)
                else
                    WinMove(newX, newY)
                WinActivate()
            } else {
                if (sizeX > 0)
                    WinMove(newX, newY, sizeX, sizeY)
                else
                    WinMove(newX, newY)
            }
        }
    }
    
    ; Use a faster timer interval
    SetTimer(CheckWindow, 5)
    
    ; Restore original settings
    SetWinDelay(originalWinDelay)
    CoordMode("Mouse", originalMouseMode)
}

; Just checks if a particular clipboard format is currently on the clipboard or not
static CheckForClipboardFormat(formatName := "", formatIDInput := unset) {
    formatId := -1

    if IsSet(formatIDInput) {
        formatId := formatIDInput
    } else if IsSet(formatName) {
        formatId := DllCall("RegisterClipboardFormat", "Str", formatName, "UInt")
        if !formatId
            Throw("Failed to register clipboard format: " formatName)
    } else {
        Throw("Error in Checking clipboard format: No format name or ID provided.")
    }

    if (formatId > 0) {
        return DllCall("IsClipboardFormatAvailable", "UInt", formatId, "UInt") ; Returns windows bool (0 or 1)
    } else {
        return false
    }   
}

; Gets the raw bytes data of a specific clipboard format, given the format's name string or ID number
static GetClipboardFormatRawData(formatName := "", formatIDInput := unset) {
    if IsSet(formatIDInput) {
        formatId := formatIDInput
    } else if IsSet(formatName) {
        formatId := DllCall("RegisterClipboardFormat", "Str", formatName, "UInt")
        if !formatId
            Throw("Failed to register clipboard format: " formatName)
    } else {
        Throw("Error in Getting clipboard format data: No format name or ID provided.")
    }

    ; Get all clipboard data
    clipData := ClipboardAll()
    
    if clipData.Size = 0
        return []
    
    ; Create buffer to read from
    bufferObj := Buffer(clipData.Size)
    DllCall("RtlMoveMemory", 
        "Ptr", bufferObj.Ptr, 
        "Ptr", clipData.Ptr, 
        "UPtr", clipData.Size)
    
    offset := 0
    while offset < clipData.Size {
        ; Read format type (4 bytes)
        currentFormat := NumGet(bufferObj, offset, "UInt")
        
        ; If we hit a zero format, we've reached the end
        if currentFormat = 0
            break
            
        ; Read size of this format's data block (4 bytes)
        dataSize := NumGet(bufferObj, offset + 4, "UInt")
        
        ; If this is the format we want
        if currentFormat = formatId {
            ; Create array to hold the bytes
            bytes := []
            bytes.Capacity := dataSize  ; Pre-allocate for better performance
            
            ; Extract each byte directly from the buffer
            loop dataSize {
                bytes.Push(NumGet(bufferObj, offset + 8 + A_Index - 1, "UChar"))
            }
            
            return bytes
        }
        
        ; Move to next format block
        offset += 8 + dataSize  ; Skip format (4) + size (4) + data block
    }
    
    return []  ; Format not found
}

static TooltipWithDelayedRemove(text, delayMs, x := unset, y := unset) {
    if (IsSet(x) && IsSet(y)) {
        ToolTip(text, x, y)
    } else {
        ToolTip(text)
    }
    
    this.RemoveToolTip(delayMs)
}

static RemoveToolTip(delayMs := 0) {
    ; Local function to use in the timer callback
    SetNoTooltip() {
        ToolTip()  ; Calling ToolTip with no parameters removes it
    }

    if delayMs > 0 {
        SetTimer(SetNoTooltip, -1 * delayMs) 
    } else {
        SetNoTooltip()
    }
}

/**
 * Resolves an ms-resource URI for a given AppX package using SHLoadIndirectString.
 * @param packageFamilyName The package family name (e.g., "Microsoft.ScreenSketch_8wekyb3d8bbwe").
 * @param ResourceUri The full ms-resource URI, such a "ms-resource://Microsoft.ScreenSketch/Resources/MarkupAndShareToast" 
 * @returns The cosntructed string that SHLoadIndirectString understands.
 */
static MakeAppxResourceString(packageFamilyName, resourceUri) {
    PackageFullName := this.GetAppxPackageFullName(packageFamilyName)
    ; Construct the special "indirect string" that SHLoadIndirectString understands.
    indirectString := "@{" . PackageFullName . "?" . ResourceUri . "}"
    return indirectString
}

/**
 * Resolves an ms-resource URI for a given AppX package using SHLoadIndirectString.
 * @param dllname The name or path of the DLL (e.g., "NotificationController.dll" or "%SystemRoot%\system32\shell32.dll")
 * @param resourceId The resource ID of the string resource, as a string. Usually this is a number with a negative sign (e.g., "-100").
 * @returns The constructed string that SHLoadIndirectString understands.
 */
static MakeDllResourceString(dllname, resourceId) {
    ; If it already contains slashes, assume it's a full path, so don't prepend the system path.
    if (InStr(dllname, "\") = 0) {
        dllname := "%SystemRoot%\system32\" . dllname
    }
    ; Construct the special "indirect string" that SHLoadIndirectString understands.
    indirectString := "@" . dllname . "," . resourceId
    return indirectString
}

/**
 * Resolves a windows localized (multilanguage) resource string using SHLoadIndirectString API.
 * @param indirectString The indirect string to resolve, formatted as "@{PackageFullName?ms-resource-uri}" or "@%SystemRoot%\system32\shell32.dll,-100".
 * @returns The resolved string if successful, or an error message if not.
 */
static ResolveWindowsResource(indirectString) {
    ; Prepare a buffer to receive the output string.
    outputBuffer := Buffer(4096 * 2, 0)

    ; Call the SHLoadIndirectString function from shlwapi.dll.
    hResult := DllCall("shlwapi\SHLoadIndirectString", "WStr", indirectString, "Ptr", outputBuffer, "UInt", outputBuffer.Size // 2, "Ptr", 0)

    ; Check the result.
    if (hResult = 0) {
        return StrGet(outputBuffer, "UTF-16")
    } else {
        OutputDebug("`nError: Failed to load indirect string. HRESULT: " . Format("0x{:X}", hResult))
        return false
    }
}

/**
 * Gets the first full package name for a given package family name.
 * @param PackageFamilyName The package family name (e.g., "Microsoft.ScreenSketch_8wekyb3d8bbwe").
 * @returns The full package name string, or 'false' if not found or an error occurs.
 */
static GetAppxPackageFullName(PackageFamilyName) {
    ; Constants needed for the API call
    PACKAGE_FILTER_HEAD := 0x00000010
    ERROR_SUCCESS := 0
    ERROR_INSUFFICIENT_BUFFER := 122

    ; First DllCall: Get the required buffer size and number of packages.
    ; We pass 0 for the buffer pointers to signal that we are querying for size.
    hResult := DllCall("kernel32\FindPackagesByPackageFamily",
        "WStr", PackageFamilyName,
        "UInt", PACKAGE_FILTER_HEAD,
        "UIntP", &count := 0,      ; Output: number of packages found
        "Ptr", 0,
        "UIntP", &bufferLen := 0,  ; Output: required buffer length
        "Ptr", 0,
        "Ptr", 0)

    ; The first call is successful if it returns ERROR_SUCCESS (no packages found)
    ; or ERROR_INSUFFICIENT_BUFFER (packages found, sizes returned).
    if (hResult != ERROR_SUCCESS && hResult != ERROR_INSUFFICIENT_BUFFER) {
        OutputDebug("`nError: Could not query for package size. Code: " . hResult)
        return false
    }

    ; If no packages were found, we can exit now.
    if (count = 0) {
        OutputDebug("`nError: No packages found for family name '" . PackageFamilyName . "'")
        return false
    }

    ; Allocate buffers with the sizes we just retrieved.
    fullNamesBuffer := Buffer(A_PtrSize * count)
    stringBuffer := Buffer(bufferLen * 2) ; WCHARs are 2 bytes

    ; Second DllCall: Get the actual package names.
    hResult := DllCall("kernel32\FindPackagesByPackageFamily",
        "WStr", PackageFamilyName,
        "UInt", PACKAGE_FILTER_HEAD,
        "UIntP", &count,
        "Ptr", fullNamesBuffer.Ptr, ; Pointer to an array of string pointers
        "UIntP", &bufferLen,
        "Ptr", stringBuffer.Ptr,    ; Pointer to the buffer for the strings themselves
        "Ptr", 0)

    ; Check the result. ERROR_SUCCESS is 0.
    if (hResult != ERROR_SUCCESS) {
        OutputDebug("`nError: Could not retrieve package names. Code: " . hResult)
        return false
    }

    ; If we found at least one package, retrieve the first full name.
    if (count > 0) {
        ; The fullNamesBuffer now contains an array of pointers. Get the first pointer.
        firstPtr := NumGet(fullNamesBuffer, 0, "Ptr")
        ; Convert that pointer to a string.
        return StrGet(firstPtr)
    }

    OutputDebug("`nError: No packages found for family name '" . PackageFamilyName . "'")
    return false
}

; ------------------------- High precision timer functions --------------------
; Note: Using these as function calls will add significant overhead if measuring small time intervals (Under ~0.1 ms)
static StartTimer() {
    DllCall("QueryPerformanceFrequency", "Int64*", &freq := 0) ; Get the frequency of the counter
    DllCall("QueryPerformanceCounter", "Int64*", &CounterBefore := 0)
    return CounterBefore
}

static EndTimer(CounterBefore, showMsgBox := true) {
    DllCall("QueryPerformanceCounter", "Int64*", &CounterAfter := 0)
    DllCall("QueryPerformanceFrequency", "Int64*", &freq := 0) ; Call this again to avoid having to pass it as a parameter
    if (showMsgBox){
        MsgBox("Elapsed QPC time is " . (CounterAfter - CounterBefore) / freq * 1000 " ms")
    }
    return (CounterAfter - CounterBefore) / freq * 1000
}

; -------------------------------------------------------------------------------
} ; End of ThioUtils class
