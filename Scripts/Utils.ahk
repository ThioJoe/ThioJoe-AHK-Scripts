#Requires AutoHotkey v2.0.19

; Various utility functions. This script doesn't run anything by itself, but is meant to be included in other scripts

class ThioUtils {
; -------------------------------------------------------------------------------

/**
 * Windows RECT structure for representing rectangular areas.
 * @property {Int32} left Left coordinate
 * @property {Int32} top Top coordinate  
 * @property {Int32} right Right coordinate
 * @property {Int32} bottom Bottom coordinate
 */
class RECT {
    left: i32
    top: i32
    right: i32
    bottom: i32
}

class ControlInfo {
    Hwnd: iptr
    Class := ""
    ClassNN := ""
    ControlID: iptr
    Text := ""  ; Optional, only if getText is true
}

; Type Definition Notes. For some reason need to use 'i32' instead of 'Integer'

;; ------------------- General --------------------

/**
 * Check if a string starts with a specific prefix.
 * @param Haystack 
 * @param Needle 
 * @param {String|1|0} CaseSense 
 * @returns {Boolean} 
 */
static StartsWith(Haystack, Needle, CaseSense := "Off") {
    return (InStr(Haystack, Needle, CaseSense) = 1)
}

/**
 * 
 * @param {string} prompt 
 * @param {string} title 
 * @param {Number} width 
 * @param {Number} height 
 */
static GetUserInput(prompt, title, width := 300, height := 200) {
    ; Pop up a message box with input
    inputObj := InputBox(prompt, title, "w" . width . " h" . height)
    if (inputObj.Result == "Cancel")
        return ""

    input := inputObj.Value

    return input
}


static JoinSelectedLines(addSpace := true) {
    ; Store the original clipboard contents in a variable.
    ;    We save both ClipboardAll (for all data types) and the plain text version.
    local OriginalFullClipboard := ClipboardAll()
    local originalClipboardText := A_Clipboard

    Send("{Blind}{Ctrl Up}{Shift Up}") ; Ensure Ctrl and Shift are released

    try {
        ; Clear the clipboard and copy the currently selected text.
        A_Clipboard := ""
        Send("^x")
        ClipWait(1) ; Wait a maximum of 0.5 seconds for the copy to complete.
        ; Now the clipboard contains the copied text.
    } catch {
        A_Clipboard := OriginalFullClipboard
        return
    }

    ; Get the copied text and replace all newline characters with a space.
    local clipboardText := A_Clipboard

    ; Split the newlines and trim each line
    local prepText := StrReplace(clipboardText, "`r`n", "`n") ; Normalize newlines to `n
    prepText := StrReplace(prepText, "`r", "`n") ; Normalize any remaining `r to `n
    local clipLinesArray := StrSplit(prepText, "`n")
    for index, line in clipLinesArray {
        clipLinesArray[index] := Trim(line)
    }

    ; Join the lines with space or nothing based on parameter
    local joiner := ""
    if (addSpace = true) {
        joiner := " "
    }

    local joinedText := ""
    loop clipLinesArray.Length {
        if (A_Index = 1) { ; Special case for first line to avoid leading joiner
            joinedText := clipLinesArray[A_Index]
        } else {
            joinedText := joinedText . joiner . clipLinesArray[A_Index]
        }
    }

    ; Type the modified text, replacing the original selection.
    ; SendInput(joinedText)
    A_Clipboard := "" ; Clea so we can use Clipwait again
    A_Clipboard := joinedText
    ClipWait(1)
    Send("^v")

    ; Restore the original clipboard. We don't need it anymore.
    A_Clipboard := OriginalFullClipboard
}

;; ------------------- Mouse and Cursor Related --------------------

/**
 * Check if mouse is over a specific window and control. Allows for wildcards in the control name.
 * @param {String} winTitle The window title or identifier to check against (e.g., "ahk_exe dopus.exe")
 * @param {String} ctl The control class name to match, supports wildcards with * (default: '' matches any control)
 * @returns {Bool} True if mouse is over the specified window and control, false otherwise
 * @example #HotIf CheckMouseOverControlAndWindow("ahk_exe dopus.exe", "dopus.tabctrl1")
 */
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

/**
 * Optimized version for exact control matches (no wildcards).
 * @param {String} winTitle The window title or identifier to check against
 * @param {String} ctl The exact control class name to match
 * @returns {Bool} True if mouse is over the specified window and exact control, false otherwise
 */
static CheckMouseOverControlAndWindowExact(winTitle, ctl) {
    MouseGetPos(unset, unset, &hWnd, &classNN)
    if classNN = "" {
        return false
    }

    return WinExist(winTitle " ahk_id" hWnd) && (ctl = classNN)
}

/**
 * Check if mouse is over a control with a specific class name, regardless of window.
 * @param {String} ctl The control class name to match, supports wildcards with *
 * @returns {Bool} True if mouse is over the specified control, false otherwise
 */
static CheckMouseOverControl(ctl) {
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

/**
 * Check if mouse is over a control with an exact class name match (no wildcards).
 * @param {String} ctl The exact control class name to match
 * @returns {Bool} True if mouse is over the exact control, false otherwise
 */
static CheckMouseOverControlExact(ctl) {
    MouseGetPos(unset, unset, unset, &classNN)
    return (ctl = classNN)
}

/**
 * Check if mouse is over a specific window and control with additional property constraints.
 * @param {String} winTitle The window title or identifier to check against (default: '' matches any window)
 * @param {String} ctl The control class name to match, supports wildcards with * (default: '' matches any control)  
 * @param {Integer} ctlMinWidth Minimum width the control must have in pixels (default: 0 for no constraint)
 * @returns {Bool} True if mouse is over the specified window, control, and meets constraints, false otherwise
 */
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
    if (matched) {
        if (ctlMinWidth > 0) {
            ctrlHwnd := ControlGetHwnd(classNN, "ahk_id " hWnd) ; Get the control's handle ID
            return (checkWidth(ctrlHwnd, hWnd, ctlMinWidth) = true)
        }
    }
    ; If nothing returned true, return false
    return false
}

/**
 * Check if mouse is over a specific window by program name (even if not focused).
 * @param {String} programTitle The program identifier (e.g., "ahk_exe notepad.exe")
 * @returns {Bool} True if mouse is over the specified program window, false otherwise
 * @example ' #HotIf CheckMouseOverProgram("ahk_exe notepad.exe") '
 */
static CheckMouseOverProgram(programTitle) {
    MouseGetPos(unset, unset, &hWnd, unset, unset)
    return WinExist(programTitle " ahk_id" hWnd)
}

/**
 * Check if mouse is over a window with a specific window class.
 * @param {String} classNNCheck The window class name to check for (e.g., "#32770")
 * @returns {Bool} True if mouse is over a window with the specified class, false otherwise
 * @example CheckMouseOverSpecificWindowClass("#32770")
 */
static CheckMouseOverSpecificWindowClass(classNNCheck) {
    MouseGetPos(unset, unset, &hWnd, unset, unset)
    ; Get the classNN of the window
    windowClass := WinGetClass("ahk_id " hWnd)
    ; Check if the classNN matches the one provided
    return (windowClass == classNNCheck)
}

/**
 * Get the handle (HWND) of the window under the mouse cursor.
 * @returns {Int} The window handle under the mouse cursor
 */
static GetWindowHwndUnderMouse() {
    MouseGetPos(unset, unset, &WindowhwndOut, unset, unset)
    ;MsgBox("Window Hwnd: " Windowhwnd)
    return WindowhwndOut
}

/**
 * Get the class name of the control under the mouse cursor.
 * @returns {String} The control class name under the mouse cursor
 */
static GetControlClassUnderMouse() {
    MouseGetPos(unset, unset, unset, &classNN, unset)
    return classNN
}

/**
 * Get the handle (HWND) of the specific control under the mouse cursor (not the window's handle).
 * @returns {Int} The control handle under the mouse cursor, or empty string if none
 */
static GetControlUnderMouseHandleID() {
    MouseGetPos(unset, unset, unset, &controlHandleID, 2) ; If controlHandleID is not provided, it will be an empty string
    return controlHandleID
}

;; ------------------ Prompts -----------------------

/**
 * Shows a dialog box to enter a directory path and validates it.
 * @returns {String} Valid directory path, or empty string if cancelled or invalid
 */
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

;; ------------------ Manipulate User Input --------------------

/**
 * Uses Windows API SendMessage to directly send a mouse wheel movement message to a window. Instead o fusing multiple wheel scroll events.
 * This is useful for apps that ignore the Windows scroll speed / scroll line amount settings.
 * @param {Number} multiplier Scroll multiplier - positive for scrolling up, negative for scrolling down
 * @param {Bool} forceWindowHandle Whether to force using window handle instead of control handle (default: false)
 * @param {Bool} useSendMessage Whether to use SendMessage instead of PostMessage (default: false)
 * @param {Integer} targetHandleID Specific window/control handle to target (default: unset, auto-detected)
 * @param {Integer} mousePosX Mouse X position for the message (default: unset, auto-detected)
 * @param {Integer} mousePosY Mouse Y position for the message (default: unset, auto-detected)
 * @returns {Int|unset} Result from SendMessage if used, otherwise unset
 */
static MouseScrollMultiplied(multiplier, forceWindowHandle := false, useSendMessage := false, targetHandleID := unset,  mousePosX := unset, mousePosY := unset) {
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
    wParam := delta << 16 ; Construct wParam: shift delta to high-order word (left 16 bits)
    lParam := (mousePosY << 16) | (mousePosX & 0xFFFF) ; Construct lParam: combine x and y coordinates: x goes in low word, y in high word
    WM_MOUSEWHEEL := 0x020A

    ; PostMessage is safer because it doesn't require waiting for a response from the window like if it freezes
    ; SendMessage might be faster but could be a bit glitchy. I set the timeout to -1 so it shouldn't wait but not sure how well that works.
    if (useSendMessage == true){
        try { 
            result := SendMessage(WM_MOUSEWHEEL, wParam, lParam, unset, "ahk_id " targetHandleID, unset, unset, unset, 25) 
        } catch {
        }   ; No catch, it will always 'fail' because of the timeout but we don't care
    }
    else {
        PostMessage(WM_MOUSEWHEEL, wParam, lParam, unset, "ahk_id " targetHandleID)
    }

    ; Uncomment below for debugging. Sometimes apps return a failed result even if it works, so leaving this commented out since it's not reliable
    ; ---------------------------------------------
    ; using := (IsSet(controlHandleID) && controlHandleID) ? "Control" : ((IsSet(windowHandleID) && windowHandleID) ? "Window" : "Given Parameter")
    ; resultStr := result > 0 ? "Failed (" result ")" : "Success (" result ")"
    ; ToolTip("SendMessage WM_MOUSEWHEEL returned: " resultStr "`n wParam: " wParam "`n Delta: " delta "`n" using " Handle: " targetHandleID)
    ; ---------------------------------------------

    ; return result
}

;; ------------------ Windows and Controls -------------------------

/**
 * Checks if a window has a control with a specific class name. Allows wildcards in the form of an asterisk (*).
 * @param {Integer} hwnd The window handle to check
 * @param {String} pattern The exact control class name to match (supports * wildcards)
 * @returns {Bool} True if the window contains a matching control, false otherwise
 */
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

/**
 * Lists all windows for a given process name if provided, otherwise lists all windows.
 * Results are copied to clipboard.
 * @param {String} procString The process name to filter by (e.g., "notepad.exe"), or unset for all windows (default: unset)
 * @param {Bool} detectHidden Whether to include hidden windows in the search (default: false)
 */
static ListAllWindowsForProcess(procString := unset, detectHidden := false) {
    local origDetectHiddenWindowSetting := DetectHiddenWindows(detectHidden) ; Store the setting to restore later if different

    try {
        if (IsSet(procString)) {
            local hWndArray := WinGetList("ahk_exe " procString, unset, unset)
            local ProcessName := procString
        } else {
            local hWndArray := WinGetList(unset, unset, unset)
            local ProcessName := "All"
        }

        finalString := ""

        ; Check if the array was created and contains any windows
        if IsSet(hWndArray) && hWndArray.Length > 0 {
            ; OutputDebug("Found " hWndArray.Length " window(s) for " ProcessName ":`n")
            ; Loop through each HWND found
            for index, hwnd in hWndArray {
                ; Get the title and class for the current HWND
                local title := ""
                local class := ""
                local positionString := ""
                local procName := ""

                try {
                    title := WinGetTitle("ahk_id " hwnd)
                } catch {
                    ; Nothing, we'll know it failed to find byecause it will be blank
                }

                try {
                    class := WinGetClass("ahk_id " hwnd)
                } catch {
                    ; Nothing, we'll know it failed to find byecause it will be blank
                }

                try {
                    ; Get the extended frame bounds (visible window bounds)
                    frameBounds := ThioUtils.RECT()
                    DllCall("dwmapi\DwmGetWindowAttribute",
                        "ptr", hwnd,
                        "uint", 9,  ; DWMWA_EXTENDED_FRAME_BOUNDS
                        "ptr", frameBounds,
                        "uint", 16)

                    ; Calculate visible window dimensions
                    visibleWidth := frameBounds.right - frameBounds.left
                    visibleHeight := frameBounds.bottom - frameBounds.top

                    positionString := Format("Position: {},{} Size: {}x{}", frameBounds.left, frameBounds.top, visibleWidth, visibleHeight)
                }

                try {
                    ; Get the full process name owning the window
                    procName := WinGetProcessName("ahk_id " hwnd)
                }

                ; Add to the final string
                finalString .= Format("`n  Window {}:`n    Process: {} `n    HWND:  {} `n    Class: {} `n    Title: {} `n    {} `n--------------------", index, procName, hwnd, class, title, positionString)
            }
        } else {
            ThioUtils.TooltipWithDelayedRemove("No windows found for process: " ProcessName, 1500)
            return
        }       

        ; Add final string to clipboard
        A_Clipboard := finalString
        ThioUtils.TooltipWithDelayedRemove("Result added to clipboard", 1500)

    ; Always be sure to restore the original setting    
    } finally {
        DetectHiddenWindows(origDetectHiddenWindowSetting) ; Restore the original setting
    }
}

/**
 * Gets all the controls of a window as objects from the Windows API, given the window's HWND.
 * Can be used as a replacement for WinGetControls() which only returns control names.
 * @param {Integer} windowHwnd The window handle to enumerate controls for
 * @param {Boolean} getText Whether to also get the text of each control (default: false)
 * @param {String} filterClassNN Optional classNN filter to only include controls whose classNN contains this substring (default: unset)
 * @param {String} filterText Optional text filter to only include controls whose text contains this substring (default: unset)
 * @returns {ControlInfo[]} Array of custom control objects with properties: Class, ClassNN, ControlID, Hwnd, and optionally Text
 */
static GetAllControlsAsObjects_ViaWindowsAPI(windowHwnd, getText := unset, filterClassNN := unset, filterText := unset) {
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

        ; ; Add control info to the array
        ; if (!IsSet(filterText) and !IsSet(filterName))
        ;     controlsArray.Push(controlObject)
        ; else if (IsSet(getText) and controlObject.Text && InStr(controlObject.Text, filterText))
        ;     controlsArray.Push(controlObject)

        ; Add control info to the array
        if (!IsSet(filterText) and !IsSet(filterClassNN)) {
            controlsArray.Push(controlObject)
        } else {
            textMatch := !IsSet(filterText) or (IsSet(getText) and controlObject.Text and InStr(controlObject.Text, filterText))
            nameMatch := !IsSet(filterClassNN) or (controlObject.ClassNN and InStr(controlObject.ClassNN, filterClassNN))

            if (textMatch and nameMatch) {
                controlsArray.Push(controlObject)
            }
        }

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

; Adapted from https://www.autohotkey.com/boards/viewtopic.php?style=19&f=82&t=133886#p588184
/**
 * Sets the theme of context menus for the current process.
 * @param {Integer|String} appMode The app mode setting: 0=Default, 1=AllowDark, 2=ForceDark, 3=ForceLight, 4=Max (default: 0)
 * @returns {Int} Previous app mode setting, or -1 on error
 * @example SetContextMenuTheme("AllowDark") ; Follow system theme
 * @note The call to this function must be put before creating any menus, otherwise the app must restart to change it.
 */
static SetContextMenuTheme(appMode := 0) {
    static preferredAppMode := { Default: 0, AllowDark: 1, ForceDark: 2, ForceLight: 3, Max: 4 }
    static uxtheme := dllCall("Kernel32.dll\GetModuleHandle", "Str", "uxtheme", "Ptr")

    if (uxtheme) {
        fnSetPreferredAppMode := dllCall("Kernel32.dll\GetProcAddress", "Ptr", uxtheme, "Ptr", 135, "Ptr")
        fnFlushMenuThemes := dllCall("Kernel32.dll\GetProcAddress", "Ptr", uxtheme, "Ptr", 136, "Ptr")
    } else {
        return -1
    }

    if (preferredAppMode.hasProp(appMode))
        appMode := preferredAppMode.%appMode%

    if (fnSetPreferredAppMode && fnFlushMenuThemes) { ; Ensure the functions were found
        prev := dllCall(fnSetPreferredAppMode, "Int", appMode)
        dllCall(fnFlushMenuThemes)
        return prev
    } else {
        return -1
    }
}

;; --------------------------- Windows Apps and Resources -----------------------------------

/**
 * Resolves an ms-resource URI for a given AppX package using SHLoadIndirectString.
 * @param {String} packageFamilyName The package family name (e.g., "Microsoft.ScreenSketch_8wekyb3d8bbwe").
 * @param {String} ResourceUri The full ms-resource URI, such a "ms-resource://Microsoft.ScreenSketch/Resources/MarkupAndShareToast" 
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
 * @param {String} dllname The name or path of the DLL (e.g., "NotificationController.dll" or "%SystemRoot%\system32\shell32.dll")
 * @param {String} resourceId The resource ID of the string resource, as a string. Usually this is a number with a negative sign (e.g., "-100").
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
 * @param {String} indirectString The indirect string to resolve, formatted as "@{PackageFullName?ms-resource-uri}" or "@%SystemRoot%\system32\shell32.dll,-100".
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
 * @param {String} PackageFamilyName The package family name (e.g., "Microsoft.ScreenSketch_8wekyb3d8bbwe").
 * @returns {String|false} The full package name string, or 'false' if not found or an error occurs.
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

/**
 * Launch any program and move it to the mouse position, with parameters for relative offset vs mouse position.
 * @param {String} programTitle The program executable name or window title to launch
 * @param {Integer} xOffset Horizontal offset from mouse position in pixels (default: 0)
 * @param {Integer} yOffset Vertical offset from mouse position in pixels (default: 0)
 * @param {String} exePath Optional path to the executable (may be faster and more reliable) (default: "")
 * @param {Bool} forceWinActivate Whether to force window activation after positioning (default: false)
 * @param {Integer} sizeX Optional window width in pixels (default: 0 for no resize)
 * @param {Integer} sizeY Optional window height in pixels (default: 0 for no resize)
 */
static LaunchProgramAtMouse(programTitle, xOffset := 0, yOffset := 0, exePath := "", forceWinActivate := false, sizeX := 0, sizeY := 0) {
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
    if (exePath != "")
        Run(exePath)
    else
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

;; --------------------------- Clipboard -----------------------------------

/**
 * Checks if a particular clipboard format is currently available on the clipboard.
 * @param {String} formatName The name of the clipboard format to check for (default: "")
 * @param {Integer} formatIDInput The numeric ID of the clipboard format (default: unset)
 * @returns {Bool} True if the format is available on the clipboard, false otherwise
 * @note Either formatName or formatIDInput must be provided
 */
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

/**
 * Gets the raw bytes data of a specific clipboard format.
 * @param {String} formatName The name of the clipboard format to retrieve (default: "")
 * @param {Integer} formatIDInput The numeric ID of the clipboard format (default: unset)
 * @returns {Array} Array of bytes representing the clipboard format data, or empty array if not found
 * @note Either formatName or formatIDInput must be provided
 */
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

/**
 * Simulate typing a string letter by letter
 * @param {string} text 
 * @param {integer} delayMs
 * @returns {void}
 */
static TypeString(text, delayMs) {
    loop StrLen(text) {
        SendText(SubStr(text, A_Index, 1))
        Sleep(delayMs)
    }
}

;; ------------------------- Tooltip ------------------------------

/**
 * Display a tooltip with automatic removal after a specified delay.
 * @param {String} text The text to display in the tooltip
 * @param {Integer} delayMs The delay in milliseconds before removing the tooltip
 * @param {Integer} x Optional X coordinate for tooltip position relative to mouse (default: unset for mouse position)
 * @param {Integer} y Optional Y coordinate for tooltip position relative to mouse (default: unset for mouse position)
 * @param {Bool} repositionbottom Whether to reposition the tooltip above the mouse if y is negative. Set to false if wanting to manually position. (default: true)
 * @param {Integer} whichToolTip The tooltip ID to use (1 or 2) (default: 1)
 */
static TooltipWithDelayedRemove(text, delayMs, x := unset, y := unset, repositionbottom := true, whichToolTip := 1) {
    if (IsSet(x) && IsSet(y)) {
        ; Set the coordinates relative to the mouse
        originalTooltipCoordMode := A_CoordModeToolTip
        originalMouseCoordMode := A_CoordModeMouse
        CoordMode("ToolTip", "Screen")

        ; Get the current mouse position, adjust coords based on that
        MouseGetPos(&mouseX, &mouseY)
        finalX := mouseX + x
        finalY := mouseY + y

        ; If the tooltip Y position is negative and repositionBottom is true, we'll recalculate the position so the bottom of tooltip is at the y
        if (y < 0 && repositionbottom) {
            tooltipSize := this.GetTooltipSize(text)
            finalY := finalY - tooltipSize.Height
        }

        ; Show the tooltip
        ToolTip(text, finalX, finalY)

        ; Restore the coordinate modes in case they were different
        CoordMode("ToolTip", originalTooltipCoordMode)
        CoordMode("Mouse", originalMouseCoordMode)
    } else {
        ToolTip(text)
    }

    this.RemoveToolTip(delayMs, whichToolTip)
}

/**
 * Remove the current tooltip, optionally after a delay.
 * @param {Integer} delayMs The delay in milliseconds before removing the tooltip (default: 0 for immediate removal)
 */
static RemoveToolTip(delayMs := 0, whichToolTip := 1) {
    ; Local function to use in the timer callback
    SetNoTooltip(whichToolTipLocal := unset) {
        ToolTip(unset, unset, unset, whichToolTip)  ; Calling ToolTip with no parameters removes it
    }

    if delayMs > 0 {
        SetTimer(SetNoTooltip, -1 * delayMs)
    } else {
        SetNoTooltip(whichToolTip)
    }
}

/**
 * Gets the dimensions of a tooltip displaying a given text
 * @param {string} text Text of the tooltip to get the size of
 * @param {Integer} whichToolTip 
 * @returns {Object} 
 */
static GetTooltipSize(text, whichToolTip := 2) {
    origTooltipCoordMode := CoordMode("ToolTip", "Screen")
    ; Set the tooltip to show off screen so we can get the dimensions of it without it being visible, to reshow it where we want.
    ; Apparently doesn't work in AHK V2 it moves it on screen, but this is fast enough where it apparently isn't visible anyway
    ToolTip(text, unset, A_ScreenHeight + 500, whichToolTip) 
    WinGetPos(&X, &Y, &tW, &tH, "ahk_class tooltips_class32")
    ThioUtils.RemoveToolTip(unset, whichToolTip)
    CoordMode("ToolTip", origTooltipCoordMode)
    return { Width: tW, Height: tH }
}


;; ------------------------- High precision timer functions --------------------

/**
 * Start a high precision timer using Windows Performance Counter.
 * @returns {Int64} The performance counter value to pass to EndTimer
 * @note Using these as function calls will add significant overhead if measuring small time intervals (Under ~0.1 ms)
 */
static StartTimer() {
    DllCall("QueryPerformanceFrequency", "Int64*", &freq := 0) ; Get the frequency of the counter
    DllCall("QueryPerformanceCounter", "Int64*", &CounterBefore := 0)
    return CounterBefore
}

/**
 * End a high precision timer and calculate elapsed time.
 * @param {Int64} CounterBefore The performance counter value from StartTimer
 * @param {Bool} showMsgBox Whether to display the result in a message box (default: true)
 * @returns {Float} Elapsed time in milliseconds
 */
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
