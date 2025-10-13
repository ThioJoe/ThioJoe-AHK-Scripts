; This script lets you press a key (default middle mouse) within an Explorer Save/Open dialog window, and it will show a list of paths from any currently open Directory Opus and/or Windows Explorer windows.
; Source Repo: https://github.com/ThioJoe/AHK-Scripts
; Parts of the logic from this script: https://gist.github.com/akaleeroy/f23bd4dd2ddae63ece2582ede842b028#file-currently-opened-folders-md

; HOW TO SET UP:
; - Either run this script by itself, or include it in your main script using #Include
; - Ensure the required RemoteTreeView class file is in the location in the #Include line
; - Configure any settings using the GUI or by editing the settings file directly

; Set to false if you want to to not have the script automatically create an instance of the class when it is included, if you want to call the constructor manually.
;       You would call the class constructor manually and can pass a SystemTraySettings object to the constructor if you want to customize that.
autoCreateInstance := true 

; ---------------------------------------------------------------------------------------------------

#Requires AutoHotkey v2.0
#SingleInstance force
SetWorkingDir(A_ScriptDir)
; Set the path to the RemoteTreeView class file as necessary. Here it is up one directory then in the Lib folder. Necessary to navigate legacy folder dialogs.
; Can be acquired from: https://github.com/ThioJoe/AHK-RemoteTreeView-V2/blob/main/RemoteTreeView.ahk
#Include "..\Lib\RemoteTreeView.ahk"

; ------------------------------------------ INITIALIZATION ----------------------------------------------------

; Set global variables about the program and compiler directives. These use regex to extract data from the lines above them (A_PriorLine)
; Keep the line pairs together!
global g_pathSelector_version := "1.5.0.0"
;@Ahk2Exe-Let ProgramVersion=%A_PriorLine~U)^(.+"){1}(.+)".*$~$2%

global g_pathSelector_programName := "Explorer Dialog Path Selector"
;@Ahk2Exe-Let ProgramName=%A_PriorLine~U)^(.+"){1}(.+)".*$~$2%

; Compiler Options for exe manifest - Arguments: RequireAdmin, Name, Version, UIAccess
;    - The UIAccess option is necessary to allow the script to run in elevated windows protected by UAC without running as admin
;    - Be aware enabling UI Access for compiled would require the script to be signed to work properly and then placed in a trusted location
;    - If you enable the UI Access argument and DON'T sign it, Windows won't run it and will give an error message, therefore not recommended to set the last argument to '1' unless you will self-sign the exe or have a trusted certificate
;    - If you do sign it, it will still run even if not in a trusted location, but it just won't work in elevated dialog Windows
; Note: Though this may have UI Access argument as 0, for my releases I set it to 1 and sign the exe
;@Ahk2Exe-UpdateManifest 0, %U_ProgramName%, %U_ProgramVersion%, 1
;@Ahk2Exe-SetVersion %U_ProgramVersion%
;@Ahk2Exe-SetProductName %U_ProgramName%
;@Ahk2Exe-SetName %U_ProgramName%
;@Ahk2Exe-SetCopyright ThioJoe
;@Ahk2Exe-SetDescription Explorer Dialog Path Selector

global EDPS := unset
; Create an instance of the class and set it as the main instance
if (autoCreateInstance) {
    ; Technically the setAsMainInstance argument is not necessary because we're assigning it to EDPS anyway
    EDPS := ExplorerDialogPathSelector(unset, unset, false, true) 
}

class ExplorerDialogPathSelector {
    ; Class level properties to hold current settings
    g_pth_Settings := {}
    g_pth_SettingsFile := {}
    ; Create property to map to the DefaultSettingsClass class, to be able to use "this.DefaultSettings", because otherwise nested classes need to be referenced with the entire parent class name apparently
    DefaultSettings := ExplorerDialogPathSelector.Settings()
    ; Dictionary to hold localized strings. Defaults to English initially for clarity here but will be loaded during initialization
    localstr := {
        isEnglish: true,
        addressBar: "Address: ",
        thisPC: "This PC",
        computer: "Computer",
        myComputer: "My Computer",
        desktop: "Desktop",
        addressBarKey: "d" ; The 'Alt' key to focus the address bar in Explorer dialogs. May be different in other languages so need to get that too
    }

    ; Need to bind certain function objects that are used in callbacks
    ; Used for the hotkey to show the menu
    boundDisplayFunc := this.DisplayDialogPathMenu.Bind(this)
    ShowPathSelectorSettingsGUI_Bound := this.ShowPathSelectorSettingsGUI.Bind(this)
    ShowPathSelectorHelpWindow_Bound := this.ShowPathSelectorHelpWindow.Bind(this)

    ; Constructor
    __New(settingsObject :=  ExplorerDialogPathSelector.Settings(), systemTraySettingsObject := ExplorerDialogPathSelector.SystemTraySettings(), dontLoadSettingsFile := false, setAsMainInstance := unset) {
        ; Assign the localized strings needed for the script
        this.GetNecessaryLocalizedStrings()

        ; ------------------- Validate the parameters -------------------
        if (systemTraySettingsObject.base.__Class != ExplorerDialogPathSelector.SystemTraySettings.prototype.__Class) {
            MsgBox("Invalid object type was used for Explorer Dialog Path Selector system tray settings. Must be type: ExplorerDialogPathSelector.SystemTraySettings . `n`n Using default settings instead.")
            systemTraySettingsObject := ExplorerDialogPathSelector.SystemTraySettings()
        }
        if (settingsObject.base.__Class != ExplorerDialogPathSelector.Settings.prototype.__Class) {
            MsgBox("Invalid object type was used for Explorer Dialog Path Selector settings. Must be type: ExplorerDialogPathSelector.Settings . `n Using default settings instead.")
            settingsObject := ExplorerDialogPathSelector.Settings()
        }
        ; ---------------------------------------------------------------

        ; Construct object with info about the settings file. Class creates the various object properties using inputs for the file name and folder name in AppData
        this.g_pth_SettingsFile := ExplorerDialogPathSelector.SettingsFile(
            "ExplorerDialogPathSelector-Settings.ini", ; Name of the settings file
            "Explorer-Dialog-Path-Selector"            ; Name of the folder in AppData where the settings file will be stored
        )

        ; ---------- Conditional Default Settings ----------
        ; Check for the existence of the hard coded default Directory Opus exe path and disable Directory Opus integration if it doesn't exist
        ; This is to prevent the script from trying to use Directory Opus integration if the path is invalid, but still load whatever value from user settings file if there is one
        if !FileExist(this.DefaultSettings.dopusRTPath) {
            this.DefaultSettings.dopusRTPath := ""
        }
        ; If the script is not running standalone or compiled, disable UI Access by default
        if !this.ThisScriptRunningStandalone() or A_IsCompiled {
            this.DefaultSettings.enableUIAccess := false
        }

        ; Set initial settings. This will be overridden by settings in the settings ini file if it exists.
        ; If it wasn't passed in as an argument to the constructor, it will contain defaults
        this.g_pth_Settings := this.CopySettingsObject(settingsObject)

        ; ------------------ Load settings Files ------------------
        if (!dontLoadSettingsFile) {
            try {
                this.LoadSettingsFromSettingsFilePath()
            } catch Error as err {
                MsgBox("Error reading settings file: " err.Message "`n`nUsing default settings.")
                ; Set each property individually instead of directly assigning the object to the settings object to avoid making a reference
                for k, v in this.DefaultSettings.OwnProps() {
                    this.g_pth_Settings.%k% := this.DefaultSettings.%k%
                }
            }
        }
        ; ----------------------------------------------------------
        
        ; ----- Special handling for certain settings -----
        ; For UI Access, always disable if not running standalone
        if !this.ThisScriptRunningStandalone() or A_IsCompiled {
            this.g_pth_Settings.enableUIAccess := false
        }
        ; If the script is running standalone and UI access is installed...
        ; Reload self with UI Access for the script - Allows usage within elevated windows protected by UAC without running the script as admin
        ; See Docs: https://www.autohotkey.com/docs/v1/FAQ.htm#uac
        if (this.g_pth_Settings.enableUIAccess = true) and !A_IsCompiled and this.ThisScriptRunningStandalone() and !InStr(A_AhkPath, "_UIA") {
            Run("*uiAccess " A_ScriptFullPath)
            ExitApp()
        }
        ; --------------------------------------------------------

        this.SetMenuTheme("AllowDark") ; Set the menu theme to follow system theme
        this.SetupSystemTray(systemTraySettingsObject, this.g_pth_Settings.hideTrayIcon)
        this.UpdateHotkey("", "") ; Initialize the hotkey. It will use the hotkey from settings

        ; ---------- Main Instance Handling ---------------------
        ; If setAsMainInstance is explicitly true, set this instance as the main instance
        if (IsSet(setAsMainInstance) && setAsMainInstance == true) {
            global EDPS := this
        ; If there is no primary instance yet, set this instance as the main instance, unless setAsMainInstance is explicitly false. 
        } else if (!IsSet(EDPS) && (!IsSet(setAsMainInstance) || (IsSet(setAsMainInstance) && setAsMainInstance != false))) { 
            global EDPS := this
        ; If there is already a primary instance, only set this instance as the main instance if setAsMainInstance is explicitly true
        } else if (IsSet(EDPS) && (IsSet(setAsMainInstance) && setAsMainInstance == true)) { 
            global EDPS := this
        }
        ; If setAsMainInstance is false, do not set the global EDPS variable to this instance of the class
        ; --------------------------------------------------------
    }

    CopySettingsObject(original) {
        newObj := {}
        for k, v in original.OwnProps() {
            newObj.%k% := original.%k%
        }
        return newObj
    }

    ; ---------------------------------------- DEFAULT USER SETTINGS ----------------------------------------
    ; These will be overridden by settings in the settings ini file if it exists. Otherwise these defaults will be used.
    class Settings {
        ; Hotkey to show the menu. Default is Middle Mouse Button. If including this script in another script, you could choose to set this hotkey in the main script and comment this line out
        dialogMenuHotkey := "~MButton"
        ; Enable debug mode to show tooltips with debug info
        enableExplorerDialogMenuDebug := false
        ; Whether to show the disabled clipboard path menu item when no valid path is found on the clipboard, or only when a valid path is found on the clipboard
        alwaysShowClipboardmenuItem := true
        ; Whether to enable UI access by default to allow the script to run in elevated windows without running as admin
        enableUIAccess := true
        ; Group paths by Window or not
        groupPathsByWindow := true
        ; Use simulated bold text for active paths in the menu
        useBoldTextActive := true
        activeTabSuffix := ""            ;  Appears to the right of the active path for each window group
        activeTabPrefix := "► "          ;  Appears to the left of the active path for each window group
        standardEntryPrefix := "    "    ; Indentation for inactive tabs, so they line up
        ; Path to dopusrt.exe - can be empty to explicitly disable Directory Opus integration, but it will automatically disable if the file is not found anyway
        dopusRTPath := "C:\Program Files\GPSoftware\Directory Opus\dopusrt.exe"
        favoritePaths := []              ; Array of favorite paths to show at the top of the menu
        conditionalFavorites := []       ; Array of objects with 'path' and 'condition' properties. If the condition is true when the menu is shown, the path will be added to the favorites

        ; Settings appearing in the settings file only, not in the GUI
        maxMenuLength := 120             ; Maximum length of menu items. The true max is MAX_PATH, but thats really long so this is a reasonable limit
        hideTrayIcon := false            ; Whether to hide the tray icon or not
    }

    class SystemTraySettings {
        settingsMenuItemName := "Path Selector Settings   " ; Added spaces to the name or else it can get cut off when default menu item
        showSettingsTrayMenuItem := true  ; Show the settings menu item in the system tray menu. If set to false none of these other settings matter
        forcePositionIndex := false       ; If this is false, the position index will be increased by 1 if the script is running as included in another script (so it shows up after the parent script's menu items)
        positionIndex := 1                ; Position index for the settings menu item. 1 is the first item, 2 is the second, etc.
        addSeparatorBefore := false       ; Add a separator before the settings menu item
        addSeparatorAfter := true         ; Add a separator after the settings menu item
        alwaysDefaultItem := false        ; If true, the path selector menu item will always be the default even if the script is included in another script
        customIconFilePath := ""          ; Path to a custom icon file for the system tray icon. If not set, the default icon will be used
        ;hideTrayIcon := unset            ; This will be set only by g_pth_Settings.hideTrayIcon
    }

    ; ---------------------------------------- EASIER NAMED METHODS TO BE CALLED FROM OTHER SCRIPTS  ----------------------------------------------

    ; Create both a static and instance method to update the primary instance of the class because apparently static methods aren't visible in instance objects
    static ShowSettingsGUI => EDPS.ShowPathSelectorSettingsGUI_Bound
    ShowSettingsGUI => this.ShowPathSelectorSettingsGUI_Bound

    ; In case the user wants to create a new instance of the class and set it as the primary instance
    static UpdatePrimaryInstance(EDPS_Object) {
        if (IsObject(EDPS_Object) && (EDPS.base.__Class == ExplorerDialogPathSelector.prototype.__Class)) {
            global EDPS := EDPS_Object
        } else {
            MsgBox("Invalid object passed to UpdatePrimaryInstance. Must be an instance of ExplorerDialogPathSelector class.`n This can be created like:`n`nsomeInstanceName := ExplorerDialogPathSelector()")
            return
        }
    }

    ; ---------------------------------------- INITIALIZATION FUNCTIONS AND CLASSES  ----------------------------------------------

    ; Updates the hotkey. If no new hotkey is provided, it will use the hotkey from settings
    UpdateHotkey(newHotkey := "", previousHotkeyString := "") {
        ; Use the hotkey from settings if no new hotkey is provided
        if (newHotkey = "") {
            newHotkey := this.g_pth_Settings.dialogMenuHotkey
        }

        ; If the new hotkey is the same as before, return. Otherwise it will disable itself and re-enable itself unnecessarily
        ; By now newHotkey should have a value either from being set or from the settings
        if (newHotkey = "") or (newHotkey = previousHotkeyString)
            return

        if (previousHotkeyString != "") {
            try {
                HotKey(previousHotkeyString, "Off")
            } catch Error as hotkeyUnsetErr {
                MsgBox("Error disabling previous hotkey: " hotkeyUnsetErr.Message "`n`nHotkey Attempted to Disable:`n" previousHotkeyString "`n`nWill still try to set new hotkey.")
            }
        }

        try {
            HotKey(newHotkey, this.boundDisplayFunc, "On") ; Include 'On' option to ensure it's enabled if it had been disabled before, like changing the hotkey back again
        } catch Error as hotkeySetErr {
            MsgBox("Error setting hotkey: " hotkeySetErr.Message "`n`nHotkey was trying to be set to:`n" newHotkey)
        }
    }

    ; Loads Windows resources to get localized versions of strings that are used to locate dialogs and their controls
    GetNecessaryLocalizedStrings() {
        ; Resource string addresses
        addressbarRes := "@%SystemRoot%\system32\explorerframe.dll,-13829" ; "Address: %s" (Need to remove the %s part)
        addressbarKeyStrRes := "@%SystemRoot%\system32\explorerframe.dll,-12896" ; "A&ddress" (Need to get the character after the &). NOT resource 13137
        thisPCRes := "@%SystemRoot%\system32\shell32.dll,-30621" ; "This PC"
        computerRes := "@%SystemRoot%\system32\shell32.dll,-30480" ; "Computer"
        myComputerRes := "@%SystemRoot%\system32\shell32.dll,-9012" ; "My Computer"
        desktopRes := "@%SystemRoot%\system32\shell32.dll,-4162" ; "Desktop"

        ; Ones that need further processing
        addressBarRaw := this.ResolveWindowsResource(addressbarRes)
        addressbarKeyStrRaw := this.ResolveWindowsResource(addressbarKeyStrRes)
        thisPCResRaw := this.ResolveWindowsResource(thisPCRes)
        computerResRaw := this.ResolveWindowsResource(computerRes)
        myComputerResRaw := this.ResolveWindowsResource(myComputerRes)
        desktopResRaw := this.ResolveWindowsResource(desktopRes)

        ; If the result isn't false, we can set the string
        if (addressBarRaw)
            this.localstr.addressBar := StrReplace(addressBarRaw, "%s", "")
        if (addressbarKeyStrRaw)
            this.localstr.addressBarKey := SubStr(addressbarKeyStrRaw, InStr(addressbarKeyStrRaw, "&") + 1, 1)
        if (thisPCResRaw)
            this.localstr.thisPC := thisPCResRaw
        if (computerResRaw)
            this.localstr.computer := computerResRaw
        if (myComputerResRaw)
            this.localstr.myComputer := myComputerResRaw
        if (desktopResRaw)
            this.localstr.desktop := desktopResRaw

        ; If the system language is not English, set isEnglish to false
            this.localstr.isEnglish := this.IsSystemLanguageEnglish()
    }

    ; Stores info about the settings file - Using a class instead of object so that we can dynamically set certain properties based on the file name and folder name
    class SettingsFile {
        ; Explicitly declare these properties as static because there should really only be one instance of it
        fileName             := unset
        directoryPath        := unset
        appDataDirName       := unset
        filePath             := unset
        appDataDirectoryPath := unset
        appDataFilePath      := unset
        usingSettingsFile    := unset

        ; Actually set the properties using inputs for the file name and folder name in AppData
        __New(fileName, appDataFolderName) {
            this.fileName               := fileName
            this.directoryPath          := A_ScriptDir
            this.appDataDirName         := appDataFolderName
            this.filePath               := this.directoryPath "\" fileName
            this.appDataDirectoryPath   := A_AppData "\" this.appDataDirName
            this.appDataFilePath        := this.appDataDirectoryPath "\" fileName
            this.usingSettingsFile      := false
        }
    }

    ; ---------------------------------------- UTILITY FUNCTIONS  ----------------------------------------------
    ; Function to check if the script is running standalone or included in another script
    ThisScriptRunningStandalone() {
        ;MsgBox("A_ScriptName: " A_ScriptFullPath "`n`nA_LineFile: " A_LineFile "`n`nRunning Standalone: " (A_ScriptFullPath = A_LineFile ? "True" : "False"))
        return A_ScriptFullPath = A_LineFile
    }

    RemoveEmptyArrayEntries(arr) {
        for i, v in arr {
            if (v = "")
                arr.RemoveAt(i)
        }
        return arr
    }

    ValidatePathCharacters(path) {
        return RegExMatch(path, "^[^<>`"\/|?*]+$")
    }

    ValidatePathCharacters_AllowWildCards(path) {
        return RegExMatch(path, "^[^<>`"\/|?]+$")
    }

    ; Enum class for condition types
    class ConditionType {
        static DialogOwnerExe := {
            StringID: "DialogOwnerExe",
            FriendlyName: "Executable Name Match",
            Description: "If the dialog window was opened by an executable file name that matches the value. ( * is a wildcard )"
        }
        static CurrentDialogPath := {
            StringID: "CurrentDialogPath",
            FriendlyName: "Dialog Path Match",
            Description: "If the current path of the dialog window matches the value.`n( * is a wildcard )"
        }
        static ParentWindowTitle := {
            StringID: "ParentWindowTitle",
            FriendlyName: "Parent Process Window Title",
            Description: "If the dialog window was opened by a window who's title matches the value. ( * is a wildcard )"
        }
        static DialogFileTypeFilter := {
            StringID: "DialogFileTypeFilter",
            FriendlyName: "Dialog File Type Filter",
            Description: "If the dialog window filters for file extensions that matches the value. ( * is a wildcard )"
        }
    }

    ; Check for match using strings with wildcard asterisks
    StringMatchWithWildcards(str, matchStr) {
        return RegExMatch(str, "i)^" RegExReplace(matchStr, "\*", ".*") "$")
    }

    ; Check with the windows API which language the system is using
    IsSystemLanguageEnglish() {
        local lang := A_Language
        ; If it's not english we need to use the localized strings
        if (lang != "1033") { ; 1033 is the LCID for English (United States)
            return false
        } else {
            return true
        }
    }

    /**
     * Resolves a windows localized (multilanguage) resource string using SHLoadIndirectString API.
     * @param indirectString The indirect string to resolve, formatted as "@{PackageFullName?ms-resource-uri}" or "@%SystemRoot%\system32\shell32.dll,-100".
     * @returns The resolved string if successful, or an error message if not.
     */
    ResolveWindowsResource(indirectString) {
        try {
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
        } catch as err {
            OutputDebug("`nError: " err.Message "`n`nFailed to resolve resource string: " indirectString)
            return false
        }
    }

    GetLocalizedPathParts(pathPartsList) {
        truePath := "" ; The true path to be built progressively and used in the lookups
        isFirstPart := true
        displayPathPartsList := [] ; Initialize an empty list to hold the display parts

        ; Progressively build the display path by resolving each part
        for part in pathPartsList {
            if (part = "") {
                continue ; Skip empty parts (like leading backslash)
            }

            if (truePath != "" && !isFirstPart) {
                truePath .= "\"
            }

            ; displayPath .= GetFolderDisplayName(displayPath . "\" . part) ; Get the display name for the current part
            newPart := this.GetFolderDisplayName(truePath . part) ; Get the display name for the current part

            if (newPart) {
                displayPathPartsList.Push(newPart) ; Add the display name to the list
                truePath .= part ; Append the original part to the true path
            } else {
                ; If we can't get the display name, use the original part
                displayPathPartsList.Push(part) ; Add the original part to the list
                truePath .= part
            }

            if (isFirstPart)
                isFirstPart := false
        }

        return displayPathPartsList ; Return the list of display names
    }

    /**
     * Gets the localized display name of a file or folder. It only gets the name of the last part of the path, not the full path.
     * This uses the SHGetFileInfoW function to get the name as it appears in Windows Explorer.
     * @param path {String} The full path to the file or folder.
     * @returns {String|Boolean} The display name on success, or `false` on failure.
     */
    GetFolderDisplayName(path) {
        try {
            ; --- Define constants for SHGetFileInfoW ---
            local SHGFI_DISPLAYNAME := 0x200 ; Flag to retrieve the display name.
            local MAX_PATH := 260            ; The maximum length for the szDisplayName field.

            ; --- Prepare the SHFILEINFOW structure ---
            ; The structure has the following members in order:
            ; HICON hIcon          (Ptr size)
            ; int   iIcon          (4 bytes)
            ; DWORD dwAttributes   (4 bytes)
            ; WCHAR szDisplayName[MAX_PATH] (260 * 2 bytes)
            ; WCHAR szTypeName[80] (80 * 2 bytes)
            local shfi_size := A_PtrSize + 4 + 4 + (MAX_PATH * 2) + (80 * 2)
            local shfi := Buffer(shfi_size, 0)

            ; --- Call the SHGetFileInfoW function ---
            ; For more info, see: https://learn.microsoft.com/en-us/windows/win32/api/shellapi/nf-shellapi-shgetfileinfow
            local hResult := DllCall("shell32\SHGetFileInfoW",
                "WStr", path,       ; [in] LPCWSTR pszPath
                "UInt", 0,          ; [in] DWORD dwFileAttributes
                "Ptr", shfi,       ; [out] SHFILEINFOW *psfi
                "UInt", shfi.Size,  ; [in] UINT cbFileInfo
                "UInt", SHGFI_DISPLAYNAME) ; [in] UINT uFlags

            ; Check if the function call was successful. A non-zero return value indicates success.
            if (hResult != 0) {
                ; --- Extract the display name from the structure ---
                ; The display name is not at the start of the buffer. We need to calculate its offset.
                ; Offset = size of hIcon + size of iIcon + size of dwAttributes
                local displayNameOffset := A_PtrSize + 4 + 4

                ; Retrieve the null-terminated UTF-16 string from the calculated offset.
                local displayName := StrGet(shfi.Ptr + displayNameOffset, "UTF-16")
                return displayName
            } else {
                ; The function call failed.
                OutputDebug("`nError: SHGetFileInfoW failed for path: " path)
                return false
            }
        } catch as err {
            ; Catch any other errors that might occur.
            OutputDebug("`nError: " err.Message "`n`nFailed to get display name for path: " path)
            return false
        }
    }

    ; Sets the theme of menus by the process - Adapted from https://www.autohotkey.com/boards/viewtopic.php?style=19&f=82&t=133886#p588184
    ; Usage: Put this before creating any menus. AllowDark will folow system theme. Seems that once set, the app must restart to change it.
    ;SetMenuTheme("ForceDark")
    SetMenuTheme(appMode:=0) {
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

    ; ------------------------------------ MAIN LOGIC FUNCTIONS ---------------------------------------------------

    ; Get the paths of all tabs from all Windows of Windows Explorer, and identify which are the active tab for each window
    GetAllExplorerPaths() {
        paths := []
        explorerHwnds := WinGetList("ahk_class CabinetWClass")
        shell := ComObject("Shell.Application")

        static IID_IShellBrowser := "{000214E2-0000-0000-C000-000000000046}"

        ; First make a pass through all explorer windows to get the active tab for each
        activeTabs := Map()
        for explorerHwnd in explorerHwnds {
            try {
                if activeTab := ControlGetHwnd("ShellTabWindowClass1", "ahk_id " explorerHwnd) {
                    activeTabs[explorerHwnd] := activeTab
                }
            }
        }

        ; shell.Windows gives us a collection of all open open tabs as a flat list, not separated by window, so now we loop through and match them up by Hwnd
        ; Now do a single pass through all tabs
        for tab in shell.Windows {
            try {
                ; Ensure we have the handle of the tab
                if tab && tab.hwnd {
                    parentWindowHwnd := tab.hwnd
                    path := tab.Document.Folder.Self.Path
                    if path {
                        ; Check if this tab is active
                        isActive := false
                        ; If we have any active tab at all for the parent window
                        if activeTabs.Has(parentWindowHwnd) {
                            ; Get an interface to interact with the tab's shell browser
                            shellBrowser := ComObjQuery(tab, IID_IShellBrowser, IID_IShellBrowser)
                            ; Call method of Index 3 on the interface to get the tab's handle so we can see if any windows have such an active tab
                            ; We need to know the method index number from the "vtable" - Apparently the struct for a vtable is often named with "Vtbl" at the end like "IWhateverInterfaceVtbl"
                            ; IShellBrowserVtbl is in the Windows SDK inside ShObjIdl_core.h. The first 3 methods are AddRef, Release, and QueryInterface, inhereted from IUnknown,
                            ;       so the first real method is the fourth, meaning index 3, which is GetWindow and is also the one we want
                            ; The output of the ComCall GetWindow method here is the handle of the tab, not the parent window, so we can compare it to the activeTabs map
                            ComCall(3, shellBrowser, "uint*", &thisTab := 0)
                            isActive := (thisTab = activeTabs[parentWindowHwnd])
                        }

                        paths.Push({
                            Hwnd: parentWindowHwnd,
                            Path: path,
                            IsActive: isActive
                        })
                    }
                }
            }
        }
        return paths
    }

    ; Call Directory Opus' DOpusRT to create temporary XML file with info about Opus open windows. Parse the XML and return an array of path objects
    GetDOpusPaths() {
        if (this.g_pth_Settings.dopusRTPath = "") {
            return []
        }

        if !FileExist(this.g_pth_Settings.dopusRTPath) {
            MsgBox("Directory Opus Runtime (dopusrt.exe) not found at:`n" this.g_pth_Settings.dopusRTPath "`n`nDirectory Opus integration won't work. To enable it, set the correct path in the script configuration. Or set it to an empty string to avoid this error.", "DOpus Integration Error", "Icon!")
            return []
        }

        tempFile := A_Temp "\dopus_paths.xml"
        try FileDelete(tempFile) ; Delete any existing temp file

        try {
            cmd := '"' this.g_pth_Settings.dopusRTPath '" /info "' tempFile '",paths'
            RunWait(cmd, unset, "Hide")

            if !FileExist(tempFile)
                return []

            xmlContent := FileRead(tempFile)
            FileDelete(tempFile)

            ; Parse paths from XML
            pathsObjectArray := []

            ; Start after the XML declaration
            xmlContent := RegExReplace(xmlContent, "^.*?<results.*?>", "")

            ; Extract each path element with its attributes
            while RegExMatch(xmlContent, "s)<path([^>]*)>(.*?)</path>", &match) {
                ; Get attributes
                attrs := Map()
                RegExMatch(match[1], "lister=`"(0x[^`"]*)`"", &listerMatch)
                RegExMatch(match[1], "active_tab=`"([^`"]*)`"", &activeTabMatch)
                RegExMatch(match[1], "active_lister=`"([^`"]*)`"", &activeListerMatch)

                ; Unescape any XML characters as necessary - Only need to worry about &amp; and &apos; because the rest are not valid file path characters
                pathStr := match[2]
                pathStr := StrReplace(pathStr, "&amp;", "&")
                pathStr := StrReplace(pathStr, "&apos;", "'")

                ; Create path object
                pathObj := {
                    path: pathStr,
                    lister: listerMatch ? listerMatch[1] : "unknown",
                    isActiveTab: activeTabMatch ? (activeTabMatch[1] = "1") : false,
                    isActiveLister: activeListerMatch ? (activeListerMatch[1] = "1") : false
                }
                pathsObjectArray.Push(pathObj)

                ; Remove the processed path element and continue searching
                xmlContent := SubStr(xmlContent, match.Pos + match.Len)
            }

            return pathsObjectArray
        } catch as err {
            MsgBox("Error reading Directory Opus paths: " err.Message "`n`nDirectory Opus integration will be disabled.", "DOpus Integration Error", "Icon!")
            return []
        }
    }

    ; Display the menu
    DisplayDialogPathMenu(thisHotkey) { ; Called via the Hotkey function, so it must accept the hotkey as its first parameter
        ; ------------------------- LOCAL FUNCTIONS -------------------------
        ProcessCLSIDPaths(clsidInputPath) {
            returnObj := {displayText: "", clsidNewPath: ""}
            try {
                ; Create Shell.Application COM object
                shell := ComObject("Shell.Application")
                folder := shell.Namespace(clsidInputPath)
                displayName := folder.Self.Name
                clsidNewPath := clsidInputPath

                ; Further check to process library-ms files
                if (SubStr(clsidInputPath, -11) = ".library-ms") {
                    ; Remove the ".library-ms" extension or else it won't work when navigating
                    clsidNewPath := SubStr(clsidInputPath, 1, -11)
                    
                    ; If there are CLSIDs in the path, convert them to folder names to show full context instead of just showing the final CLSID folder name
                    ; Example: "::{031E4825-7B94-4DC3-B131-E946B44C8DD5}\Music" -> "Libraries\Music"
                    ;     And show Libraries\Music instead of just Music
                    if (RegExMatch(clsidNewPath, "::{[^}]+}")) {
                        fullContextDisplayName := clsidNewPath ; For libraries may end up looking like "Libraries\Documents"

                        while (RegExMatch(fullContextDisplayName, "::{[^}]+}", &match)) {
                            clsidName := shell.Namespace(match[0]).Self.Name
                            fullContextDisplayName := StrReplace(fullContextDisplayName, match[0], clsidName)
                        }
                        displayName := fullContextDisplayName
                    }
                }

                returnObj.displayText := displayName
                returnObj.clsidNewPath := clsidNewPath

                return returnObj

            } catch as err {
                OutputDebug("Error Getting Folder Name From CLSID " clsidInputPath " : " err.Message)
                return returnObj
            }
        }

        GetDialogAddressBarPath(windowHwnd) {
            controls := WinGetControls(windowHwnd)
            ; Go through controls that match "ToolbarWindow32*" in the class name and check if their text starts with "Address: "
            for controlClassNN in controls {
                if (controlClassNN ~= "ToolbarWindow32") {
                    controlText := ControlGetText(controlClassNN, windowHwnd)
                    if (controlText ~= this.localstr.addressBar) {
                        ; Get the path from the address bar
                        return SubStr(controlText, 10)
                    }
                }
            }
            return ""
        }

        /**
         * @param windowHwnd 
         * @returns {String} 
         */
        GetDialogFilterComboBoxText(windowHwnd) {
            controls := WinGetControls(windowHwnd)
            for controlClassNN in controls {
                if (controlClassNN ~= "ComboBox") {
                    controlText := ControlGetText(controlClassNN, windowHwnd)
                    if (controlText ~= "*") { ; If it contains an asterisk from the start it's gotta be the filter box since that can't be in file names
                        ; controlHwnd := ControlGetHwnd(controlClassNN, windowHwnd)
                        return controlText
                    }
                }
            }
            return ""
        }

        ; Replaces the text in the input string with a bold unicode version of the text for latin characters
        SimulateBoldText(inputText) {
            boldText := ""
            len := StrLen(inputText)
            boldSansSerifStartUpper := 0x1D5D4
            boldSansSerifStartLower := 0x1D5EE
            boldSansSerifNumbers := 0x1D7EC
        
            Loop len {
                char := SubStr(inputText, A_Index, 1)
                charCode := Ord(char)
                if IsAlpha(char) {
                    if (charCode >= Ord("A") && charCode <= Ord("Z")) ; Check if character is uppercase
                        boldText .= Chr(boldSansSerifStartUpper + charCode - Ord("A"))
                    else if (charCode >= Ord("a") && charCode <= Ord("z")) ; Check if character is lowercase
                        boldText .= Chr(boldSansSerifStartLower + charCode - Ord("a"))
                }
                else if (charCode >= Ord("0") && charCode <= Ord("9")) ; Check if character is a number
                    boldText .= Chr(boldSansSerifNumbers + charCode - Ord("0"))
                else
                    boldText .= char
            }
            return boldText
        }

        ;global menuItemTrackerObj := {} ; For debugging purposes
        InsertMenuItem(menuObj, text, path := unset, iconPath := unset, iconIndex := unset, isActiveTab := unset) {
            if text = "" { ; It's a separator
                menuObj.Insert(unset)
                currentMenuNum++
                return
            }

            pathStr := "" ; Default to empty string so we can use the same callback without passing f_path parameter unset
            if (IsSet(path)) {
                pathStr := path
            }

            ; If the menuText is greater than the max, truncate it from the right using SubStr. The path should still work since this is only the text displayed in the menu
            stringLength := StrLen(text)
            ; Note: If the path is longer than MAX_PATH (I believe), Windows will convert it to 8.3 format. Navigation still work but our window path matching might not work
            if (stringLength > maxMenuLength) {
                if (isActiveTab)
                    text := this.g_pth_Settings.activeTabPrefix "..." SubStr(text, (-1 * maxMenuLength)) this.g_pth_Settings.activeTabSuffix
                else
                    text := this.g_pth_Settings.standardEntryPrefix "..." SubStr(text, (-1 * maxMenuLength))
            }

            ; If the display text has an ampersand, double it to escape it so Autohotkey displays it correctly
            text := StrReplace(text, "&", "&&")

            ; If it's a CLSID, get the folder name from the CLSID
            usingCLSID := false
            clsidFriendlyName := ""
            if (pathStr ~= "::{") {
                usingCLSID := true
                clsidObj := ProcessCLSIDPaths(pathStr)
                clsidFriendlyName := clsidObj.displayText
                clsidPathStr := clsidObj.clsidNewPath

                if (clsidFriendlyName != "" and clsidPathStr != "") {
                    pathStr := clsidPathStr

                    if (isActiveTab)
                        text := this.g_pth_Settings.activeTabPrefix clsidFriendlyName this.g_pth_Settings.activeTabSuffix
                    else
                        text := this.g_pth_Settings.standardEntryPrefix clsidFriendlyName
                }
            }

            if IsSet(isActiveTab) and isActiveTab and this.g_pth_Settings.useBoldTextActive {
                text := SimulateBoldText(text)
            }

            ; Insert the menu item attached to the callback with parameters that will execute if that particular menu item is clicked
            ; The first 3 parameters of the callback are automatically provided by the Insert method, then we add additional parameters we'll need
            ; For some reason when refactoring the code so everything is in a class, binding the Navigate function stopped working right. 
            ;       It was no longer passing the name of the menu item as the first parameter to Navigate().
            ;       But this instead where we use fat-arrow syntax to pass the parameters works for some reason
            ;           This is probably better anyway since it makes it easier to see what parameters are being passed
            ;       So basically for the callback, we create a temporary wrapper function that calls the Navigate function with the parameters we want
            ;           And apparently this doesn't have issues with classes / this references like the other way did
            menuObj.Insert(unset, text, (iName, iPos, mObj) => this.Navigate(iName, iPos, mObj, pathStr, windowClass, windowID))
            
            currentMenuNum++

            if (!IsSet(path)) { ; It's a header because it's just text
                menuObj.Disable(currentMenuNum "&")
                return
            }
            
            if (IsSet(iconPath) and iconPath and IsSet(iconIndex) and iconIndex) {
                menuObj.SetIcon(currentMenuNum "&", iconPath, iconIndex)
            }

            ; Check if the path matches the dialog already. If so, disable the item
            if (windowPath and (windowPath = path)) {
                menuObj.Disable(currentMenuNum "&")
            }

            ; Additional check for CLSID paths to disable if the window path matches the CLSID path
            if (usingCLSID and (clsidFriendlyName != "") and (windowPath = clsidFriendlyName)) {
                menuObj.Disable(currentMenuNum "&")
            }
        }

        ; ------------------------------------------------------------------------

        debugMode := this.g_pth_Settings.enableExplorerDialogMenuDebug
        maxMenuLength := this.g_pth_Settings.maxMenuLength

        if (debugMode) {
            ToolTip("Hotkey Pressed: " A_ThisHotkey)
            Sleep(1000)
            ToolTip()
        }

        ; Detect the window under the mouse cursor
        windowID := 0
        try {
            MouseGetPos(unset, unset, &windowID)
            windowClass := WinGetClass("ahk_id " windowID)
            windowHwnd := WinExist("ahk_id " windowID)
            windowExe := WinGetProcessName("ahk_id " windowID)
        } catch as err {
            ; If we can't get window info, wait briefly and try once more
            OutputDebug("Error getting window info. Will Try again in 25ms: " err.Message)
            Sleep(25)
            try {
                MouseGetPos(unset, unset, &windowID)
                windowClass := WinGetClass("ahk_id " windowID)
                windowHwnd := WinExist("ahk_id " windowID)
                windowExe := WinGetProcessName("ahk_id " windowID)
            } catch as err {
                OutputDebug("Error getting window info: " err.Message)
                if (debugMode) {
                    ToolTip("Unable to detect active window")
                    Sleep(1000)
                    ToolTip()
                }
                return
            }
        }

        ; Verify we got valid window info
        if (!windowID || !windowClass) {
            if (debugMode) {
                ToolTip("No valid window detected")
                Sleep(3000)
                ToolTip()
            }
            return
        }

        if (debugMode) {
            ToolTip("Window ID: " windowID "`nClass: " windowClass)
            Sleep(2000)
            ToolTip()
        }

        ; Don't display menu unless it's a dialog or console window
        if !(windowClass ~= "^(?i:#32770|ConsoleWindowClass|SunAwtDialog)$") {
            if (debugMode) {
                tooltipText := "Window class does not match expected. Detected: " windowClass
                if (windowClass = "CabinetWClass") {
                    tooltipText .= "`n`nIs this a regular Windows Explorer window? This tool is only meant for 'Save As' and 'Open' type windows."
                }
                ToolTip(tooltipText)
                Sleep(4000)
                ToolTip()
            }
            return
        }

        local parentWindowTitle := ""
        local windowPath := ""
        local windowFilter := ""
        local windowParent := 0

        ; Get the path and other info about the dialog. At this point we know it's the correct type of window
        if (windowHwnd) {
            windowPath := GetDialogAddressBarPath(windowHwnd)
            ; Get info about the parent window name
            try {
                hwndParent := DllCall("GetParent", "Ptr", windowHwnd, "UInt")
                parentWindowTitle := WinGetTitle("ahk_id " hwndParent)
            } catch as err {
                OutputDebug("Error getting parent window or parent window title: " err.Message)
                if (debugMode) {
                    ToolTip("Unable to get parent window")
                    Sleep(1000)
                    ToolTip()
                }
                hwndParent := 0
                parentWindowTitle := ""
            }

            ; Try to get the window filter combobox. Not all windows will have this.
            try {
                windowFilter := ControlGetText("ComboBox2", "ahk_id " windowID)
                ; If it doesn't contain an asterisk it's probably not the right box so we'll do a more thorough search
                if (!InStr(windowFilter, "*")) {
                    windowFilterAttempt2 := GetDialogFilterComboBoxText(windowHwnd)
                    if (windowFilterAttempt2) {
                        windowFilter := windowFilterAttempt2
                    }
                }
            } catch {
                OutputDebug("Couldn't get window filetype filter.")
            }
        }

        ; Proceed to display the menu
        CurrentLocations := Menu()
        hasItems := false
        currentMenuNum := 0 ; Used to keep track of the current menu item number so we can refer to each item by index like "1&" in case of duplicate path entries

        ; Add favorite paths if there are any
        if (this.g_pth_Settings.favoritePaths.Length > 0) {
            InsertMenuItem(CurrentLocations, "Favorites", unset, unset, unset, unset) ; Header
            for favoritePath in this.g_pth_Settings.favoritePaths {
                InsertMenuItem(CurrentLocations, this.g_pth_settings.standardEntryPrefix favoritePath, favoritePath, A_WinDir . "\system32\imageres.dll", "-1024", false) ; Favorite Path
                hasItems := true
            }
        }

        ; Loop through the conditional favorites once and group them into a new object so we don't have to loop them all every time
        groupedConditionalFavorites := Map()
        for conditionalFavorite in this.g_pth_Settings.conditionalFavorites {
            if !groupedConditionalFavorites.Has(conditionalFavorite.conditionType)
                groupedConditionalFavorites[conditionalFavorite.conditionType] := []
            groupedConditionalFavorites[conditionalFavorite.conditionType].Push(conditionalFavorite)
        }

        for conditionType in ExplorerDialogPathSelector.ConditionType.OwnProps() {
            local stringToMatch := ""
            local sectionHeader := ""

            if (conditionType = "__Init" || conditionType = "Prototype") {
                continue
            } else if (conditionType = ExplorerDialogPathSelector.ConditionType.DialogOwnerExe.StringID) {
                itemToMatch := windowExe
                sectionHeader := "Conditional Favorites - Executable Match"
            } else if (conditionType = ExplorerDialogPathSelector.ConditionType.CurrentDialogPath.StringID) {
                itemToMatch := windowPath
                sectionHeader := "Conditional Favorites - Path Match"
            } else if (conditionType = ExplorerDialogPathSelector.ConditionType.ParentWindowTitle.StringID) {
                itemToMatch := parentWindowTitle
                sectionHeader := "Conditional Favorites - Parent Match"
            } else if (conditionType = ExplorerDialogPathSelector.ConditionType.DialogFileTypeFilter.StringID) {
                itemToMatch := windowFilter
                sectionHeader := "Conditional Favorites - File Filter Match"
            }

            ; Enter the group if it exists
            if groupedConditionalFavorites.Has(conditionType) {
                currentGroup := groupedConditionalFavorites[conditionType]
                ; Per-Group loop
                for favorite in currentGroup {
                    ; Users can set multiple conditions match strings per conditional favorite entry
                    for conditionValue in favorite.ConditionValues {
                        if (this.StringMatchWithWildcards(itemToMatch, conditionValue)) {
                            InsertMenuItem(CurrentLocations, sectionHeader, unset, unset, unset, unset) ; Header
                            ; Users can set multiple paths to show per conditional favorite entry upon a match
                            for conditionPath in favorite.Paths {
                                InsertMenuItem(CurrentLocations, this.g_pth_settings.standardEntryPrefix conditionPath, conditionPath, A_WinDir . "\system32\imageres.dll", "-81", false) ; Conditional Favorite Path
                                hasItems := true
                            }
                            break ; No need to keep going if we found a match
                        }
                    }
                }
            }            
        }

        ; Only get Directory Opus paths if dopusRTPath is set
        if (this.g_pth_Settings.dopusRTPath != "") {
            ; Get paths from Directory Opus using DOpusRT
            paths := this.GetDOpusPaths()

            ; Group paths by lister
            listersMap := Map()

            ; Get the ID of the active lister
            activeListerID := ""
            for pathObj in paths {
                if (pathObj.isActiveLister) {
                    activeListerID := pathObj.lister
                    break
                }
            }

            ; First, group all paths by their lister
            for pathObj in paths {
                if !listersMap.Has(pathObj.lister)
                    listersMap[pathObj.lister] := []
                listersMap[pathObj.lister].Push(pathObj)
            }

            ; Add the listers to a new list variable starting with the active lister, then the rest
            listers := []
            ; Add the active lister first if it exists
            if (activeListerID != "" and listersMap.Has(activeListerID)) {
                listers.Push(listersMap[activeListerID])
                listersMap.Delete(activeListerID)
            }
            ; Add any other listers
            for lister, listerPaths in listersMap {
                listers.Push(listerPaths)
            }
            
            windowNum := 1
            for listerPaths in listers {
                ; Add a separator if we had favorites and have Directory Opus paths to show
                if (hasItems and paths.Length > 0)
                    InsertMenuItem(CurrentLocations, "", unset, unset, unset, unset) ; Separator
                
                headerText := "Opus Window " windowNum
                InsertMenuItem(CurrentLocations, headerText, unset, unset, unset, unset) ; Header

                ; Add all paths for this lister
                for pathObj in listerPaths {
                    menuText := pathObj.path
                    ; Add prefix and suffix for active tab based on global settings
                    if (pathObj.isActiveTab)
                        menuText := this.g_pth_Settings.activeTabPrefix menuText this.g_pth_Settings.activeTabSuffix
                    else
                        menuText := this.g_pth_Settings.standardEntryPrefix menuText

                    InsertMenuItem(CurrentLocations, menuText, pathObj.path, A_WinDir . "\system32\imageres.dll", "4", pathObj.isActiveTab) ; Path
                    hasItems := true
                }

                windowNum++
            }
        }

        explorerPaths := this.GetAllExplorerPaths()

        ; Group paths by window handle (Hwnd)
        windows := Map()
        for pathObj in explorerPaths {
            if !windows.Has(pathObj.Hwnd)
                windows[pathObj.Hwnd] := []
            windows[pathObj.Hwnd].Push(pathObj)
        }

        ; Add Explorer paths if any exist
        if explorerPaths.Length > 0 {
            ; Add separator if we had Directory Opus paths
            if (hasItems)
                InsertMenuItem(CurrentLocations, "", unset, unset, unset, unset) ; Separator

            windowNum := 1
            for hwnd, windowPaths in windows {
                if (this.g_pth_Settings.groupPathsByWindow){
                    InsertMenuItem(CurrentLocations, "Explorer Window " windowNum, unset, unset, unset, unset) ; Header
                }

                for pathObj in windowPaths {
                    menuText := pathObj.Path
                    ; Add prefix and suffix for active tab based on global settings
                    if (pathObj.IsActive)
                        menuText := this.g_pth_Settings.activeTabPrefix menuText this.g_pth_Settings.activeTabSuffix
                    else
                        menuText := this.g_pth_Settings.standardEntryPrefix menuText

                    InsertMenuItem(CurrentLocations, menuText, pathObj.path, A_WinDir . "\system32\imageres.dll", "4", pathObj.IsActive) ; Path
                    hasItems := true
                }

                windowNum++
            }
        }

        ; If there is a path in the clipboard, add it to the menu
        if DllCall("Shlwapi\PathIsDirectoryW", "Str", A_Clipboard) != 0 {
            ; Add separator if we had Directory Opus or Explorer paths
            if (hasItems)
                InsertMenuItem(CurrentLocations, "", unset, unset, unset, unset) ; Separator

            path := A_Clipboard
            menuText := this.g_pth_Settings.standardEntryPrefix path Chr(0x200B) ; Add zero-with space as a janky way to make the menu item unique so it doesn't overwrite the icon of previous items with same path
            InsertMenuItem(CurrentLocations, menuText, path, A_WinDir . "\system32\imageres.dll", "-5301", false) ; Path (Clipboard)
            hasItems := true
            
        } else if this.g_pth_Settings.alwaysShowClipboardmenuItem = true {
            if (hasItems)
                InsertMenuItem(CurrentLocations, "", unset, unset, unset) ; Separator

            menuText := this.g_pth_Settings.standardEntryPrefix "Paste path from clipboard"
            InsertMenuItem(CurrentLocations, menuText, "placeholder path text", A_WinDir . "\system32\imageres.dll", "-5301", false) ; Clipboard Placeholder entry. 'Path' won't be used so using placeholder text
            CurrentLocations.Disable(currentMenuNum "&")
            hasItems := true ; Still show the menu item even if clipboard is empty, if the user has set it to always show clipboard item
        }

        RemoveToolTip() {
            SetTimer(RemoveToolTip, 0)
            ToolTip()
        }

        ; Show menu if we have items, otherwise show tooltip
        if (hasItems) {
            CurrentLocations.Show()
        } else {
            ToolTip("No folders open")
            SetTimer(RemoveToolTip, 1000)
        }

        ; Clean up
        CurrentLocations := ""
    }

    ; CallbackArgsTest(args*) {
    ;     for i, arg in args {
    ;         if (!IsSet(arg)){
    ;             OutputDebug("`nArgument " i " is not set")
    ;         } else if (arg = "") {
    ;             OutputDebug("`nArgument " i " is empty")
    ;         } else if (IsObject(arg)) {
    ;             ; Output the type
    ;             OutputDebug("`nArgument " i ": " Type(arg))
    ;         }  else {
    ;             OutputDebug("`nArgument " i ": " arg)
    ;         }
    ;     }
    ;     return
    ; }

    ; Callback function to navigate to a path. The first 3 parameters are provided by the .Add method of the Menu object
    ; The other two parameters must be provided using .Bind when specifying this function as a callback!
    Navigate(ThisMenuItemName, ThisMenuItemPos, MyMenu, f_path, windowClass, windowID) {
        ; ------------------------- LOCAL FUNCTIONS -------------------------
        NavigateDialog(path, windowHwnd, dialogInfo) {
            
            if (dialogInfo.Type = "ModernDialog") {
                OutputDebug("`n`nModern dialog detected. Navigating using address bar.")
                NavigateUsingAddressbar(path, windowHwnd)
            } else if (dialogInfo.Type = "HasEditControl") {
                OutputDebug("`n`nLegacy dialog detected. Navigating using edit control.")
                WinActivate("ahk_id " windowID)
                WM_SETTEXT := 0x000C

                ; Get the initial text of the edit control
                initialText := ControlGetText(dialogInfo.ControlHwnd, "ahk_id " windowHwnd)

                ; Try checking for SysListView32 control to see if any files are selected, then deselect them
                try {
                    countSelected := ListViewGetContent("Count Selected", "SysListView321", "ahk_id " windowHwnd)
                    if (countSelected > 0){
                        ; Focus on the listview
                        ControlFocus("SysListView321", "ahk_id " windowHwnd)
                        ; Send Ctrl + Space to deselect
                        SendInput("^{Space}")
                    }
                } catch Error as err{
                    OutputDebug("Error or no files selected in ListView. Message: " err.Message)
                }
                ; Send the path to the edit control text box using SendMessage
                DllCall("SendMessage", "Ptr", dialogInfo.ControlHwnd, "UInt", WM_SETTEXT, "Ptr", 0, "Str", path) ; 0xC is WM_SETTEXT - Sets the text of the text box
                ; Tell the dialog to accept the text box contents, which will cause it to navigate to the path
                DllCall("SendMessage", "Ptr", windowHwnd, "UInt", 0x0111, "Ptr", 0x1, "Ptr", 0) ; command ID (0x1) typically corresponds to the IDOK control which represents the primary action button, whether it's labeled "Save" or "Open".

                ; Restore the initial text of the edit control - Usually this happens automatically but just in case
                if (initialText){
                    DllCall("SendMessage", "Ptr", dialogInfo.ControlHwnd, "UInt", WM_SETTEXT, "Ptr", 0, "Str", initialText)
                ; Or if there was no initial text and now the control has the path leftover in it, then just clear the control text
                } else if (ControlGetText(dialogInfo.ControlHwnd, "ahk_id " windowHwnd) = path) {
                    DllCall("SendMessage", "Ptr", dialogInfo.ControlHwnd, "UInt", WM_SETTEXT, "Ptr", 0, "Str", "")
                }

            } else if (dialogInfo.Type = "FolderBrowserDialog") {
                OutputDebug("`n`nFolderBrowserDialog detected. Navigating using legacy folder dialog.")
                this.NavigateLegacyFolderDialog(path, dialogInfo.ControlHwnd)
            }
        }
        
        NavigateUsingAddressbar(path, windowHwnd) {
            ; ----------------- Local functions -----------------
            local eachWaitTime := 25
            CheckAddressbarReadyAndNavigate(waitedMs := 0) {
                addressbarHwnd := ControlGetFocus("ahk_id " windowHwnd)
                addressBarClassNN := ControlGetClassNN(addressbarHwnd)

                ; Regex match if the address bar is an Edit control but not Edit1, which seems to always the file name box. But the address bar box might not always be Edit2
                if (addressBarClassNN != "Edit1" and addressBarClassNN ~= "Edit\d+") {
                    ControlSetText(path, addressBarClassNN, "ahk_id " windowHwnd)
                    ControlSend("{Enter}", addressBarClassNN, "ahk_id " windowHwnd)
                    ControlFocus("Edit1", "ahk_id " windowHwnd) ; Return focus to the file name box
                } else if (waitedMs <= 300) {
                    ; Try waiting a bit longer for the address bar to be ready
                    Sleep(eachWaitTime)
                    OutputDebug("`n`nWaiting for address bar to be ready. Waited " waitedMs "ms")
                    CheckAddressbarReadyAndNavigate(waitedMs + eachWaitTime)
                } else {
                    OutputDebug("`n`nAddress bar didn't match expected class name. Found ClassNN: " addressBarClassNN)
                }
            }
            ; ----------------- End of local functions -----------------
            ; Activate the window
            WinActivate("ahk_id " windowHwnd)
            ; Try getting the text from the Edit1 control
            originalFileName := ""
            try {
                originalFileName := ControlGetText("Edit1", "ahk_id " windowHwnd)
                ControlSetText("", "Edit1", windowHwnd) ; Clear the text box just in case something weird happens so it doesn't save the file prematurely
            } catch Error as err {
                OutputDebug("`nError getting or setting Edit1 control text. Message: " err.Message)
            }

            ; Move focus to Address Bar
            Send("!{" this.localstr.addressBarKey "}") ; For some reason doesn't seem to work sending to the window. So use Alt + D to focus the address bar (or whatever language key)
            Sleep(50)
            CheckAddressbarReadyAndNavigate()

            ; Restore the original file name if it was there
            if (originalFileName != "") {
                ControlSetText(originalFileName, "Edit1", "ahk_id " windowHwnd)
            } else {
                OutputDebug("`n`nOriginal file name not found or no file name box handle")
            }
        }

        GetDialogAddressbarHwnd(windowHwnd) {
            controls := WinGetControls(windowHwnd)
            for controlClassNN in controls {
                if (controlClassNN ~= "ToolbarWindow32") {
                    controlText := ControlGetText(controlClassNN, windowHwnd)
                    if (controlText ~= this.localstr.addressBar) { ; Look for "Address: " or language specific equivalent
                        controlHwnd := ControlGetHwnd(controlClassNN, windowHwnd)
                        return controlHwnd
                    }
                }
            }
            return ""
        }

        DetectDialogType(hwnd) {
            try {
                addressbarHwnd := GetDialogAddressbarHwnd(hwnd)
                if (addressbarHwnd) {
                    return { Type: "ModernDialog", ControlHwnd: addressbarHwnd }
                }
            } catch as err {
                OutputDebug("`n`nApparently not a modern dialog. Checking other dialog types. Message: " err.Message)
                ; Nothing just move on
            }

            ; Look for an "Edit1" control, which is typically the file name edit box in file dialogs
            try {
                hFileNameEdit := ControlGetHwnd("Edit1", "ahk_class #32770")
                return { Type: "HasEditControl", ControlHwnd: hFileNameEdit }
            } catch {
                ; Try to get the handle of the TreeView control
                try {
                    hTreeView := ControlGetHwnd("SysTreeView321", "ahk_class #32770")
                    return { Type: "FolderBrowserDialog", ControlHwnd: hTreeView }
                } catch {
                    ; Neither control found
                    return 0
                }
            }
        }
        ; ------------------------------------------------------------------------

        ; Return if there's nothing
        if (f_path = "")
            return

        if (windowClass = "ConsoleWindowClass") {
            WinActivate("ahk_id " windowID)
            SetKeyDelay(0)
            Send("{Esc}pushd " f_path "{Enter}")
            return
        }
        ; Java dialogs have a different structure
        else if (windowClass = "SunAwtDialog") {
            WinActivate("ahk_id " windowID)

            ; Send Alt + N to focus the file name edit box
            DllCall("keybd_event", "UChar", 0x12, "UChar", 0x38, "UInt", 0, "UPtr", 0) ; Alt Down
            Sleep(20)
            DllCall("keybd_event", "UChar", 0x4E, "UChar", 0x31, "UInt", 0, "UPtr", 0) ; N Down
            Sleep(20)
            DllCall("keybd_event", "UChar", 0x4E, "UChar", 0x31, "UInt", 2, "UPtr", 0) ; N up
            DllCall("keybd_event", "UChar", 0x12, "UChar", 0x38, "UInt", 2, "UPtr", 0) ; Alt up

            ; Select the text in the box and delete it because otherwise it appends it
            Sleep(50)
            Send("^a") 
            Sleep(20)
            Send("{Del}")
            ; Send the path. Add a backslash just in case - the dialog seems to accept it
            Sleep(20)
            Send(f_path "\")
            Send("{Enter}")
            Sleep(20)
            ; Remove the text in the box because it doesn't remove it automatically
            Send("^a") 
            Sleep(20)
            Send("{Delete}")
            return

        } else { ; Default: #32770 or other compatible dialog
            ; Check if it's a legacy dialog
            if (dialogInfo := DetectDialogType(windowID)) {
                ; Use the legacy navigation approach
                NavigateDialog(f_path, windowID, dialogInfo)
            } else {
                ; Assume modern dialog
                NavigateUsingAddressbar(f_path, windowID)
            }
            return
        }
    }

    NavigateLegacyFolderDialog(path, hTV) {
        ; Helper function to navigate to a node with the given text under the given parent item
        NavigateToNode(treeView, parentItem, nodeText, isDriveLetter := false) {
            ; Helper function to escape special regex characters in node text
            RegExEscape(str) {
                static chars := "[\^\$\.\|\?\*\+\(\)\{\}\[\]\\]"
                return RegExReplace(str, chars, "\$0")
            }

            treeView.Expand(parentItem, true)
            hItem := treeView.GetChild(parentItem)
            while (hItem) {
                itemText := treeView.GetText(hItem)
                if (isDriveLetter) {
                    ; Special handling for drive letters. Look for them in parentheses, because they might show with name like "Primary (C:)"
                    if (itemText ~= "i)\(" . RegExEscape(nodeText) . "\)") {
                        ; Found the drive
                        return hItem
                    }
                } else {
                    ; Regular matching for other nodes
                    if (itemText ~= "i)^" . RegExEscape(nodeText) . "(\s|$)") {
                        ; Found the item
                        return hItem
                    }
                }
                hItem := treeView.GetNext(hItem)
            }
            return 0
        }

        ; -------------------------------------------------------------

        ; Initialize variables
        networkPath := ""
        driveLetter := ""
        hItem := ""

        ; Create RemoteTreeView object
        myTreeView := RemoteTreeView(hTV)

        ; Wait for the TreeView to load
        myTreeView.Wait()

        ; Split the path into components
        pathComponents := StrSplit(path, "\")
        ; If it's not english get the localized versions of path parts
        if (!this.localstr.isEnglish) {
            pathComponents := this.GetLocalizedPathParts(pathComponents)
        }

        ; Remove empty components caused by leading backslashes
        while (pathComponents.Length > 0 && pathComponents[1] = "") {
            pathComponents.RemoveAt(1)
        }

        ; Handle network paths starting with "\\"
        if (SubStr(path, 1, 2) = "\\") {
            networkPath := "\\" . pathComponents.RemoveAt(1)
            if pathComponents.Length > 0 {
                networkPath .= "\" . pathComponents.RemoveAt(1)
            }
        }

        ; Start from the "This PC" node (adjust for different Windows versions)
        startingNodes := [this.localstr.thisPC, this.localstr.computer, this.localstr.myComputer, this.localstr.desktop]
        for name in startingNodes {
            if (hItem := myTreeView.GetHandleByText(name)) {
                break
            }
        }
        if !hItem {
            MsgBox("Could not find a starting node like 'This PC' in the TreeView.")
            return
        }

        ; Expand the starting node
        myTreeView.Expand(hItem, true)

        ; If it's a network path
        if (networkPath != "") {
            ; Navigate to the network location
            hItem := NavigateToNode(myTreeView, hItem, networkPath)
            if !hItem {
                MsgBox("Could not find network path '" . networkPath . "' in the TreeView.")
                return
            }
        } else if (pathComponents.Length > 0 && pathComponents[1] ~= "^[A-Za-z]:$") {
            ; Handle drive letters
            driveLetter := pathComponents.RemoveAt(1)
            hItem := NavigateToNode(myTreeView, hItem, driveLetter, true) ; Pass true to indicate drive letter
            if !hItem {
                MsgBox("Could not find drive '" . driveLetter . "' in the TreeView.")
                return
            }
        } else {
            ; If path starts from a folder under starting node
            ; No action needed
        }

        ; Now navigate through the remaining components
        for component in pathComponents {
            hItem := NavigateToNode(myTreeView, hItem, component)
            if !hItem {
                MsgBox("Could not find folder '" . component . "' in the TreeView.")
                return
            }
        }

        ; Select the final item
        myTreeView.SetSelection(hItem, false)
        ; Optionally, send Enter to confirm selection
        ; Send("{Enter}")
    }

    ; ----------------------------------------------------------------------------------------------
    ; ---------------------------------------- GUI-RELATED  ----------------------------------------
    ; ----------------------------------------------------------------------------------------------

    ; Function to show the settings GUI
    ShowPathSelectorSettingsGUI(*) {
        ; ---------------------------------------- LOCAL FUNCTIONS ----------------------------------------
        BrowseForDopusRT(editControl) {
            selectedFile := FileSelect(3, unset, "Select dopusrt.exe", "Executable (*.exe)")
            if selectedFile
                editControl.Value := selectedFile
        }
        ; --------------------------------------------------------------------------------------------------

        ; Create the settings window
        settingsGui := Gui("+Resize", g_pathSelector_programName " - Settings")
        settingsGui.OnEvent("Size", GuiResize)
        settingsGui.SetFont("s10", "Segoe UI")

        hTT := ExplorerDialogPathSelector.CreateTooltipControl(settingsGui.Hwnd)

        ; Set Hotkey - Edit Box
        labelHotkey := settingsGui.AddText("xm y10 w120 h23 +0x200", "Menu Hotkey:")
        hotkeyEdit := settingsGui.AddEdit("x+10 yp w200", this.g_pth_Settings.dialogMenuHotkey)
        hotkeyEdit.SetFont("s12", "Consolas")
        labelhotkeyTooltipText := "Enter the key or key combination that will trigger the dialog menu (Using AutoHotkey V2 syntax)`n`nSee `"Help`" menu for link to documentation about key names."
        labelhotkeyTooltipText .= "`n`nTip: Add a tilde (~) before the key to ensure the hotkey doesn't block the key's normal functionality.`nExample:  ~MButton"
        ExplorerDialogPathSelector.AddTooltipToControl(hTT, labelHotkey.Hwnd, labelhotkeyTooltipText)
        ExplorerDialogPathSelector.AddTooltipToControl(hTT, hotkeyEdit.Hwnd, labelhotkeyTooltipText)

        ; DOpus RT Path - Edit Box
        labelOpusRTPath := settingsGui.AddText("xm y+10 w120 h23 +0x200", "DOpus RT Path:")
        dopusPathEdit := settingsGui.AddEdit("x+10 yp w200 h30 -Multi -Wrap", this.g_pth_Settings.dopusRTPath) ; Setting explicit height and -Multi because for some reason it was wrapping the control box down. Not sure if -Wrap is necessary
        labelOpusRTPathTooltipText := "*** For Directory Opus users *** `nPath to dopusrt.exe`n`nOr leave empty to disable Directory Opus integration."
        ExplorerDialogPathSelector.AddTooltipToControl(hTT, labelOpusRTPath.Hwnd, labelOpusRTPathTooltipText)
        ExplorerDialogPathSelector.AddTooltipToControl(hTT, dopusPathEdit.Hwnd, labelOpusRTPathTooltipText)
        ; DOpusRT Browse - Button
        browseBtn := settingsGui.AddButton("x+5 yp w60", "Browse...")
        browseBtn.OnEvent("Click", (*) => BrowseForDopusRT(dopusPathEdit))

        ; Set Active Prefix - Edit Box
        labelActiveTabPrefix := settingsGui.AddText("xm y+10 w120 h23 +0x200", "Active Tab Prefix:")
        prefixEdit := settingsGui.AddEdit("x+10 yp w200", this.g_pth_Settings.activeTabPrefix)
        labelActiveTabPrefixTooltipText := "Text/Characters that appears to the left of the active path for each window group"
        ExplorerDialogPathSelector.AddTooltipToControl(hTT, labelActiveTabPrefix.Hwnd, labelActiveTabPrefixTooltipText)
        ExplorerDialogPathSelector.AddTooltipToControl(hTT, prefixEdit.Hwnd, labelActiveTabPrefixTooltipText)

        ; Set Standard Prefix - Edit Box
        labelStandardEntryPrefix := settingsGui.AddText("xm y+10 w120 h23 +0x200", "Non-Active Prefix:")
        standardPrefixEdit := settingsGui.AddEdit("x+10 yp w200", this.g_pth_Settings.standardEntryPrefix)
        labelStandardEntryPrefixTooltipText := "Indentation spaces for inactive tabs, so they line up"
        ExplorerDialogPathSelector.AddTooltipToControl(hTT, labelStandardEntryPrefix.Hwnd, labelStandardEntryPrefixTooltipText)
        ExplorerDialogPathSelector.AddTooltipToControl(hTT, standardPrefixEdit.Hwnd, labelStandardEntryPrefixTooltipText)

        ; Set Active Suffix - Edit Box
        ; labelActiveTabSuffix := settingsGui.AddText("xm y+10 w120 h23 +0x200", "Active Tab Suffix:")
        ; suffixEdit := settingsGui.AddEdit("x+10 yp w200", g_settings.activeTabSuffix)
        ; labelActiveTabSuffixTooltipText := "Text/Characters will appear to the right of the active path for each window group, if you want as a label."
        ; AddTooltipToControl(hTT, labelActiveTabSuffix.Hwnd, labelActiveTabSuffixTooltipText)
        ; AddTooltipToControl(hTT, suffixEdit.Hwnd, labelActiveTabSuffixTooltipText)

        ; Bring up favorites setting GUI - Button
        favoritesBtn := settingsGui.AddButton("xm y+10 w120 h30", "Favorites")
        favoritesBtn.OnEvent("Click", (*) => this.ShowFavoritePathsGui())

        ; Bring up conditional favorites GUI - Button
        conditionalFavoritesBtn := settingsGui.AddButton("xp+130 yp+0 w150 h30", "Conditional Favorites")
        conditionalFavoritesBtn.OnEvent("Click", (*) => this.ShowConditionalFavoritesGui())
        
        ; Debug Mode - Checkbox
        debugCheck := settingsGui.AddCheckbox("xm y+15", "Enable Debug Mode")
        debugCheck.Value := this.g_pth_Settings.enableExplorerDialogMenuDebug
        labelDebugCheckTooltipText := "Show tooltips with debug information when the hotkey is pressed.`nUseful for troubleshooting."
        ExplorerDialogPathSelector.AddTooltipToControl(hTT, debugCheck.Hwnd, labelDebugCheckTooltipText)

        ; Always Show Clipboard Menu Item - Checkbox
        clipboardCheck := settingsGui.AddCheckbox("xm y+5", "Always Show Clipboard Menu Item")
        clipboardCheck.Value := this.g_pth_Settings.alwaysShowClipboardmenuItem
        labelClipboardCheckTooltipText := "If Disabled: The option to paste the clipboard path will only appear when a valid path is found on the clipboard.`nIf Enabled: The menu entry will always appear, but is disabled when no valid path is found."
        ExplorerDialogPathSelector.AddTooltipToControl(hTT, clipboardCheck.Hwnd, labelClipboardCheckTooltipText)

        ; Group Paths By Window - Checkbox
        groupByWindowCheck := settingsGui.AddCheckbox("xm y+5", "Group Explorer Paths By Window")
        groupByWindowCheck.Value := this.g_pth_Settings.groupPathsByWindow
        labelGroupByWindowCheckTooltipText := "If disabled, paths from Windows Explorer will all be listed together without grouping by window.`nIdeal for Windows 10 which does not have tabs."
        ExplorerDialogPathSelector.AddTooltipToControl(hTT, groupByWindowCheck.Hwnd, labelGroupByWindowCheckTooltipText)

        ; Use bold font for active tab - Checkbox
        useBoldTextActiveCheck := settingsGui.AddCheckbox("xm y+5", "Use Bold Font for Active Tabs")
        useBoldTextActiveCheck.Value := this.g_pth_Settings.useBoldTextActive
        useBoldTextActiveCheckTooltipText := "Use a bold font for the active tab in the menu.`nNote: It uses a simulated bold effect using unicode, `n      so it may not look perfect for all characters."
        ExplorerDialogPathSelector.AddTooltipToControl(hTT, useBoldTextActiveCheck.Hwnd, useBoldTextActiveCheckTooltipText)

        ; Enable UI Access - Checkbox
        UIAccessCheck := settingsGui.AddCheckbox("xm y+5", "Enable UI Access")
        UIAccessCheck.Value := this.g_pth_Settings.enableUIAccess
        labelUIAccessCheckTooltipText := ""
        if !this.ThisScriptRunningStandalone() or A_IsCompiled {
            UIAccessCheck.Value := 0
            UIAccessCheck.Enabled := 0

            ; Get position of the checkbox before disabling it so we can add an invisible box to apply the tooltip to
            ; Because the tooltip won't show on a disabled control
            x := 0, y := 0, w := 0, h := 0
            UIAccessCheck.GetPos(&x, &y, &w, &h)
            tooltipOverlay := settingsGui.AddText("x" x " y" y " w" w " h" h " +BackgroundTrans", "")

            if A_IsCompiled {
                labelUIAccessCheckTooltipText := "UI Access allows the script to work on dialogs run by elevated processes, without having to run as Admin itself."
                labelUIAccessCheckTooltipText .= "`nHowever this setting does not apply for the compiled Exe version of the script."
                labelUIAccessCheckTooltipText .= "`n`nInstead, you must put the exe in a `"trusted`" Windows location such as the `"C:\Program Files\...`" directory."
                labelUIAccessCheckTooltipText .= "`nYou do NOT need to run the exe as Admin for this to work."
            } else {
                labelUIAccessCheckTooltipText := "This script appears to be running as being included by another script. You should enable UI Access via the parent script instead."
            }
            ExplorerDialogPathSelector.AddTooltipToControl(hTT, tooltipOverlay.Hwnd, labelUIAccessCheckTooltipText)
        } else {
            labelUIAccessCheckTooltipText := "Enable `"UI Access`" to allow the script to run in elevated windows protected by UAC without running as admin."
            ExplorerDialogPathSelector.AddTooltipToControl(hTT, UIAccessCheck.Hwnd, labelUIAccessCheckTooltipText)
        }

        ; Add divider line for non-persistent settings
        checkBoxDivider := settingsGui.AddText("xm y+15 h2 w150 0x10")

        ; Keep Open After Saving - Checkbox
        keepOpenCheck := settingsGui.AddCheckbox("xm y+10", "Keep This Window Open After Saving")
        keepOpenCheck.SetFont("s9") ; Smaller font for this checkbox
        keepOpenCheck.Value := false ; False by default - This isn't a saved setting, just a temporary preference
        keepOpenCheck.OnEvent("Click", (*) => ToggleAlwaysOnTopCheckVisibility(keepOpenCheck.Value))
        labelKeepOpenCheckTooltipText := "Keep this window open after saving the settings.`nGood for experimenting with different settings.`n`n(Note: This checkbox setting is not saved.)"
        ExplorerDialogPathSelector.AddTooltipToControl(hTT, keepOpenCheck.Hwnd, labelKeepOpenCheckTooltipText)

        ; Keep Settings on Top - Checkbox (Hidden by default and shown only if keepOpenCheck is checked)
        keepOnTopCheck := settingsGui.AddCheckbox("xm y+5 +Hidden1", "Keep This Window Always On Top") ; +Hidden hides by default, but setting +Hidden1 to make it explicitly hidden
        keepOnTopCheck.SetFont("s9") ; Smaller font for this checkbox
        keepOnTopCheck.Value := false ; False by default - This isn't a saved setting, just a temporary preference
        keepOnTopCheck.OnEvent("Click", (*) => SetSettingsWindowAlwaysOnTop(keepOnTopCheck.Value))
        labelKeepOnTopCheckTooltipText := "Keep this window always on top of other windows.`nGood for keeping it visible while testing settings.`n`n(Note: This checkbox setting is not saved.)"
        ExplorerDialogPathSelector.AddTooltipToControl(hTT, keepOnTopCheck.Hwnd, labelKeepOnTopCheckTooltipText)

        
        ; --------- Bottom Buttons ---------- See positioning cheatsheet: https://www.reddit.com/r/AutoHotkey/comments/1968fq0/a_cheatsheet_for_building_guis_using_relative/
        buttonsY := "y+20"
        ; Reset button
        resetBtn := settingsGui.AddButton("xm " buttonsY " w80", "Defaults")
        resetBtn.OnEvent("Click", ResetSettings)
        settingsGui.AddButton("x+10 yp w80", "Cancel").OnEvent("Click", (*) => settingsGui.Destroy())
        labelButtonResetTooltipText := "Sets all settings above to their default values.`nYou'll still need to click Save to apply the changes."
        ExplorerDialogPathSelector.AddTooltipToControl(hTT, resetBtn.Hwnd, labelButtonResetTooltipText)
        ; Save button
        saveBtn := settingsGui.AddButton("x+10 yp w80 Default", "Save")
        saveBtn.OnEvent("Click", SaveSettings)
        labelButtonSaveTooltipText := "Save the current settings to a file to automatically load in the future."
        ExplorerDialogPathSelector.AddTooltipToControl(hTT, saveBtn.Hwnd, labelButtonSaveTooltipText)
        ; Help button
        helpBtn := settingsGui.AddButton("x+10 w70", "Help")
        helpBtn.OnEvent("Click", this.ShowPathSelectorHelpWindow_Bound)
        ; Smaller About button above the help button
        aboutBtn := settingsGui.AddButton("xp+0 yp-25 w70 h25", "About") ; These positions aren't right, but we'll fix them in the resize function
        aboutBtn.OnEvent("Click", ExplorerDialogPathSelector.ShowPathSelectorAboutWindow)

        ; Set variables to track when certain settings are changed for special handling.
        ; Setting this as a function so it can be called again if saving settings without closing the window
        UIAccessInitialValue := ""
        HotkeyInitialValue := ""
        RecordInitialValuesFromGlobalSettings() {
            UIAccessInitialValue := this.g_pth_Settings.enableUIAccess
            HotkeyInitialValue := this.g_pth_Settings.dialogMenuHotkey
        }
        RecordInitialValuesFromGlobalSettings()

        ; Show the GUI
        settingsGui.Show()

        ResetSettings(*) {
            hotkeyEdit.Value := this.DefaultSettings.dialogMenuHotkey
            dopusPathEdit.Value := this.DefaultSettings.dopusRTPath
            prefixEdit.Value := this.DefaultSettings.activeTabPrefix
            ;suffixEdit.Value := DefaultSettings.activeTabSuffix
            useBoldTextActiveCheck := this.DefaultSettings.useBoldTextActive
            standardPrefixEdit.Value := this.DefaultSettings.standardEntryPrefix
            debugCheck.Value := this.DefaultSettings.enableExplorerDialogMenuDebug
            clipboardCheck.Value := this.DefaultSettings.alwaysShowClipboardmenuItem
            groupByWindowCheck.Value := this.DefaultSettings.groupPathsByWindow
            UIAccessCheck.Value := this.DefaultSettings.enableUIAccess
            ; Favorites paths are updated in the favorites GUI
        }

        SaveSettings(*) {
            ; Update settings object
            this.g_pth_Settings.dialogMenuHotkey := hotkeyEdit.Value
            this.g_pth_Settings.dopusRTPath := dopusPathEdit.Value
            this.g_pth_Settings.activeTabPrefix := prefixEdit.Value
            ;g_settings.activeTabSuffix := suffixEdit.Value
            this.g_pth_Settings.useBoldTextActive := useBoldTextActiveCheck.Value
            this.g_pth_Settings.standardEntryPrefix := standardPrefixEdit.Value
            this.g_pth_Settings.enableExplorerDialogMenuDebug := debugCheck.Value
            this.g_pth_Settings.alwaysShowClipboardmenuItem := clipboardCheck.Value
            this.g_pth_Settings.groupPathsByWindow := groupByWindowCheck.Value
            this.g_pth_Settings.enableUIAccess := UIAccessCheck.Value
            ; Favorites paths are updated in the favorites GUI

            this.SaveSettingsToFile()

            ; When UI Access goes from enabled to disabled, the user must manually close and re-run the script
            if (UIAccessInitialValue = true && UIAccessCheck.Value = false) {
                MsgBox("NOTE: When changing UI Access from Enabled to Disabled, you must manually close and re-run the script/app for changes to take effect.", "Settings Saved - Process Restart Required", "Icon!")
            } else if (UIAccessInitialValue = false && UIAccessCheck.Value = true) {
                ; When enabling UI Access, we can reload the script to enable it. Ask the user if they want to do this now
                result := MsgBox("UI Access has been enabled. Do you want to restart the script now to apply the changes?", "Settings Saved - Process Restart Required", "YesNo Icon!")
                if (result = "Yes") {
                    Reload()
                }
            }

            ; The rest of the settings don't require a restart, they are pulled directly from the settings object which has been updated

            ; Disable the original hotkey by passing in the previous hotkey string
            this.UpdateHotkey("", HotkeyInitialValue)

            ; At this point all settings have been saved and applied
            RecordInitialValuesFromGlobalSettings()

            if (keepOpenCheck.Value = false) {
                settingsGui.Destroy()
            }
        }

        GuiResize(thisGui, minMax, width, height) {
            if minMax = -1  ; The window has been minimized
                return

            ; Update control positions based on new window size
            for ctrl in thisGui {
                ; For specific control objects
                if ctrl = checkBoxDivider {
                    ctrl.Move(unset, unset, width - 30)  ; Set width to fill the window
                    continue
                }

                ; For certain control types
                if ctrl.HasProp("Type") {
                    if ctrl.Type = "Edit" {
                        ; Leave space for the Browse button if this is the DOpus path edit box
                        if (ctrl.HasProp("ClassNN") && ctrl.ClassNN = "Edit2") {
                            ctrl.Move(unset, unset, width - 220)  ; Leave extra space for Browse button
                        } else {
                            ctrl.Move(unset, unset, width - 150)  ; Set consistent width for other edit boxes
                        }
                    } else if ctrl.Type = "Button" {
                        if ctrl.Text = "Browse..." {
                            ctrl.Move(width - 70)  ; Anchor Browse button to right side window edge
                        } else if ctrl.Text = "Help" {
                            ctrl.Move(width-85, height-45)  ; Right align Help button with some margin
                        } else if ctrl.Text = "About" {
                            ; Align it with the Help button and go above it
                            ctrl.Move(width-85, height-75)
                        } else if ctrl.Text = "Save" or ctrl.Text = "Defaults" or ctrl.Text = "Cancel" {
                            ctrl.Move(unset, height-45)  ; Bottom align buttons with 40px margin from bottom
                        }
                        ctrl.Redraw()
                    }
                }
            }
        }

        SetSettingsWindowAlwaysOnTop(checkValue) {
            if (checkValue) {
                settingsGui.Opt("+AlwaysOnTop")
            } else {
                settingsGui.Opt("-AlwaysOnTop")
            }
        }

        ToggleAlwaysOnTopCheckVisibility(keepOpenCheckValue) {
            if (keepOpenCheckValue) {
                keepOnTopCheck.Visible := true
            } else {
                ; Hide the checkbox but only if it's not checked
                if !keepOnTopCheck.Value {
                    keepOnTopCheck.Visible := false
                }
            }
        }
    }

    ShowConditionalFavoritesGui(*) {
        ; Create a deep copy of the settings instead of a reference
        pendingConditionalFavorites := []
        for entry in this.g_pth_settings.conditionalFavorites {
            pendingConditionalFavorites.Push({
                Index: entry.Index,
                ConditionType: entry.ConditionType,
                ConditionTypeName: entry.ConditionTypeName,
                ConditionValues: entry.ConditionValues,
                Paths: entry.Paths
            })
        }

        currentlyEditedRow := -1

        ; Automatically build a Map/Dictionary of the condition types as keys and incremental integer as an index for the dropdown
        local ConditionTypeIndexDict := Map()
        local dropdownFriendlyArray := []
        local iCondition := 1
        for conditionTypeStr in ExplorerDialogPathSelector.ConditionType.OwnProps() {
            ; Skip built in properties
            if (conditionTypeStr = "__Init" || conditionTypeStr = "Prototype")
                continue
            ConditionTypeIndexDict[conditionTypeStr] := iCondition
            dropdownFriendlyArray.Push(ExplorerDialogPathSelector.ConditionType.%conditionTypeStr%.FriendlyName)
            iCondition++
        }

        currentRowNum := 0
        
        ; Create the main window
        pathGui := Gui("+Resize +MinSize600x500", "Conditional Favorites Manager")
        pathGui.OnEvent("Size", GuiResize)
        pathGui.SetFont("s10", "Segoe UI")
        
        ; Add ListView to show existing conditions
        pathGui.AddText("w580", "Conditional Favorites Rules:")
        listView := pathGui.AddListView("w580 h200 vConditionsList", ["#", "Condition Type", "Condition Values", "Paths"])

        ; Add label to show which is the currently selected entry
        labelActiveSelection := pathGui.AddText("w580", "Currently Editing: [None]")
        
        ; Add buttons for managing entries - Add events separately so we can create a new object variable for each button
        addBtn := pathGui.AddButton("w80", "Add New")
        addBtn.OnEvent("Click", AddEntry)
        editBtn := pathGui.AddButton("x+10 yp w80", "Edit")
        editBtn.OnEvent("Click", EditOrCancelEditEntry)
        removeBtn := pathGui.AddButton("x+10 yp w80", "Remove")
        removeBtn.OnEvent("Click", RemoveEntry)
        helpBtn := pathGui.AddButton("x+10 yp w80", "Help")
        helpBtn.OnEvent("Click", this.ShowConditionalFavoritesHelpWindow)
        
        ; Add edit panel (initially disabled)
        grpBox := pathGui.AddGroupBox("xs w580 h220", "Entry Details")
        
        ; Condition Type dropdown
        pathGui.AddText("xp+10 yp+20", "Condition Value Type:")
        typeDropdown := pathGui.AddDropDownList("w200 vConditionType Choose1", dropdownFriendlyArray)
        typeDropdown.OnEvent("Change", ShowConditionTypeDescription)
        ; Condition Type Description text - Shows next to the dropdown
        typeDescription := pathGui.AddText("x+10 yp+10 w150 +Wrap", ExplorerDialogPathSelector.ConditionType.CurrentDialogPath.Description)
        ShowConditionTypeDescription()  ; Show description for the default selected type
        
        ; Condition Values
        conditionValuesLabel := pathGui.AddText("xs+10 yp+40", "Condition Values (one per line):")
        valuesEdit := pathGui.AddEdit("w560 h60 vConditionValues Multi VScroll", "")
        
        ; Paths
        pathsEditLabel := pathGui.AddText("xs+10 y+10", "Paths to show when any of the condition values match (one per line):")
        pathsEdit := pathGui.AddEdit("w560 h60 vPaths Multi VScroll", "")
        
        ; Main buttons
        saveBtn := pathGui.AddButton("xs yp+15 w80", "Save")
        saveBtn.OnEvent("Click", SaveAndClose)
        cancelBtn := pathGui.AddButton("x+10 yp w80", "Cancel")
        cancelBtn.OnEvent("Click", (*) => pathGui.Destroy())
        applyBtn := pathGui.AddButton("x+10 yp w80", "Apply")
        applyBtn.OnEvent("Click", ValidateAndApplyEntry)
            
        ; Initially disable edit panel controls
        EnableDisableEditPanel(false, -1)
        local isEditing := false

        ; Set initial sizes
        initialWindowWidth := 600
        defaultValColWidth := 120 ; Note that this will be adjusted automatically for scaling

        ; Properly set the height of the groupbox to fit all controls
        grpBox.GetPos(unset, &grpBoxY, unset, &grpBoxHeight)
        pathsEdit.GetPos(unset, &pathEditY, unset, &pathEditHeight)
        totalHeight := pathEditY + pathEditHeight - grpBoxY + 20
        grpBox.Move(unset, unset, initialWindowWidth - 20, totalHeight) ; Set width to fill the window

        ; Set initial height to fit everything
        grpBox.GetPos(unset, &grpBoxY, unset, &grpBoxHeight)
        initialHeight := grpBoxY + grpBoxHeight + 75

        PopulateListView()
        
        ; Show the GUI
        pathGui.Show("w" initialWindowWidth " h" initialHeight)

        ; ------------------------- LOCAL FUNCTIONS -------------------------

        GetDpiScaleFactor(){
            return A_ScreenDPI / 96
        }

        GetListViewColumnWidth(columnIndex) {
            static LVM_GETCOLUMNWIDTH := 0x101D
            try {
                return SendMessage(LVM_GETCOLUMNWIDTH, columnIndex - 1, 0, listView.Hwnd, unset, unset, unset, unset, 75)
            } catch {
                ; If it times out or fails return default width
                return defaultValColWidth
            }
        }

        IncreaseRelativeColumnWidth(columnIndex, increaseAmount) {
            currentWidth := GetListViewColumnWidth(columnIndex) / GetDpiScaleFactor() ; Need to use the unscaled width
            listView.ModifyCol(columnIndex, currentWidth + increaseAmount)
        }

        /**
         * Recreate the listview with the condition table.
         * @param {Integer|null} rowToSelect If unset or 0, will not select any. If a number is set, that row will be selected (1-index based). If -1, will select the last item.
         */
        PopulateListView(rowToSelect := unset) {
            listView.Delete()

            ; Create strings to show in ListView
            valueString := ""
            pathString := ""
            for entry in pendingConditionalFavorites {
                valueString := JoinDelimited(entry.ConditionValues, "; ")
                pathString := JoinDelimited(entry.Paths, "; ")

                if (IsSet(rowToSelect) and (rowToSelect = -1) and (entry.Index = pendingConditionalFavorites.Length))
                    listView.Add("Select", entry.Index, entry.ConditionTypeName, valueString, pathString)
                else if (IsSet(rowToSelect) and (rowToSelect = entry.Index))
                    listView.Add("Select", entry.Index, entry.ConditionTypeName, valueString, pathString)
                else
                    listView.Add(unset, entry.Index, entry.ConditionTypeName, valueString, pathString)                
            }

            ; Set column widths
            listView.ModifyCol(1, "Auto")  ; Index - AutoHdr to auto-size based on column contents. AutoHdr would be header text
            listView.ModifyCol(2, "Auto") ; Condition Type - Auto to auto-size based on content since there are only 2 options
            listView.ModifyCol(3, defaultValColWidth) ; Condition Values column - Fixed starting size auto-size based on content, then adjust further if necessary
            listView.ModifyCol(4, "Auto") ; Paths - Just expand as big as necessary. It will auto overflow to the right and user can expand window or scroll

            IncreaseRelativeColumnWidth(1, 3)
            IncreaseRelativeColumnWidth(2, 5)

            ; Set optimal column width for the values column. It's really the only one that needs more special handling since paths will just overflow to the right, and other columns are fixed size
            valuesColumnWidth := GetListViewColumnWidth(3)
            if (valuesColumnWidth < defaultValColWidth) {
                listView.ModifyCol(3, defaultValColWidth) ; Set to 180 if it's less than that
            }

            ; listView.Redraw()
        }

        ; Helper function to show the description for the selected condition type
        ShowConditionTypeDescription(*) {
            dropdownIndex := typeDropdown.Value
            for key, value in ConditionTypeIndexDict {
                if (value = dropdownIndex) {
                    ; Since 'key' will be the StringID, we need to find the matching enum value, we can dynamically access properties with %% syntax.
                    typeDescription.Value := ExplorerDialogPathSelector.ConditionType.%key%.Description
                    break
                }
            }
            typeDescription.Redraw()
        }

        GetConditionTypeStringIDFromFriendlyName(friendlyName) {
            for key, typeObj in ExplorerDialogPathSelector.ConditionType.OwnProps() {  ; Use OwnProps() to get static properties
                if (typeObj.FriendlyName = friendlyName) {
                    return typeObj.StringID
                }
            }
            return ""
        }
        
        /**
         * Enable or disable the edit panel.
         * @overload
         * @param {true} enable True to enable the edit panel for a specific row.
         * @param {Number} numToEdit The row number to edit.
        */
        /**
         * Enable or disable the edit panel.
         * @overload
         * @param {false} enable False to disable the edit panel.
         * @param {-1} numToEdit When 
         */
        EnableDisableEditPanel(enable := true, numToEdit) {
            typeDropdown.Enabled := enable
            valuesEdit.Enabled := enable
            pathsEdit.Enabled := enable

            if (enable) {
                isEditing := true
                editBtn.Text := "Cancel Edit"

                currentlyEditedRow := numToEdit
                labelActiveSelection.Value := "Currently Editing: #" numToEdit
                labelActiveSelection.SetFont("s10 Bold cRed")

                listView.Enabled := false
            } else {
                valuesEdit.Value := ""
                pathsEdit.Value := ""
                editBtn.Text := "Edit"
                isEditing := false

                currentlyEditedRow := -1
                labelActiveSelection.Value := "Currently Editing: [None]"
                labelActiveSelection.SetFont() ; Default same as pathGUI object
                labelActiveSelection.SetFont("cDefault") ; The above is supposed to reset the color too but it doesn't so we do it like this

                listView.Enabled := true
            } 
        }
   
        ; Handle adding new entry
        AddEntry(*) {
            typeDropdown.Value := ConditionTypeIndexDict[ExplorerDialogPathSelector.ConditionType.DialogOwnerExe.StringID] ; Initial default but can be changed by user
            valuesEdit.Value := ""
            pathsEdit.Value := ""
            ; Create new entry object in pendingConditionalFavorites - Use default placeholders
            entry := {
                Index : pendingConditionalFavorites.Length + 1,
                ConditionType: ExplorerDialogPathSelector.ConditionType.DialogOwnerExe.StringID,
                ConditionTypeName: ExplorerDialogPathSelector.ConditionType.DialogOwnerExe.FriendlyName,
                ConditionValues: [],
                Paths: []
            }
            pendingConditionalFavorites.Push(entry)
            ; Add new entry to ListView
            PopulateListView(-1)
            ; Update the current 
            EnableDisableEditPanel(true, entry.Index)
        }
        
        ; Handle editing selected entry
        EditOrCancelEditEntry(*) {
            currentRowNum := listView.GetNext()
            if (currentRowNum > 0)
                entry := pendingConditionalFavorites[currentRowNum] ; The index in the array is 0-based, but the ListView index is 1-based

            ; Already editing, so the user wants to cancel
            if (isEditing) {
                EnableDisableEditPanel(false, -1)
                ; If cancelling editing and the entry was empty to begin with, like after adding new, remove it again
                if (entry and entry.ConditionValues.Length = 0 and entry.Paths.Length = 0)
                    RemoveEntry()
            ; Enable Editing
            } else if (currentRowNum) {
                typeDropdown.Value := ConditionTypeIndexDict[entry.ConditionType]
                valuesEdit.Value := Join(entry.ConditionValues)
                pathsEdit.Value := Join(entry.Paths)
                EnableDisableEditPanel(true, entry.Index)
            }
        }
        
        ; Handle removing selected entry
        RemoveEntry(*) {
            if (row := listView.GetNext()) {
                listView.Delete(row)
                pendingConditionalFavorites.RemoveAt(row)
            }

            ; Update the indexes of the entries in case any were removed
            for i, entry in pendingConditionalFavorites {
                entry.Index := i ; Autohotkey is 1 index based
            }

            PopulateListView()
            EnableDisableEditPanel(false, -1)
        }

        ValidateAndApplyEntry(*) {
            if (typeDropdown.Enabled) {  ; If panel is enabled, include current values
                entry := {
                    ConditionType: GetConditionTypeStringIDFromFriendlyName(typeDropdown.Text),
                    conditionTypeName: typeDropdown.Text,
                    ConditionValues: SplitAndTrim(valuesEdit.Value),
                    Paths: SplitAndTrim(pathsEdit.Value),
                    Index: currentlyEditedRow
                }

                ; Validate values
                for value in entry.ConditionValues {
                    if (!this.ValidatePathCharacters_AllowWildCards(value)) {
                        MsgBox("Invalid characters found in value:`n" value "`n`nCannot contain these characters:`n< > : `" / | ? `n`nPlease correct and try again.", "Error", "Icon!")
                        return
                    }
                }
                
                ; Validate paths
                for path in entry.Paths {
                    if (!this.ValidatePathCharacters(path)) {
                        MsgBox("Invalid characters found in path:`n" path "`n`nCannot contain these characters:`n< > : `" / | * ? `n`nPlease correct and try again.", "Error", "Icon!")
                        return
                    }
                }
                
                ; Add to array if both values and paths are provided
                if (entry.ConditionValues.Length > 0 && entry.Paths.Length > 0) {
                    pendingConditionalFavorites[currentlyEditedRow] := entry
                } else {
                    MsgBox("Both Condition Values and Paths are required to save a conditional favorite.", "Error", "Icon!")
                    return
                }

                PopulateListView(currentlyEditedRow)
                EnableDisableEditPanel(false, -1)
            }
        }
    
        SaveAndClose(*) {
            ; If any entry is being edited, apply it
            if (currentlyEditedRow != -1) {
                ValidateAndApplyEntry()
            }

            ; Remove any pending conditional favorites that are empty
            for index, entry in pendingConditionalFavorites {
                if (entry.ConditionValues.Length = 0 || entry.Paths.Length = 0) {
                    pendingConditionalFavorites.RemoveAt(index)
                }
            }

            ; Update the indexes of the entries in case any were removed
            for i, entry in pendingConditionalFavorites {
                entry.Index := i ; Autohotkey is 1 index based
            }

            this.g_pth_Settings.conditionalFavorites := pendingConditionalFavorites
            pathGui.Destroy()
        }
        
        ; Helper function to split and trim text into array
        SplitAndTrim(text) {
            arr := []
            for line in StrSplit(text, "`n", "`r") {
                if (Trim(line) != "") {
                    arr.Push(Trim(line))
                }
            }
            return arr
        }
        
        ; Helper function to join array into multiline text
        Join(arr) {
            text := ""
            for item in arr {
                text .= item "`n"
            }
            return RTrim(text, "`n")
        }

        JoinDelimited(arr, delimiter) {
            text := ""
            for item in arr {
                text .= item delimiter
            }
            return RTrim(text, delimiter)
        }

        GetControlRightEdge(ctrl) {
            ctrl.GetPos(&ctrl_X, &ctrl_Y, &ctrlWidth, &ctrlHeight)
            return ctrl_X + ctrlWidth
        }

        GetControlBottomEdge(ctrl) {
            ctrl.GetPos(&ctrl_X, &ctrl_Y, &ctrlWidth, &ctrlHeight)
            return ctrl_Y + ctrlHeight
        }

        HeightToFill(ctrl, parentHeight) {
            ctrl.GetPos(&ctrl_X, &ctrl_Y, &ctrlWidth, &ctrlHeight)
            return parentHeight - ctrl_Y
        }

        GuiResize(thisGui, minMax, window_width, window_height) {
            if minMax = -1  ; The window has been minimized
                return

            ; ---------------------------------- Update Entry Details Controls Sizes ----------------------------------
            ; Update the groupbox height based on the window size
            grpBox.GetPos(&grpBox_X, &grpBox_Y, &grpBoxWidth, &grpBoxHeight)
            newGrpBoxHeight := HeightToFill(grpBox, window_height) - 60
            grpBox.Move(unset, unset, window_width - 30, newGrpBoxHeight)
            ; grpBox.Redraw()

            ; Update the text positions and sizes of the condition values and paths edit boxes. Need to be sure to keep their labels positioned correctly
            ; The two together should fill the remaining height of the groupbox with about half of the height each
            grpBoxAvailableHeight := newGrpBoxHeight - (GetControlBottomEdge(conditionValuesLabel) - grpBox_Y) - 50 ; Subtract out space for controls above the edit boxes and buffer between
            ; First do the values edit box and label
            grpBox.GetPos(&grpBox_X, &grpBox_Y, &grpBoxWidth, &grpBoxHeight)
            valuesEdit.GetPos(&ctrl_X, &ctrl_Y, &ctrlWidth, &ctrlHeight)

            valueEditHeight := grpBoxAvailableHeight / 2  ; Half the height of the groupbox minus some margin
            valuesEdit.Move(unset, unset, window_width - 50, valueEditHeight)  ; Set width to fill the window
            ; valuesEdit.Redraw()

            ; Now do the paths edit box and label
            startY := GetControlBottomEdge(valuesEdit)
            pathsEditLabel.Move(unset, startY + 10)  ; Move the label down 10 pixels from the bottom of the values edit box
            pathsEdit.Move(unset, startY + 30, window_width - 50, valueEditHeight)  ; Set width to fill the window

            ;------------------------------------------------------------------------------------------------------------

            ; Update control positions based on new window size
            for ctrl in thisGui {
                ctrl.GetPos(&ctrl_X, &ctrl_Y, &ctrlWidth, &ctrlHeight)

                ; For specific control objects
                if ctrl = typeDescription {
                    typeDropdown.GetPos(unset, &dropdownY, unset, unset)
                    ; Set height to the same as the dropdown, and fill the remaining width.
                    ctrl.Move(unset, dropdownY, window_width - ctrl_X - 20)  ; Set width to fill to the right edge of the window (with some margin)
                    ; ctrl.Redraw()
                    continue

                } else if ctrl = listView {
                    ctrl.Move(unset, unset, window_width - 30)  ; Set consistent width for edit boxes
                        ; ctrl.Redraw()

                ; Buttons
                } else if ctrl = saveBtn or ctrl = cancelBtn {
                    ctrl.Move(unset, window_height-45)
                    ; ctrl.Redraw()
                } else if ctrl = applyBtn {
                    ctrl.Move(window_width - ctrlWidth - 20, window_height-45)  ; Bottom align buttons with 40px margin from bottom
                    ; ctrl.Redraw()
                } else if ctrl = helpBtn {
                    ctrl.Move(window_width - ctrlWidth - 20)  ; Bottom align buttons with 40px margin from bottom

                ; For general control types
                } else if ctrl.HasProp("Type") {
                    ; Nothing here yet
                }
            }
        }
    }

    ShowConditionalFavoritesHelpWindow(*) {
        helpGui := Gui(unset, "Conditional Favorites - Help")
        helpGui.SetFont("s10", "Segoe UI")
        winWidth := 500
        textWidth := winWidth - 20  ; Subtract some margin from the right

        WidthForMargin(winWidth, textControl, margin) {
            textControl.GetPos(&ctrl_X, &ctrl_Y, &ctrlWidth, &ctrlHeight)
            remainingWidth := winWidth - ctrl_X
            return remainingWidth - margin
        }

        winWidthString := "w" winWidth
        txtWStr := "w" textWidth

        ; Title
        titleText := helpGui.AddText("xm y+10 " txtWStr " h20", "Conditional Favorites")
        titleText.SetFont("s12 bold")

        ; Overview
        helpGui.AddText("xm y+10 " txtWStr, "Conditional Favorites allow you to have the menu show certain paths only under specific conditions.")

        ; Supported Conditions Section
        conditionsHeader := helpGui.AddText("xm y+20 " txtWStr " h20", "Supported Conditions:")
        conditionsHeader.SetFont("s10 bold underline")
        
        conditionLabel1 := helpGui.AddText("xm+15 y+5 " txtWStr,  "• " ExplorerDialogPathSelector.ConditionType.DialogOwnerExe.FriendlyName)
        conditionLabel1.SetFont("s10 bold")
        conditionDesc1 := helpGui.AddText("xm+30 y+2 " txtWStr, "When the dialog window was launched by a certain program, based on the program's executable file name.")
        conditionDesc1.Move(unset, unset, WidthForMargin(winWidth, conditionDesc1, 20))
        
        conditionLabel2 := helpGui.AddText("xm+15 y+10 " txtWStr, "• " ExplorerDialogPathSelector.ConditionType.CurrentDialogPath.FriendlyName)
        conditionLabel2.SetFont("s10 bold")
        conditionDesc2 := helpGui.AddText("xm+30 y+2 " txtWStr " +Wrap", "When the current path of the dialog matches a specified path you set.")
        conditionDesc2.Move(unset, unset, WidthForMargin(winWidth, conditionDesc2, 20))

        conditionLabel3 := helpGui.AddText("xm+15 y+10 " txtWStr, "• " ExplorerDialogPathSelector.ConditionType.ParentWindowTitle.FriendlyName)
        conditionLabel3.SetFont("s10 bold")
        conditionDesc3 := helpGui.AddText("xm+30 y+2 " txtWStr " +Wrap", "When the title of the parent window (the app window that opened the dialog) contains a specified text string.")
        conditionDesc3.Move(unset, unset, WidthForMargin(winWidth, conditionDesc3, 20))

        conditionLabel3 := helpGui.AddText("xm+15 y+10 " txtWStr, "• " ExplorerDialogPathSelector.ConditionType.DialogFileTypeFilter.FriendlyName)
        conditionLabel3.SetFont("s10 bold")
        conditionDesc3 := helpGui.AddText("xm+30 y+2 " txtWStr " +Wrap", "When the dialog window's currently active file type filter matches your pattern. The filter can be the `"Save As Type`" dropdown for Save dialogs, or at the bottom right of Open dialogs. For example `"All Files (*.*)`". Applies to entire string, not just the file extensions.")
        conditionDesc3.Move(unset, unset, WidthForMargin(winWidth, conditionDesc3, 20))

        ; Condition Values Section
        valuesHeader := helpGui.AddText("xm y+20 " txtWStr " h20", "Condition Values:")
        valuesHeader.SetFont("s10 bold underline")
        helpGui.AddText("xm y+5 " txtWStr, "These let you define what will trigger the conditional favorite paths to show.")
        
        ; --- Notes
        notesText := helpGui.AddText("xm y+10 " txtWStr " h20", "Notes:")
        notesText.SetFont("italic")
        helpGui.AddText("xm+15 y+2 " txtWStr, "• Values are not case sensitive")
        helpGui.AddText("xm+15 y+2 " txtWStr, "• Asterisks can be used as wildcards")
        helpGui.AddText("xm+15 y+2 " txtWStr, "• Multiple values can be set per rule (one per line)")
        helpGui.AddText("xm+15 y+2 " txtWStr, "• If any values match, all associated paths will be shown")

        ; ---- Executable name match Examples
        execExamplesText := helpGui.AddText("xm y+10 " txtWStr, "Executable Name Match Examples: ")
        execExamplesText.SetFont("italic")
        execExample1 := helpGui.AddText("xm+15 y+5 " txtWStr, "• A conditional value of 'Photoshop.exe' will cause the associated paths to show when the dialog window is opened by Photoshop (Photoshop.exe).")
        execExample2 := helpGui.AddText("xm+15 y+5 " txtWStr, "• A conditional value 'Adobe*' will match when the window is opened by 'Adobe Reader.exe' or any other process that matches the wildcard pattern.")
        execExample1.Move(unset, unset, WidthForMargin(winWidth, execExample1, 20))
        execExample2.Move(unset, unset, WidthForMargin(winWidth, execExample2, 20))

        ; ---- Path match Examples
        pathExamplesText := helpGui.AddText("xm y+10 " txtWStr, "Path Match Examples: ")
        pathExamplesText.SetFont("italic")
        pathExample1 := helpGui.AddText("xm+15 y+5 " txtWStr, "• A conditional value of 'C:\Program Files\Adobe' will cause the associated paths to show when the dialog window's path is exactly`n'C:\Program Files\Adobe'")
        pathExample2 := helpGui.AddText("xm+15 y+5 " txtWStr, "• A conditional value of 'C:\Program Files\*' will match when the dialog window's path starts with 'C:\Program Files\'")
        pathExample1.Move(unset, unset, WidthForMargin(winWidth, pathExample1, 20))
        pathExample2.Move(unset, unset, WidthForMargin(winWidth, pathExample2, 20))

        ; ---- Parent Window Title match Examples
        parentExamplesText := helpGui.AddText("xm y+10 " txtWStr, "Parent Title Match Examples: ")
        parentExamplesText.SetFont("italic")
        parentExample1 := helpGui.AddText("xm+15 y+5 " txtWStr, "• A conditional value of '*YouTube*' would match if the dialog was opened by a browser tab titled `"Channel Dashboard - YouTube Studio`", for example. ")
        parentExample1.Move(unset, unset, WidthForMargin(winWidth, parentExample1, 20))

        ; ---- Filetype filter Example
        filterExampleText := helpGui.AddText("xm y+10 " txtWStr, "File Type Filter Match Examples: ")
        filterExampleText.SetFont("italic")
        filterExample1 := helpGui.AddText("xm+15 y+5 " txtWStr, "• If a Open/Save dialog's file type filter is `"Text Documents (*.txt)`", you could match with *.txt* or *Text* or even *Document* for example")
        filterExample1.Move(unset, unset, WidthForMargin(winWidth, filterExample1, 20))

        ; Paths Section
        pathsHeader := helpGui.AddText("xm y+20 " txtWStr " h20", "Paths:")
        pathsHeader.SetFont("s10 bold underline")
        helpGui.AddText("xm y+5 " txtWStr, "These are the paths that will be shown when the condition values match.")

        ; Close button
        closeButton := helpGui.AddButton("xm y+20 w80 Default", "Close")
        closeButton.OnEvent("Click", (*) => helpGui.Destroy())

        ; Get bottom position of close button to set the height of the window
        closeButton.GetPos(&ctrl_X, &ctrl_Y, &ctrlWidth, &ctrlHeight)
        windowHeight := ctrl_Y + ctrlHeight + 20

        ; Show with specific initial size
        helpGui.Show(winWidthString " h" windowHeight)

        ; Default to focus on the close button
        closeButton.Focus()
    }

    ShowFavoritePathsGui(*) {
        ; Create the main window
        pathGui := Gui("+Resize +MinSize400x300", "Favorite Paths Manager")
        
        ; Add instructions text
        pathGui.AddText(, "Enter folder paths to always show (one per line):")
        
        ; Add multi-line edit control with scrollbars
        editPaths := pathGui.AddEdit("vPaths r15 w400 Multi VScroll", "")
        
        ; Create a horizontal button layout using a GroupBox
        buttonGroup := pathGui.AddGroupBox("w400 h50", "")
        
        ; Add OK and Cancel buttons
        saveBtn := pathGui.AddButton("xp+20 yp+15 w80", "OK").OnEvent("Click", SavePaths)
        cancelBtn := pathGui.AddButton("x+10 yp w80", "Cancel").OnEvent("Click", (*) => pathGui.Destroy())
        
        ; Populate edit control with existing paths
        if this.g_pth_settings.HasProp("favoritePaths") && this.g_pth_settings.favoritePaths.Length > 0 {
            existingPaths := ""
            for path in this.g_pth_settings.favoritePaths {
                existingPaths .= path "`n"
            }
            editPaths.Value := RTrim(existingPaths, "`n")
        }
        
        ; Show the GUI
        pathGui.Show()
        
        ; Handle saving paths
        SavePaths(*) {
            ; Get the paths from the edit control
            rawPaths := editPaths.Value
            
            ; Split into array and remove empty lines
            pathArray := []
            for path in StrSplit(rawPaths, "`n", "`r") {
                if (Trim(path) != "") {
                    pathArray.Push(RTrim(Trim(path), "\"))
                }
            }

            ; Ensure none have invalid characters
            for path in pathArray {
                if (!this.ValidatePathCharacters(path)) {
                    MsgBox("Invalid characters found in path:`n" path "`n`nCannot contain these characters:`n< > : `" / | ? * `n`nPlease correct and try again.", "Error", "Icon!")
                    return
                }
            }
            
            ; Update the settings
            this.g_pth_settings.favoritePaths := pathArray
            
            ; Close the GUI
            pathGui.Destroy()
        }
    }

    ShowPathSelectorHelpWindow(*) {
        ; Added MinSize to prevent window from becoming too small
        helpGui := Gui("+Resize +MinSize400x300", g_pathSelector_programName " - Help & Tips")
        helpGui.SetFont("s10", "Segoe UI")
        helpGui.OnEvent("Size", GuiResize)

        ; hTT := ExplorerDialogPathSelector.CreateTooltipControl(helpGui.Hwnd) ; Currently no tooltips in help window

        ; Settings file info
        settingsFileDescription := helpGui.AddText("xm y+10 w300 h20", "Current config file path:") ; Creating this separately so we can set the font
        settingsFileDescription.SetFont("s10 bold")
        labelFileLocationText := ""
        labelFileLocationEdit := ""
        if this.g_pth_SettingsFile.usingSettingsFile {
            ; labelFileLocation := helpGui.AddText("xm y+0 w300 +Wrap", g_settingsFilePath)
            ; Show an edit text box so the user can copy the path and also so it word wraps properly even with no spaces
            labelFileLocationEdit := helpGui.AddEdit("xm y+5 w300 h30 +ReadOnly", this.g_pth_SettingsFile.filePath)
        } else {
            labelFileLocation := helpGui.AddText("xm y+0 w300 +Wrap", "N/A - Using default settings (no config file)")
        }

        ; AHK Key Names documentation link
        linkDescription := helpGui.AddText("xm y+20 w300 h20", "For information about key names in AutoHotkey, click here:") ; Creating this separately so we can set the font
        linkDescription.SetFont("s10 bold")
        linkText := '<a href="https://www.autohotkey.com/docs/v2/lib/Send.htm#keynames">https://www.autohotkey.com/docs/v2/lib/Send.htm</a>'
        keyNameLink := helpGui.AddLink("xm y+0 w300", linkText)

        ; Tips section
        tipsHeader := helpGui.AddText("xm y+20 w300 h20", "Tips:")
        tipsHeader.SetFont("s10 bold")

        ; Display info about UI Access depending on the mode the script is running in
        elevatedTipText := ""
        if A_IsCompiled {
            elevatedTipText := "• To make this program work with dialogs launched by elevated processes without having to run it as admin, place the executable in a trusted location such as `"C:\Program Files\...`""
            elevatedTipText .= "  (You do NOT need to run this exe itself as Admin for this to work.)"
        } else if !this.ThisScriptRunningStandalone() {
            elevatedTipText := "• To make this work with dialogs launched by elevated processes, enable UI Access via the parent script."
            elevatedTipText .= 'See this documentation page for more info:`n <a href="https://www.autohotkey.com/docs/v1/FAQ.htm#uac">https://www.autohotkey.com/docs/v1/FAQ.htm#uac</a>'
        } else {
            elevatedTipText := "• Enable `"UI Access`" setting to allow the script to work in dialogs from elevated windows without running this script itself as admin."
        }
        labelElevatedTip := helpGui.AddLink("xm y+5 w300", elevatedTipText)

        ; Display clickable link to source repo
        projectPageText := 'Check for updates on the project page: `n <a href="https://github.com/ThioJoe/ThioJoe-AHK-Scripts">https://github.com/ThioJoe/ThioJoe-AHK-Scripts</a>'
        labelProjectPage := helpGui.AddLink("xm y+0 w300", projectPageText)
        labelProjectPage.SetFont("s10 bold")


        ; ------------------------------------------------------------------------
        closeButton := helpGui.AddButton("xm y+10 w80 Default", "Close")
        closeButton.OnEvent("Click", (*) => helpGui.Destroy())

        ; Show with specific initial size
        helpGui.Show("w500 h360")

        GuiResize(thisGui, minMax, width, height) {
            if minMax = -1  ; The window has been minimized
                return

            ; Update control positions based on new window size
            for ctrl in thisGui {
                if ctrl.HasProp("Type") {
                    if ctrl.Type = "Text" or ctrl.Type = "Link" {
                        ctrl.Move(unset, unset, width - 25)  ; Add some margin to the right
                        ctrl.Redraw()
                    } else if ctrl.Type = "Button" {
                        if ctrl.Text = "Close" {
                            ctrl.Move(unset, height - 40)  ; Bottom align Close button with 40px margin from bottom
                        }
                    } else if ctrl.Type = "Edit" {
                        ctrl.Move(unset, unset, width - 25)  ; Add some margin to the right
                    }
                }
            }
        }

        if (labelFileLocationEdit != "") {
            labelFileLocationEdit.Value := labelFileLocationEdit.Value ; This is hacky but it forces the edit box to update and show the text properly, before it was strangely shifted until clicked on
        }
        ; Default to focus on the close button. Doesn't really matter which control, just so it doesn't select the text in the edit box
        closeButton.Focus()
    }

    ; 
    /**
     * Create a tooltip control window and return its handle
     * @param {Integer} guiHwnd 
     * @returns {Integer} Handle to the tooltip control
     */
    static CreateTooltipControl(guiHwnd) {
        ; Create tooltip window
        static ICC_TAB_CLASSES := 0x8
        static CW_USEDEFAULT := 0x80000000
        static TTS_ALWAYSTIP := 0x01
        static TTS_NOPREFIX := 0x02
        static WS_POPUP := 0x80000000

        ; Initialize common controls
        INITCOMMONCONTROLSEX := Buffer(8, 0)
        NumPut("UInt", 8, "UInt", ICC_TAB_CLASSES, INITCOMMONCONTROLSEX)
        DllCall("comctl32\InitCommonControlsEx", "Ptr", INITCOMMONCONTROLSEX)

        ; Create tooltip window
        hTT := DllCall("CreateWindowEx", "UInt", 0
            , "Str", "tooltips_class32"
            , "Ptr", 0
            , "UInt", TTS_ALWAYSTIP | TTS_NOPREFIX | WS_POPUP
            , "Int", CW_USEDEFAULT
            , "Int", CW_USEDEFAULT
            , "Int", CW_USEDEFAULT
            , "Int", CW_USEDEFAULT
            , "Ptr", guiHwnd
            , "Ptr", 0
            , "Ptr", 0
            , "Ptr", 0
            , "Ptr")

        ; Set maximum width to enable word wrapping and newlines in tooltips
        static TTM_SETMAXTIPWIDTH := 0x418
        PostMessage(
            TTM_SETMAXTIPWIDTH, ; Msg
            0,                  ; wParam (Must be zero for TTM_SETMAXTIPWIDTH)
            600,                ; lParam (Maximum tooltip window width, in pixels; or -1 to allow any width)
            hTT,                ; Hwnd to tooltip
        )

        return hTT
    }

    ; Add a tooltip to a control
    /**
     * 
     * @param {Integer} hTT Handle to the tooltip control window
     * @param {Integer} controlHwnd Handle to the control getting the tooltip
     * @param {String} text Text of the tooltip 
     */
    static AddTooltipToControl(hTT, controlHwnd, text) {
        ; TTM_ADDTOOLW - Unicode version only
        static TTM_ADDTOOL := 0x432
        ; Enum values used in TOOLINFO structure - See: https://learn.microsoft.com/en-us/windows/win32/api/commctrl/ns-commctrl-tttoolinfow
        static TTF_IDISHWND := 0x1
        static TTF_SUBCLASS := 0x10
        ; Static control style - See: https://learn.microsoft.com/en-us/windows/win32/controls/static-control-styles
        static SS_NOTIFY := 0x100
        static GWL_STYLE := -16 ; Used in SetWindowLongPtr: https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-setwindowlongptrw

        ; Check if this is a static text control and add SS_NOTIFY style if needed
        className := Buffer(256)
        if DllCall("GetClassName", "Ptr", controlHwnd, "Ptr", className, "Int", 256) {
            if (StrGet(className) = "Static") {
                ; Get current style
                currentStyle := DllCall("GetWindowLongPtr", "Ptr", controlHwnd, "Int", GWL_STYLE, "Ptr")
                ; Add SS_NOTIFY if it's not already there
                if !(currentStyle & SS_NOTIFY)
                    DllCall("SetWindowLongPtr", "Ptr", controlHwnd, "Int", GWL_STYLE, "Ptr", currentStyle | SS_NOTIFY)
            }
        }

        ; Create and populate TOOLINFO structure
        TOOLINFO := Buffer(A_PtrSize = 8 ? 72 : 44, 0)  ; Size differs between 32 and 64 bit

        ; Calculate size of TOOLINFO structure
        cbSize := A_PtrSize = 8 ? 72 : 44

        ; Populate TOOLINFO structure
        NumPut("UInt", cbSize, TOOLINFO, 0)   ; cbSize
        NumPut("UInt", TTF_IDISHWND | TTF_SUBCLASS, TOOLINFO, 4)   ; uFlags
        NumPut("Ptr",  controlHwnd,  TOOLINFO, A_PtrSize = 8 ? 16 : 12)  ; hwnd
        NumPut("Ptr",  controlHwnd,  TOOLINFO, A_PtrSize = 8 ? 24 : 16)  ; uId
        NumPut("Ptr",  StrPtr(text), TOOLINFO, A_PtrSize = 8 ? 48 : 36)  ; lpszText

        try {
            ; Add the tool - It seems only SendMessage works here. With PostMessage the tooltip text is gibberish.
            SendMessage(
                TTM_ADDTOOL,    ; Msg
                0,              ; wParam (Must be zero)
                TOOLINFO,       ; lParam (Must be TOOLINFO structure)
                hTT,            ; AHK "Control" Param - In this case handle to tooltip
                unset, unset, unset, unset, ; WinTitle, WinText, ExcludeTitle, ExcludeText
                50              ; Timeout, set it to something short
            )
        } catch {
            ; Try/Catch just in case it takes too long, so don't bother doing anything. It's not critical.
        }
        
    }



    SaveSettingsToFile() {
        ; ---------------------------- LOCAL FUNCTIONS ----------------------------
        GetConditionalFavoritesDelimitedString() {
            conditionalFavoritesString := ""
            ; Put double pipe between each entry, and single pipe between values. First value is the condition type
            i := 0
            for entry in this.g_pth_settings.conditionalFavorites {
                conditionalFavoritesString .= entry.ConditionType "||"
                j := 0
                for value in entry.ConditionValues {
                    conditionalFavoritesString .= value
                    ; Add a pipe between values, but not after the last one
                    if (j < entry.ConditionValues.Length - 1) {
                        conditionalFavoritesString .= "|"
                    }
                    j++
                }
                conditionalFavoritesString .= "||"
        
                j := 0
                for path in entry.Paths {
                    conditionalFavoritesString .= path
                    ; Add a pipe between paths, but not after the last one
                    if (j < entry.Paths.Length - 1) {
                        conditionalFavoritesString .= "|"
                    }
                    j++
                }
                ; Add a separator between entries, but not after the last one
                if (i < this.g_pth_settings.conditionalFavorites.Length - 1) {
                    conditionalFavoritesString .= "|||"
                }
                i++
            }
            return conditionalFavoritesString
        }

        GetFavoritesDelimitedString() {
            favoritePathsString := ""
            for path in this.g_pth_settings.favoritePaths {
                favoritePathsString .= path "|"
            }
            return favoritePathsString
        }
        ; -------------------------------------------------------------------------

        SaveToPath(settingsFileDir) {
            settingsFilePath := settingsFileDir "\" this.g_pth_SettingsFile.fileName

            fileAlreadyExisted := (FileExist(settingsFilePath) != "") ; If an empty string is returned from FileExist, the file was not found

            ; Create the necessary folders
            DirCreate(settingsFileDir)

            ; Save all settings to INI file
            IniWrite(this.g_pth_Settings.dialogMenuHotkey, settingsFilePath, "Settings", "dialogMenuHotkey")
            IniWrite(this.g_pth_Settings.dopusRTPath, settingsFilePath, "Settings", "dopusRTPath")
            ; Put quotes around the prefix and suffix values, otherwise spaces will be trimmed by the OS. The quotes will be removed when the values are read back in.
            IniWrite('"' this.g_pth_Settings.activeTabPrefix '"', settingsFilePath, "Settings", "activeTabPrefix")
            IniWrite('"' this.g_pth_Settings.activeTabSuffix '"', settingsFilePath, "Settings", "activeTabSuffix")
            IniWrite('"' this.g_pth_Settings.standardEntryPrefix '"', settingsFilePath, "Settings", "standardEntryPrefix")
            IniWrite(this.g_pth_Settings.enableExplorerDialogMenuDebug ? "1" : "0", settingsFilePath, "Settings", "enableExplorerDialogMenuDebug")
            IniWrite(this.g_pth_Settings.alwaysShowClipboardmenuItem ? "1" : "0", settingsFilePath, "Settings", "alwaysShowClipboardmenuItem")
            IniWrite(this.g_pth_Settings.groupPathsByWindow ? "1" : "0", settingsFilePath, "Settings", "groupPathsByWindow")
            IniWrite(this.g_pth_Settings.enableUIAccess ? "1" : "0", settingsFilePath, "Settings", "enableUIAccess")
            IniWrite(this.g_pth_Settings.maxMenuLength, settingsFilePath, "Settings", "maxMenuLength")
            IniWrite(this.g_pth_Settings.hideTrayIcon ? "1" : "0", settingsFilePath, "Settings", "hideTrayIcon")
            IniWrite(this.g_pth_Settings.useBoldTextActive ? "1" : "0", settingsFilePath, "Settings", "useBoldTextActive")
            IniWrite(GetFavoritesDelimitedString(), settingsFilePath, "Settings", "favoritePaths")
            IniWrite(GetConditionalFavoritesDelimitedString(), settingsFilePath, "Settings", "conditionalFavorites")

            this.g_pth_SettingsFile.usingSettingsFile := true

            if (!fileAlreadyExisted) {
                MsgBox("Settings saved to file:`n" this.g_pth_SettingsFile.fileName "`n`nIn Location:`n" settingsFilePath "`n`n Settings will be automatically loaded from file from now on.", "Settings File Created", "Iconi")
            }
        }

        ; Try saving to the current default settings path
        try {
            SaveToPath(this.g_pth_SettingsFile.directoryPath)
        } catch OSError as oErr {
            ; If it's error number 5, it's access denied, so try appdata path instead unless it's already the appdata path
            if (oErr.Number = 5 && this.g_pth_SettingsFile.filePath != this.g_pth_SettingsFile.appDataFilePath) {
                try {
                    ; Try to save to AppData path
                    SaveToPath(this.g_pth_SettingsFile.appDataDirectoryPath)
                    this.g_pth_SettingsFile.filePath := this.g_pth_SettingsFile.appDataFilePath ; If successful, update the global settings file path
                    this.g_pth_SettingsFile.directoryPath := this.g_pth_SettingsFile.appDataDirectoryPath
                } catch Error as innerErr {
                    MsgBox("Error saving settings to file:`n" innerErr.Message "`n`nTried to save in: `n" this.g_pth_SettingsFile.appDataFilePath, "Error Saving Settings", "Icon!")
                }
            } else if (oErr.Number = 5) {
                MsgBox("Error saving settings to file:`n" oErr.Message "`n`nTried to save in: `n" this.g_pth_SettingsFile.filePath, "Error Saving Settings", "Icon!")
            }
        } catch {
            MsgBox("Error saving settings to file:`n" A_LastError "`n`nTried to save in: `n" this.g_pth_SettingsFile.filePath, "Error Saving Settings", "Icon!")
        }

    }

    LoadSettingsFromSettingsFilePath() {
        ; If the settings file isn't in the current directory, but it is in AppData, use the AppData path
        if (!FileExist(this.g_pth_SettingsFile.filePath)) and FileExist(this.g_pth_SettingsFile.appDataFilePath) {
            this.g_pth_SettingsFile.filePath := this.g_pth_SettingsFile.appDataFilePath
            this.g_pth_SettingsFile.directoryPath := this.g_pth_SettingsFile.appDataDirectoryPath
        }

        local settingsFilePath := this.g_pth_SettingsFile.filePath

        ; ---------------------------- LOCAL FUNCTIONS ----------------------------
        ParseConditionalFavoritesString(conditionalFavoritesString) {
            conditionalFavorites := []
            if (conditionalFavoritesString = "") {
                return conditionalFavorites
            }
            entryIndex := 1
            entries := StrSplit(conditionalFavoritesString, "|||")
            for entry in entries {
                entryParts := StrSplit(entry, "||")
                if (entryParts.Length > 0) {
                    conditionTypeID := entryParts[1]
                    conditionValues := []
                    paths := []
                    for i, part in entryParts {
                        if (i = 1) {
                            continue ; Skip the first part which is the condition type
                        } else if (i = 2) { ; Condition values
                            conditionValues := StrSplit(part, "|")
                        } else { ; Paths
                            paths := StrSplit(part, "|") 
                        }
                    }
                    conditionalFavorites.Push({
                        ConditionType: conditionTypeID,
                        conditionTypeName: ExplorerDialogPathSelector.ConditionType.%conditionTypeID%.FriendlyName,
                        ConditionValues: conditionValues,
                        Paths: paths,
                        Index: entryIndex
                    })
                }
                entryIndex++
            }
        
            return conditionalFavorites
        }
        ; -------------------------------------------------------------------------
        
        if FileExist(settingsFilePath) {
            ; Load each setting from the INI file
            this.g_pth_Settings.dialogMenuHotkey := IniRead(settingsFilePath, "Settings", "dialogMenuHotkey", this.DefaultSettings.dialogMenuHotkey)
            this.g_pth_Settings.dopusRTPath := IniRead(settingsFilePath, "Settings", "dopusRTPath", this.DefaultSettings.dopusRTPath)
            this.g_pth_Settings.activeTabPrefix := IniRead(settingsFilePath, "Settings", "activeTabPrefix", this.DefaultSettings.activeTabPrefix)
            this.g_pth_Settings.activeTabSuffix := IniRead(settingsFilePath, "Settings", "activeTabSuffix", this.DefaultSettings.activeTabSuffix)
            this.g_pth_Settings.standardEntryPrefix := IniRead(settingsFilePath, "Settings", "standardEntryPrefix", this.DefaultSettings.standardEntryPrefix)
            this.g_pth_Settings.enableExplorerDialogMenuDebug := IniRead(settingsFilePath, "Settings", "enableExplorerDialogMenuDebug", this.DefaultSettings.enableExplorerDialogMenuDebug)
            this.g_pth_Settings.alwaysShowClipboardmenuItem := IniRead(settingsFilePath, "Settings", "alwaysShowClipboardmenuItem", this.DefaultSettings.alwaysShowClipboardmenuItem)
            this.g_pth_Settings.groupPathsByWindow := IniRead(settingsFilePath, "Settings", "groupPathsByWindow", this.DefaultSettings.groupPathsByWindow)
            this.g_pth_Settings.enableUIAccess := IniRead(settingsFilePath, "Settings", "enableUIAccess", this.DefaultSettings.enableUIAccess)
            this.g_pth_settings.maxMenuLength := IniRead(settingsFilePath, "Settings", "maxMenuLength", this.DefaultSettings.maxMenuLength)
            this.g_pth_Settings.hideTrayIcon := IniRead(settingsFilePath, "Settings", "hideTrayIcon", this.DefaultSettings.hideTrayIcon)
            this.g_pth_Settings.useBoldTextActive := IniRead(settingsFilePath, "Settings", "useBoldTextActive", this.DefaultSettings.useBoldTextActive)
            this.g_pth_settings.favoritePaths := StrSplit(IniRead(settingsFilePath, "Settings", "favoritePaths", ""), "|") ; Split the delimited string to an array
            this.g_pth_settings.conditionalFavorites := ParseConditionalFavoritesString(IniRead(settingsFilePath, "Settings", "conditionalFavorites", ""))

            ; Convert string boolean values to actual booleans
            this.g_pth_Settings.enableExplorerDialogMenuDebug := this.g_pth_Settings.enableExplorerDialogMenuDebug = "1"
            this.g_pth_Settings.alwaysShowClipboardmenuItem := this.g_pth_Settings.alwaysShowClipboardmenuItem = "1"
            this.g_pth_Settings.groupPathsByWindow := this.g_pth_Settings.groupPathsByWindow = "1"
            this.g_pth_Settings.enableUIAccess := this.g_pth_Settings.enableUIAccess = "1"
            this.g_pth_Settings.hideTrayIcon := this.g_pth_Settings.hideTrayIcon = "1"
            this.g_pth_Settings.useBoldTextActive := this.g_pth_Settings.useBoldTextActive = "1"

            ; Convert to int where necessary
            this.g_pth_settings.maxMenuLength := this.g_pth_settings.maxMenuLength + 0

            ; Remove empty entries from arrays
            this.g_pth_settings.favoritePaths := this.RemoveEmptyArrayEntries(this.g_pth_settings.favoritePaths)

            this.g_pth_SettingsFile.usingSettingsFile := true
        } else {
            ; If no settings file exists, use defaults
            for k, v in this.DefaultSettings.OwnProps() {
                this.g_pth_Settings.%k% := this.DefaultSettings.%k%
            }
        }
    }

    static ShowPathSelectorAboutWindow(*) {
        MsgBox(g_pathSelector_programName "`nVersion: " g_pathSelector_version "`n`nAuthor: ThioJoe`n`nProject Repository: https://github.com/ThioJoe/AHK-Scripts", "About", "Iconi")
    }

    ; ---------------------------- SYSTEM TRAY MENU CUSTOMIZATION ----------------------------
    SetupSystemTray(traySettings, hideTrayIconOption) {
        ; Ensure the parameter type is correct, otherwise use a new instance of the class to use defaults
        if (traySettings.base.__Class != ExplorerDialogPathSelector.SystemTraySettings.prototype.__Class) {
            MsgBox("Invalid object type was used for Explorer Dialog Path Selector system tray settings. Must be type: ExplorerDialogPathSelector.SystemTraySettings . `n Using default settings instead.")
            traySettings := ExplorerDialogPathSelector.SystemTraySettings()
        }

        if (hideTrayIconOption) {
            A_IconHidden := true
        }

        if (!traySettings.showSettingsTrayMenuItem) {
            return
        }

        if (traySettings.customIconFilePath != "") {
            ; Set the icon for the tray menu
            try {
                TraySetIcon(traySettings.customIconFilePath)
            } catch as err {
                MsgBox("Error setting custom icon: " err.Message)
            }
        }

        ; Set a variable for this because it might be updated
        local positionIndex := traySettings.positionIndex

        if (!traySettings.forcePositionIndex) {
            if (this.ThisScriptRunningStandalone() or A_IsCompiled) {
                ; If it's running standalone or compiled, use the default position
            } else {
                ; If it's running as an included script, put it next so it goes after the parent script's menu items
                positionIndex := positionIndex + 1
            }
        }

        ; Uses the & symbol to indicate the 1-indexed position of the menu item
        if (traySettings.addSeparatorBefore) {
            A_TrayMenu.Insert(positionIndex . "&", "")  ; Separator
            positionIndex := positionIndex + 1
        }

        A_TrayMenu.Insert(positionIndex . "&", traySettings.settingsMenuItemName, this.ShowPathSelectorSettingsGUI_Bound)

        if (traySettings.addSeparatorAfter) {
            A_TrayMenu.Insert((positionIndex + 1) . "&", "")  ; Separator - Comment/Uncomment if you want to add a separator or not
        }

        ; If the script is compiled or running standalone, make the settings the default menu item
        if (this.ThisScriptRunningStandalone() or A_IsCompiled or traySettings.alwaysDefaultItem) {
            A_TrayMenu.Default := traySettings.settingsMenuItemName
        }
    }

} ; ----------------------- END OF ExplorerDialogPathSelector -----------------------