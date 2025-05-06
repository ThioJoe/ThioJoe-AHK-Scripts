#Requires AutoHotkey v2.0

; ------- Snipping Tool Controller Script -------
; Purpose: This script allows you to launch Snipping Tool directly to specific modes.
; Author:  ThioJoe
; Repo:    https://github.com/ThioJoe/ThioJoe-AHK-Scripts
; Version: 1.0.0
; -----------------------------------------------
;
; REQUIRED: Set the path to the required UIA.ahk class file. Here it is up one directory then in the Lib folder. If it's in the same folder it would be:  #Include "UIA.ahk"
;       You can acquire UIA.ahk here:  https://github.com/Descolada/UIA-v2/blob/main/Lib/UIA.ahk
#Include "..\Lib\UIA.ahk"

; ========================================== HOW TO USE THIS SCRIPT ==========================================
; 1. Include this script in your main script using #Include, or add a hotkey directly in this script and run it standalone.
; 2. Call the ActivateSnippingToolAction() function with the desired "Action" item as an argument
;       Optional: You can also use the SetSnippingToolMode() function directly if you want to set a specific mode without launching the tool, like if you know it's already open.

; EXAMPLE USAGE:  "Ctrl + PrintScreen"  to go directly to text extractor mode. You can even uncomment this exact line to use it if you want. Or call the same function from your main script.
;   ^PrintScreen:: ActivateSnippingToolAction(Actions.TextExtractor)

; Optional Parameters:
;    Parameter Position 2:  autoClickToast: If set to true, it will automatically click the toast notification that appears after taking a screenshot (within 15 seconds).
; ============================================================================================================

; ------ Available options for Actions (Do Not Change - For Reference Only)  ------
; Pass them as a parameter to ActivateSnippingToolAction() in the form like "Actions.Rectangle" or "Actions.TextExtractor".
class Actions {
    static Rectangle := 1
    static Window    := 2
    static FullScreen := 3
    static Freeform  := 4
    static TextExtractor := 5
    static Video := 6
    static Close := 7
}

; ====================================================================================================
; ====================================================================================================
; ====================================================================================================

; Create a new template class that will store names, UIA paths, AutomationIDs, etc.
class SnipToolbarUIA {
    __New(name, path, automationId, type, parent := 0) {
        this.Name := name
        this.Path := path
        this.AutomationId := automationId
        this.Type := type
        this.ParentElement := parent
    }
}

; Parent Elements
snipToolString := "Snipping Tool Overlay ahk_exe SnippingTool.exe"
snippingOverlayElement   := SnipToolbarUIA("Snipping Tool Overlay"      , "YR/80", ""                            , "Window")
modeDropdownElement      := SnipToolbarUIA("Snipping Mode Dropdown Menu", "YR/3" , "SnippingModeComboBox"        , "ComboBox")
popUpHost                := SnipToolbarUIA("PopupHost"                   , "YR/3" , ""                           , "Pane", modeDropdownElement)

; Top level Button and Switches
captureModeToggleElement := SnipToolbarUIA("[Capture Mode Toggle]"       ,  "YR/0", "CaptureModeToggleSwitch"    , "Button") ; Name varies based on mode
textExtractorElement     := SnipToolbarUIA("Text extractor"              , "YR/80", "AuxiliaryModeToolbarButton" , "Button")
mainCloseButtonElement   := SnipToolbarUIA("Close"                       , "YR/0/", "CloseButton"                , "Button")

; List items
rectangleModeElement     := SnipToolbarUIA("Rectangle"                   , "X37"  , ""                           , "ListItem", popUpHost)
windowModeElement        := SnipToolbarUIA("Window"                      , "X37q" , ""                           , "ListItem", popUpHost)
fullScreenModeElement    := SnipToolbarUIA("Full screen"                 , "X37r" , ""                           , "ListItem", popUpHost)
freeformModeElement      := SnipToolbarUIA("Freeform"                    , "X37/" , ""                           , "ListItem", popUpHost)


CheckIfVideoMode() {
    try {
        local SnipToolOverlay := UIA.ElementFromHandle(snipToolString)
        local Condition := UIA.CreateCondition("AutomationId", modeDropdownElement.AutomationId)
        ; local dropdown := snipToolOverlay.WaitElement(Condition, UIA.TreeScope.Descendants, 1000) ; Wait for the element to be present
        Sleep(50)
        local dropdown := SnipToolOverlay.FindFirst(Condition, UIA.TreeScope.Subtree)

        ; If the dropdown is not enabled, it means we're in video mode
        if (IsObject(dropdown) && !dropdown.GetPropertyValue("IsEnabled")) {
            return true
        } else {
            return false
        }
    } catch as e {
        OutputDebug("`nError checking video mode: " e.Message "`nAt line: " e.Line)
        return false
    }
}


ActivateSnippingToolAction(elementEnum, autoClickToast := false) {
    try {
        Launch_UWP_With_Args("Microsoft.ScreenSketch_8wekyb3d8bbwe!App", "new-snip")
    } catch Error as e {
        ; Fallback requires the print screen key to be set as the hotkey for the Snipping Tool
        Send("{PrintScreen}")
        OutputDebug("`nFailed to launch snipping tool via Activation Context, falling back to Print Screen: " . e.Message)
    }
    
    WinWaitActive("Snipping Tool Overlay", unset, 2) ; Add a small timeout
    if !WinActive("Snipping Tool Overlay") {
        OutputDebug("`nSnipping Tool Overlay did not become active.")
        return
    }

    ; Once the snipping tool is active, we can proceed to set or invoke the desired action
    SetSnippingToolMode(elementEnum)

    if (autoClickToast && elementEnum != Actions.Close && elementEnum != Actions.TextExtractor && elementEnum != Actions.Video) {
        ; Check for the toast notification and click it
        CallFunctionWithTimeout(CheckAndClickToast, 350, 15) ; Check every 350ms for 15 seconds
    }
}


SetSnippingToolMode(elementEnum) {
    
    if (elementEnum == Actions.Rectangle || elementEnum == Actions.Window || elementEnum == Actions.FullScreen || elementEnum == Actions.Freeform)
    {
        if (CheckIfVideoMode())
            InvokeElement(captureModeToggleElement, snipToolString)

        ; Now we can invoke the desired mode
        if (elementEnum == Actions.Rectangle)
            InvokeElement(rectangleModeElement, snipToolString)
        else if (elementEnum == Actions.Window)
            InvokeElement(windowModeElement, snipToolString)
        else if (elementEnum == Actions.FullScreen)
            InvokeElement(fullScreenModeElement, snipToolString)
        else if (elementEnum == Actions.Freeform)
            InvokeElement(freeformModeElement, snipToolString)
    } else if (elementEnum == Actions.Video) {

        ; Check if we're in video mode first, and switch if necessary
        if (!CheckIfVideoMode())
            InvokeElement(captureModeToggleElement, snipToolString)

    } else if (elementEnum == Actions.TextExtractor) {
        InvokeElement(textExtractorElement, snipToolString)
    } else if (elementEnum == Actions.Close) {
        InvokeElement(mainCloseButtonElement, snipToolString)
    } else {
        OutputDebug("`nUnknown action selected.")
    }
}


InvokeElement(element, initialElementString) {
    local button := 0 ; Reset variable

    try {
        ; Check if the element is valid
        if !IsObject(element) {
            OutputDebug("`nInvalid element passed to InvokeElement.")
            return false
        }

        ; Get the main window element
        local MainElement := UIA.ElementFromHandle(initialElementString)
        if !IsObject(MainElement) {
            OutputDebug("`nFailed to get element.")
            return false
        }

        ; If it has a parent element that needs to be opened first, do that
        if (element.ParentElement !== 0) {
            local parentSuccess := InvokeElement(element.ParentElement, initialElementString) ; Recursively invoke the parent element
            if (parentSuccess) {
                OutputDebug("`nParent Element `"" element.ParentElement.Name "`" opened.")
            } else {
                OutputDebug("`nFailed to open parent element `"" element.ParentElement.Name "`". Will keep going.")
                ; Keep going just in case it somehow still works
            }
        }

        local condition := 0
        if (element.AutomationId !== "") {
            OutputDebug("`nAttempting to find element using AutomationId: " element.AutomationId)
            Condition := UIA.CreateCondition("AutomationId", element.AutomationId)
        } else {
            Sleep(10) ; These need some time to load apparently
            OutputDebug("`nAttempting to find element using Name: " element.Name)
            Condition := UIA.CreateCondition("Name", element.Name)
        }

        try { 
            button := MainElement.FindFirst(Condition, UIA.TreeScope.Subtree) 
        } catch as e {
            Sleep(50) ; Give it a little time to load and try again
            try {
                button := MainElement.FindFirst(Condition, UIA.TreeScope.Subtree)
            }
        }

        if IsObject(button) {
            OutputDebug("`n" element.Name " found through AutomationId.")
        } else {
            ; Fall back to find the button using the path
            try {
                button := MainElement.ElementFromPath(element.Path)
            }
            
            if IsObject(button) {
                OutputDebug("`n" element.Name " found through path.")
            }
        }

        if IsObject(button) {
            try {
                button.Click() ; This will automatically try to use the proper method (Invoke, Toggle, etc.). It does not move the mouse to work.
            }
            OutputDebug("`n`"" element.Name "`" invoked successfully.")
            return true
        } else {
            OutputDebug("`n`"" element.Name "`" not found.")
            return false
        }

    } catch as e {
        OutputDebug("`n" element.Name " -- Error invoking element: " e.Message "`nAt line: " e.Line)
        return false
    }
}

/**
 * Launches a UWP application using its AUMID and an activation context string
 * via the IApplicationActivationManager COM interface using ComCall.
 */
Launch_UWP_With_Args(appUserModelId, activationContext, options := 0) {
    local CLSID_ApplicationActivationManager := "{45BA127D-10A8-46EA-8AB7-56EA9078943C}"
    local IID_IApplicationActivationManager  := "{2E941141-7F97-4756-BA1D-9DECDE894A3D}"
    local S_OK := 0

    local processIdBuffer := Buffer(4, 0)
    local manager := unset

    try {
        ; Create the COM object - this returns a ComValue wrapper because we specified a non-IDispatch IID
        manager := ComObject(CLSID_ApplicationActivationManager, IID_IApplicationActivationManager)
        if !IsObject(manager) {
            throw(Error("Failed to create IApplicationActivationManager COM object."))
        }

        ; --- Use ComCall to invoke the method via the interface pointer ---
        ; HRESULT ActivateApplication(LPCWSTR, LPCWSTR, ACTIVATEOPTIONS, DWORD*) is VTable index 3
        hResult := ComCall(3, manager, "WStr", appUserModelId, "WStr", activationContext, "Int", options, "Ptr", processIdBuffer.Ptr)

        ; Check the HRESULT returned by the COM method itself
        if (hResult != S_OK) {
            errMsg := "IApplicationActivationManager.ActivateApplication (via ComCall) failed with HRESULT: 0x" . Format("{:X}", hResult)
            if (hResult == 0x80070002) errMsg .= " (Error: File Not Found - Check AUMID?)"
            if (hResult == 0x800704C7) errMsg .= " (Error: Cancelled?)"
            if (hResult == 0x80270254) errMsg .= " (Error: Activation Failed - Privileges/Disabled?)"
            throw(Error(errMsg))
        }

        local launchedPID := NumGet(processIdBuffer, 0, "UInt")
        return launchedPID

    } catch Error as e {
        throw(Error("Error during UWP activation: " . e.Message))
    }
    ; 'manager' (ComValue) is automatically released when it goes out of scope.
}


; Loop that checks for the presence of a toast notification ahk_exe ShellExperienceHost.exe, title of New notification
toastString := "New notification ahk_exe ShellExperienceHost.exe"
isToastClickTimerRunning := false

CheckAndClickToast() {
    local existResult := WinExist("ahk_exe ShellExperienceHost.exe")
    local MainElement := 0
    local button := 0

    if (existResult)
    {
        try {
            MainElement := UIA.ElementFromHandle(toastString)
            if (isObject(MainElement))
            {
                Condition := UIA.CreateCondition("AutomationId", "VerbButton")
                try {
                    button := MainElement.FindFirst(Condition, UIA.TreeScope.Subtree)
                }

                if (isObject(button))
                {
                    Sleep(250) ; Wait for the button to be ready
                    button.Click()
                    CancelToastCheckTimer() ; Cancel the timer
                    WaitAndActivateSnipWindow() ; Wait for the snipping tool window to exist and activate it
                    OutputDebug("`nToast found and clicked.")
                    return true
                } 
            }
        } catch as e {
            OutputDebug("`nError checking for toast: " e.Message "`nAt line: " e.Line)
        }
    }
    ; OutputDebug("`nToast not found or button not clickable.")
    ; Return false unless we find the button
    return false
}

; Wait for the snipping tool window to exist and activate it and bring to front
WaitAndActivateSnipWindow() {
    local snipWindow := WinWait("Snipping Tool", unset, 2) ; Add a small timeout
    if (snipWindow) {
        WinActivate(snipWindow)
        WinWaitActive(snipWindow)
        OutputDebug("`nSnipping Tool Overlay activated.")
    } else {
        OutputDebug("`nSnipping Tool Overlay not found.")
    }
}

CancelToastCheckTimer() {
    ; Cancels the timer
    SetTimer(CheckAndClickToast, 0)
    global isToastClickTimerRunning := false
    OutputDebug("`nTimer cancelled.")
}


CallFunctionWithTimeout(FuncToCall, IntervalMs, TimeoutSeconds) {
    local startTime := A_TickCount
    local endTime := startTime + (TimeoutSeconds * 1000)

    if (isToastClickTimerRunning) {
        ; OutputDebug("`nTimer already running. Exiting.")
        return
    } else {
        global isToastClickTimerRunning := true
    }

    ; Use SetTimer to call the function at intervals but stop after the timeout or if it returns true
    SetTimer(FuncToCall, IntervalMs)

    ; Cancel the timer after the timeout. Using negative makes it run once after the timeout
    SetTimer(CancelToastCheckTimer, (TimeoutSeconds * -1000)) 

}