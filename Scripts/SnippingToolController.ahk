#Requires AutoHotkey v2.0

; ------- Snipping Tool Controller Script -------
; Purpose: This script allows you to launch Snipping Tool directly to specific modes.
; Author:  ThioJoe
; Repo:    https://github.com/ThioJoe/ThioJoe-AHK-Scripts
; Version: 1.3.0
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
;   ^PrintScreen:: ActivateSnippingToolAction(SnippingToolActions.TextExtractor, true)

; Optional Parameters:
;    Parameter Position 2:  autoClickToast: If set to true, it will automatically click the toast notification that appears after taking a screenshot (within 15 seconds).
; ============================================================================================================

; ------ Available options of SnippingToolActions (Do Not Change - For Reference Only)  ------
; Pass them as a parameter to ActivateSnippingToolAction() in the form like "SnippingToolActions.Rectangle" or "SnippingToolActions.TextExtractor".
class SnippingToolActions {
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

; Simple Helper "proxy" function to call the classe's function without having to create an instance of the class or specifying it with a long name
/**
 * Launches the Snipping Tool and activates a specific snipping action or mode.
 * It will first attempt to launch the Snipping Tool UWP app with arguments. If that fails, it falls back to sending the PrintScreen key.
 * Once the Snipping Tool overlay is active, it will set the desired snipping mode.
 *
 * @param {SnippingToolActions} elementEnum A value from the `SnippingToolActions` class (e.g., `SnippingToolActions.Rectangle`, `SnippingToolActions.TextExtractor`).
 * This determines which mode or action the Snipping Tool will be set to.
 * @param {Boolean} [autoClickToast=false] Optional: If set to true, the script will attempt to automatically click the "Markup and share" toast notification
 * that appears after a snip is taken (excluding TextExtractor, Video, and Close actions).
 * The script will monitor for this toast for up to 15 seconds. Default is `false`.
 * @returns This function does not explicitly return a value. Its effects are the launching and controlling of the Snipping Tool.
 */
ActivateSnippingToolAction(elementEnum, autoClickToast := false) {
    if !IsSet(elementEnum) || !IsNumber(elementEnum) {
            OutputDebug("`nInvalid elementEnum passed to ActivateAction.")
            throw(Error("Invalid elementEnum passed to ActivateAction."))
        }
    SnippingToolController.ActivateAction(elementEnum, autoClickToast)
}

; ----------------------------------- Main Class ---------------------------------
class SnippingToolController {

    ; Class level variables
    static toastString := "New notification ahk_exe ShellExperienceHost.exe"
    static isToastClickTimerRunning := false
    static haveClickedToast := false

    static ActivateAction(elementEnum, autoClickToast := false) {
        if (elementEnum == SnippingToolActions.Rectangle 
            || elementEnum == SnippingToolActions.Window 
            || elementEnum == SnippingToolActions.Freeform 
            || elementEnum == SnippingToolActions.Video 
            || elementEnum == SnippingToolActions.FullScreen)
        {
            this.ActivateActionDirectly(elementEnum, autoClickToast)
        } else {
            this.ActivateActionWithUIA(elementEnum, autoClickToast)
        }
    }

    ; For certain modes we can launch directly using ms-screenclip protocol commands
    static ActivateActionDirectly(elementEnum, autoClickToast) {
        local shell := ComObject("Shell.Application")

        if (elementEnum == SnippingToolActions.Rectangle) {
            shell.ShellExecute("ms-screenclip:capture?clippingMode=Rectangle")
        } else if (elementEnum == SnippingToolActions.Window) {
            shell.ShellExecute("ms-screenclip:capture?clippingMode=Window")
        } else if (elementEnum == SnippingToolActions.Freeform) {
            shell.ShellExecute("ms-screenclip:capture?clippingMode=Freeform")
        } else if (elementEnum == SnippingToolActions.Video) {
            shell.ShellExecute("ms-screenclip:capture?type=recording")
        } else if (elementEnum == SnippingToolActions.FullScreen) {
            shell.ShellExecute("ms-screenclip:capture?type=snapshot")
        }

        this.EnableAutoclickIfApplicable(elementEnum, autoClickToast)
    }

    static EnableAutoclickIfApplicable(elementEnum, autoClickToast) {
        if (autoClickToast && elementEnum != SnippingToolActions.Close && elementEnum != SnippingToolActions.TextExtractor && elementEnum != SnippingToolActions.Video) {
            ; Check for the toast notification and click it
            this.haveClickedToast := false ; Reset the flag
            this.CallFunctionWithTimeout(350, 15) ; Check every 350ms for 15 seconds
        }
    }

    static ActivateActionWithUIA(elementEnum, autoClickToast) {
        try {
            this.Launch_UWP_With_Args("Microsoft.ScreenSketch_8wekyb3d8bbwe!App", "new-snip")
        } catch Error as e {
            ; Fallback requires the print screen key to be set as the hotkey for the Snipping Tool
            Send("{PrintScreen}")
            OutputDebug("`nFailed to launch snipping tool via Activation Context, falling back to Print Screen: " . e.Message)
        }
        
        WinWaitActive(this.snipToolString, unset, 2) ; Add a small timeout
        if !WinActive("Snipping Tool Overlay") {
            OutputDebug("`nSnipping Tool Overlay did not become active.")
            return
        }

        ; Once the snipping tool is active, we can proceed to set or invoke the desired action
        this.SetSnippingToolMode(elementEnum)

        this.EnableAutoclickIfApplicable(elementEnum, autoClickToast)
    }

    ; Create a new template class that will store names, UIA paths, AutomationIDs, etc.
    class SnipToolbarUIA {
        __New(name, path, automationId, type, parent := 0, parentRootString := "") {
            this.Name := name
            this.Path := path
            this.AutomationId := automationId
            this.Type := type
            this.ParentElement := parent
            this.ParentRootString := parentRootString
        }
    }

    ; Parent Elements
    static snipToolString := "Snipping Tool Overlay ahk_exe SnippingTool.exe"
    static dropDownString := "PopupHost ahk_exe SnippingTool.exe" ; A new window is created for the dropdown

    ; snippingOverlayElement   := SnipToolbarUIA("Snipping Tool Overlay"      , "YR/80", ""                            , "Window")
    ; popUpHost                := SnipToolbarUIA("PopupHost"                   , "YR/3" , ""                           , "Pane")

        ; Top level Button and Switches
    static captureModeToggleElement := this.SnipToolbarUIA("[Capture Mode Toggle]"       ,  "YR/0", "CaptureModeToggleSwitch"    , "Button") ; Name varies based on mode
    static textExtractorElement     := this.SnipToolbarUIA("Text extractor"              , "YR/80", "AuxiliaryModeToolbarButton" , "Button")
    static mainCloseButtonElement   := this.SnipToolbarUIA("Close"                       , "YR/0/", "CloseButton"                , "Button")

    ; List items
    static modeDropdownElement      := this.SnipToolbarUIA("Snipping Mode Dropdown Menu", "YR/3" , "SnippingModeComboBox"        , "ComboBox", 0                       , this.dropDownString)
    static rectangleModeElement     := this.SnipToolbarUIA("Rectangle"                   , "X37"  , ""                           , "ListItem", this.modeDropdownElement, unset)
    static windowModeElement        := this.SnipToolbarUIA("Window"                      , "X37q" , ""                           , "ListItem", this.modeDropdownElement, unset)
    static fullScreenModeElement    := this.SnipToolbarUIA("Full screen"                 , "X37r" , ""                           , "ListItem", this.modeDropdownElement, unset)
    static freeformModeElement      := this.SnipToolbarUIA("Freeform"                    , "X37/" , ""                           , "ListItem", this.modeDropdownElement, unset)

    static SetSnippingToolMode(elementEnum) {
    
        local currentMode := this.CheckCurrentMode()

        ; If we want to active a non-video mode, we need to check if we're in video mode first
        if (elementEnum == SnippingToolActions.Rectangle || elementEnum == SnippingToolActions.Window || elementEnum == SnippingToolActions.FullScreen || elementEnum == SnippingToolActions.Freeform)
        {
            ; If we are in video mode, toggle to the capture mode to image mode
            if (currentMode == SnippingToolActions.Video)
                this.InvokeElement(this.captureModeToggleElement, this.snipToolString)

            ; Now we can invoke the desired mode. If we're already in the desired mode, we don't need to do anything
            if (elementEnum == SnippingToolActions.Rectangle && currentMode != SnippingToolActions.Rectangle)
                this.InvokeElement(this.rectangleModeElement, this.snipToolString)
            else if (elementEnum == SnippingToolActions.Window && currentMode != SnippingToolActions.Window)
                this.InvokeElement(this.windowModeElement, this.snipToolString)
            else if (elementEnum == SnippingToolActions.FullScreen && currentMode != SnippingToolActions.FullScreen)
                this.InvokeElement(this.fullScreenModeElement, this.snipToolString)
            else if (elementEnum == SnippingToolActions.Freeform && currentMode != SnippingToolActions.Freeform)
                this.InvokeElement(this.freeformModeElement, this.snipToolString)

        } else if (elementEnum == SnippingToolActions.Video) {
            ; Check if we're in video mode first, and switch if necessary
            if (currentMode != SnippingToolActions.Video)
                this.InvokeElement(this.captureModeToggleElement, this.snipToolString)
            else
                OutputDebug("`nAlready in video mode. No need to switch.")
        } else if (elementEnum == SnippingToolActions.TextExtractor) {
            this.InvokeElement(this.textExtractorElement, this.snipToolString)
        } else if (elementEnum == SnippingToolActions.Close) {
            this.InvokeElement(this.mainCloseButtonElement, this.snipToolString)
        } else {
            OutputDebug("`nUnknown action selected.")
        }

    }
    
    static CheckCurrentMode() {
        try {
            local SnipToolOverlay := UIA.ElementFromHandle(this.snipToolString)
            local Condition := UIA.CreateCondition("AutomationId", this.modeDropdownElement.AutomationId)
            local dropdown := snipToolOverlay.WaitElement(Condition, 1000, UIA.TreeScope.Descendants) ; Wait for the element to be present

            ; If the dropdown is not enabled, it means we're in video mode
            if (IsObject(dropdown)) {
                ; OutputDebug("`nDropdown found: " dropdown.Name)
                if (!dropdown.GetPropertyValue("IsEnabled")) {
                    OutputDebug("`nDropdown is not enabled. We're in video mode.")
                    return SnippingToolActions.Video
                } else {
                    ; Try to get the selected item's name (current mode), whhich should be the first and only child of the dropdown
                    local selectedItem := 0
                    try {
                        selectedItem := dropdown.WaitElement(UIA.CreateCondition("AutomationId", "ContentPresenter"), 250, UIA.TreeScope.Descendants) ; Working
                    }
                    
                    ; If not found, try to get it by Type
                    if (IsObject(selectedItem)) {
                        OutputDebug("`nFound selected item via AutomationId: " selectedItem.Name)
                    } else {
                        OutputDebug("`n!  Couldn't find selected dropdown item via AutomationId while checking mode.")
                        try {
                            selectedItem := dropdown.WaitElement(UIA.CreateCondition("Type", UIA.Type.ListItem), 250, UIA.TreeScope.Descendants)
                        }

                        if (IsObject(selectedItem)) {
                            OutputDebug("`nFound selected dropdown item via Type instead.")
                        } else {
                            OutputDebug("`n!  Couldn't find selected dropdown item via Name while checking mode.")
                            return 0 ; Not known mode
                        }
                    }
                    
                    OutputDebug("`nCurrently selected mode: " selectedItem.Name)
                    if (selectedItem.Name == this.rectangleModeElement.Name) {
                        return SnippingToolActions.Rectangle ; It's video mode
                    } else if (selectedItem.Name == this.windowModeElement.Name) {
                        return SnippingToolActions.Window ; It's window mode
                    } else if (selectedItem.Name == this.fullScreenModeElement.Name) {
                        return SnippingToolActions.FullScreen ; It's full screen mode
                    } else if (selectedItem.Name == this.freeformModeElement.Name) {
                        return SnippingToolActions.Freeform ; It's freeform mode
                    } else {
                        OutputDebug("`nUnknown mode: " selectedItem.Name)
                        return 0 ; Not known mode
                    }
                }
            } else {
                OutputDebug("`n!  Dropdown not found while checking mode.")
            }
        } catch as e {
            OutputDebug("`nError checking video mode: " e.Message "`nAt line: " e.Line)
            return 0
        }

        return 0
    }

    static InvokeElement(element, initialElementString) {
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
                local stringToUse := ""
                if (element.ParentRootString !== "") {
                    stringToUse := element.ParentRootString
                    WinActivate(stringToUse) ; If there's a new root element to use, wait for it to be active
                    WinWaitActive(stringToUse, unset, 2) 
                } else {
                    stringToUse := initialElementString
                }

                local parentSuccess := this.InvokeElement(element.ParentElement, stringToUse) ; Recursively invoke the parent element
                if (parentSuccess) {
                    OutputDebug("`nParent Element `"" element.ParentElement.Name "`" opened.")
                } else {
                    OutputDebug("`nFailed to open parent element `"" element.ParentElement.Name "`". Will keep going.")
                    ; Keep going just in case it somehow still works
                }
            }

            ; Prepare to find the element by defining the condition
            local condition := 0
            if (element.AutomationId !== "") {
                OutputDebug("`nAttempting to find element using AutomationId: " element.AutomationId)
                Condition := UIA.CreateCondition("AutomationId", element.AutomationId)
            } else {
                ; Sleep(10) ; These need some time to load apparently
                OutputDebug("`nAttempting to find element using Name: " element.Name)
                Condition := UIA.CreateCondition("Name", element.Name)
            }

            ; A key line to find the element. This will search the entire subtree of the main element.
            try { 
                button := MainElement.WaitElement(Condition, 1000, UIA.TreeScope.Subtree) ; Wait for the element to be present
            } 
            catch as e {
                Sleep(50) ; Give it a little time to load and try again. This shouldn't happen with WaitElement but just in case.
                try {
                    button := MainElement.FindFirst(Condition, UIA.TreeScope.Subtree)
                }
            }

            ; Check if we found the button / element. Fallback to using the path if not.
            if IsObject(button) {
                OutputDebug("`n" element.Name " found.")
            } else {
                ; Fall back to find the button using the path
                try {
                    button := MainElement.ElementFromPath(element.Path)
                }
                
                if IsObject(button) {
                    OutputDebug("`n" element.Name " found through path.")
                }
            }

            ; Final check to see if we found the button, and invoke it
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
    static Launch_UWP_With_Args(appUserModelId, activationContext, options := 0) {
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

    ; Wait for the snipping tool window to exist and activate it and bring to front
    static WaitAndActivateSnipWindow() {
        local snipWindow := WinWait("Snipping Tool ahk_exe SnippingTool.exe", unset, 3) ; Waits for the window to exist, with a timeout
        if (snipWindow) {
            WinActivate(snipWindow)
            WinWaitActive(snipWindow, unset, 0.5) ; Wait for the window to be active. Should be instant but just in case
            if (WinActive("Snipping Tool ahk_exe SnippingTool.exe")) {
                OutputDebug("`nSnipping Tool Overlay activated.")
                return true
            } else {
                OutputDebug("`nSnipping Tool Overlay did not become active.")
                return false
            }
        } else {
            OutputDebug("`nSnipping Tool Overlay not found.")
        }

        return false
    }

    ; Looped cycle that checks for the presence of a toast notification ahk_exe ShellExperienceHost.exe, title of New notification
    static CheckAndClickToast() {
        local existResult := WinExist("ahk_exe ShellExperienceHost.exe")
        local MainElement := 0
        local button := 0

        ; Check if the toast has already been clicked
        if (this.haveClickedToast) {
            OutputDebug("`nToast already clicked. Exiting.")
            return
        }

        if (existResult)
        {
            try {
                MainElement := UIA.ElementFromHandle(this.toastString)
                if (isObject(MainElement))
                {
                    Condition := UIA.CreateCondition("AutomationId", "VerbButton")
                    try {
                        button := MainElement.FindFirst(Condition, UIA.TreeScope.Subtree)
                    }

                    if (isObject(button) && button.Name == "Markup and share")
                    {
                        Sleep(250) ; Wait for the button to be ready. Even when not automating, the button is not clickable immediately.
                        button.Click()
                        this.haveClickedToast := true ; Set the flag to true to prevent clicking again unless the click fails
                        
                        local activateResult := this.WaitAndActivateSnipWindow() ; Wait for the snipping tool window to exist and activate it
                        if (activateResult) {
                            this.CancelToastCheckTimer() ; Cancel the timer
                            OutputDebug("`nToast found and clicked.")
                            return true
                        } else {
                            ; If it timed out
                            OutputDebug("`nFailed to activate Snipping Tool Overlay after clicking toast.")
                            this.haveClickedToast := false ; Reset the flag to allow for another click
                            return false
                        }
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

    ; The SetTimer function was not accepting the bound function directly, so we create a static bound function to use instead for some reason
    static boundCheckToast := SnippingToolController.CheckAndClickToast.Bind(SnippingToolController)
    static boundCancelTimer := SnippingToolController.CancelToastCheckTimer.Bind(SnippingToolController)

    static CancelToastCheckTimer() {
        ; Cancels the timer
        SetTimer(this.boundCheckToast, 0)
        this.isToastClickTimerRunning := false
        OutputDebug("`nTimer cancelled.")
    }
   

    static CallFunctionWithTimeout(IntervalMs, TimeoutSeconds) {
        if (this.isToastClickTimerRunning) {
            OutputDebug("`nTimer already running. Exiting CallFunctionWithTimeout.")
            return
        }

        this.isToastClickTimerRunning := true
        OutputDebug("`nSetting up timers. isToastClickTimerRunning = true")

        ; Use SetTimer with the BoundFunc objects
        SetTimer(this.boundCheckToast, IntervalMs)
        ; A timeout timer to cancel the toast check after a certain time if not clicked
        SetTimer(this.boundCancelTimer, (TimeoutSeconds * -1000))
    }
}
