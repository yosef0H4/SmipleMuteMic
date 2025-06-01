#SingleInstance Force

; Main application entry point
app := MicrophoneController()
app.Initialize()

; Main controller class
class MicrophoneController {
    
    __New() {
        this.configManager := ConfigManager()
        this.audioManager := AudioDeviceManager()
        this.configGUI := ""
        this.selectedDevice := ""
        this.currentHotkey := ""
        this.defaultHotkey := "CapsLock"
        this.showTooltips := true
        this.enableBeeps := true
    }
    
    Initialize() {
        ; Load configuration or show setup if none exists
        if (this.configManager.ConfigExists()) {
            this.LoadConfiguration()
        } else {
            this.ShowConfiguration()
        }
        
        ; Add tray menu
        A_TrayMenu.Add("Configure Microphone", (*) => this.ShowConfiguration())
    }
    
    LoadConfiguration() {
        this.selectedDevice := this.configManager.LoadDevice()
        this.currentHotkey := this.configManager.LoadHotkey()
        this.showTooltips := this.configManager.LoadTooltips()
        this.enableBeeps := this.configManager.LoadBeeps()
        
        if (!this.selectedDevice) {
            this.ShowConfiguration()
            return
        }
        
        ; Set up the hotkey
        this.SetupHotkey()
    }
    
    SetupHotkey() {
        ; Remove old hotkey if it exists
        if (this.currentHotkey) {
            try {
                Hotkey(this.currentHotkey, "Off")
            } catch {
                ; Ignore errors when removing non-existent hotkeys
            }
        }
        
        ; Use default if no hotkey is configured
        if (!this.currentHotkey) {
            this.currentHotkey := this.defaultHotkey
        }
        
        ; Set CapsLock state for CapsLock hotkey
        if (this.currentHotkey == "CapsLock") {
            SetCapsLockState("AlwaysOff")
        }
        
        ; Create the new hotkey
        Hotkey(this.currentHotkey, (*) => this.ToggleMute())
        if (this.showTooltips) {
            ToolTip("Hotkey set to: " . this.currentHotkey, , , 1)
            SetTimer(() => ToolTip(), -2000)
        }
    }
    
    ShowConfiguration() {
        this.configGUI := ConfigGUI(this)
        this.configGUI.Show()
    }
    
    ToggleMute() {
        if (this.selectedDevice) {
            try {
                this.audioManager.ToggleMute(this.selectedDevice)
                this.ShowMuteStatus()
            } catch as e {
                MsgBox("Error: " e.Message)
            }
        } else {
            MsgBox("No microphone selected. Please configure the script first.")
            this.ShowConfiguration()
        }
    }
    
    ShowMuteStatus() {
        isMuted := this.audioManager.IsMuted(this.selectedDevice)
        
        ; Show tooltip if enabled
        if (this.showTooltips) {
            ToolTip("Microphone " (isMuted ? "Muted" : "Unmuted"), , , 1)
            SetTimer(() => ToolTip(), -2000)
        }
        
        ; Audio feedback if enabled
        if (this.enableBeeps) {
            if (isMuted) {
                SoundBeep(100, 100)
            } else {
                SoundBeep(300, 100)
            }
        }
    }
    
    SaveConfiguration(deviceName, hotkeyString, showTooltips, enableBeeps) {
        ; Validate device selection
        if (!deviceName || deviceName == "") {
            MsgBox("No device name provided", "Error", "Icon!")
            return
        }
        
        if (InStr(deviceName, "No microphones detected")) {
            MsgBox("Invalid device selection", "Error", "Icon!")
            return
        }
        
        ; Validate hotkey
        if (!hotkeyString || hotkeyString == "") {
            hotkeyString := this.defaultHotkey
        }
        
        ; Save all settings
        this.selectedDevice := deviceName
        this.currentHotkey := hotkeyString
        this.showTooltips := showTooltips
        this.enableBeeps := enableBeeps
        
        this.configManager.SaveDevice(deviceName)
        this.configManager.SaveHotkey(hotkeyString)
        this.configManager.SaveTooltips(showTooltips)
        this.configManager.SaveBeeps(enableBeeps)
        
        this.SetupHotkey()
        
        if (this.configGUI) {
            this.configGUI.Close()
            this.configGUI := ""
        }
        
        if (this.showTooltips) {
            ToolTip("Configuration saved!`nMicrophone: " . deviceName . "`nHotkey: " . hotkeyString . "`nReloading script...", , , 1)
            SetTimer(() => ToolTip(), -1500)
        }
        
        ; Reload the script to ensure all changes take effect
        SetTimer(() => Reload(), -700)
    }
    
    CancelConfiguration() {
        if (this.configGUI) {
            this.configGUI.Close()
            this.configGUI := ""
        }
        
        if (!this.selectedDevice) {
            MsgBox("No microphone selected. The script will exit.")
            ExitApp()
        }
    }
}

; Configuration management class
class ConfigManager {
    
    __New() {
        this.configFile := A_ScriptDir "\AudioConfig.ini"
    }
    
    ConfigExists() {
        return FileExist(this.configFile)
    }
    
    LoadDevice() {
        try {
            return IniRead(this.configFile, "Settings", "Device", "")
        } catch {
            return ""
        }
    }
    
    LoadHotkey() {
        try {
            return IniRead(this.configFile, "Settings", "Hotkey", "")
        } catch {
            return ""
        }
    }
    
    LoadTooltips() {
        try {
            value := IniRead(this.configFile, "Settings", "ShowTooltips", "1")
            return (value == "1" || value == "true")
        } catch {
            return true  ; Default to enabled
        }
    }
    
    LoadBeeps() {
        try {
            value := IniRead(this.configFile, "Settings", "EnableBeeps", "1")
            return (value == "1" || value == "true")
        } catch {
            return true  ; Default to enabled
        }
    }
    
    SaveDevice(deviceName) {
        IniWrite(deviceName, this.configFile, "Settings", "Device")
    }
    
    SaveHotkey(hotkeyString) {
        IniWrite(hotkeyString, this.configFile, "Settings", "Hotkey")
    }
    
    SaveTooltips(enabled) {
        IniWrite(enabled ? "1" : "0", this.configFile, "Settings", "ShowTooltips")
    }
    
    SaveBeeps(enabled) {
        IniWrite(enabled ? "1" : "0", this.configFile, "Settings", "EnableBeeps")
    }
}

; Audio device management class
class AudioDeviceManager {
    
    GetMicrophoneDevices() {
        microphones := []
        
        loop {
            try {
                deviceName := SoundGetName(, A_Index)
                
                if (this.IsMicrophoneDevice(deviceName)) {
                    microphones.Push(deviceName)
                }
            } catch {
                break
            }
        }
        
        return microphones
    }
    
    IsMicrophoneDevice(deviceName) {
        microphoneKeywords := ["mic", "input", "recording", "VoiceMeeter", "headset", "array"]
        
        for keyword in microphoneKeywords {
            if (InStr(deviceName, keyword, false)) {
                return true
            }
        }
        
        return false
    }
    
    ToggleMute(deviceName) {
        SoundSetMute(-1, "", deviceName)
    }
    
    IsMuted(deviceName) {
        return SoundGetMute("", deviceName)
    }
}

; Compact Configuration GUI class
class ConfigGUI {
    
    __New(controller) {
        this.controller := controller
        this.gui := ""
        this.deviceList := ""
        this.hotkeyControl := ""
        this.tooltipCheckbox := ""
        this.beepCheckbox := ""
        this.audioManager := AudioDeviceManager()
    }
    
    Show() {
        this.CreateGUI()
        this.PopulateDeviceList()
        this.gui.Show()
    }
    
    CreateGUI() {
        ; Create compact GUI with minimal margins
        this.gui := Gui("", "Microphone Configuration")
        
        ; Set tight margins for compact layout
        this.gui.MarginX := 8
        this.gui.MarginY := 6
        
        ; === MICROPHONE SELECTION SECTION ===
        this.gui.Add("Text", "w350 cBlue", "Microphone Selection")
        this.gui.Add("Text", "w350 y+2", "Select the microphone you want to control:")
        
        ; Compact device list
        this.deviceList := this.gui.Add("ListBox", "w350 h120 y+4 vSelectedDevice")
        
        ; === HOTKEY CONFIGURATION SECTION ===
        this.gui.Add("Text", "w350 y+8 cBlue", "Hotkey Configuration")
        this.gui.Add("Text", "w350 y+2", "Current: " . (this.controller.currentHotkey ? this.controller.currentHotkey : "CapsLock (default)"))
        
        this.gui.Add("Text", "w350 y+4", "Set hotkey (F1, Ctrl+Space, Alt+M, etc.):")
        
        ; Hotkey input with reset button on same line
        this.hotkeyControl := this.gui.Add("Hotkey", "w200 y+2 vNewHotkey", this.controller.currentHotkey ? this.controller.currentHotkey : this.controller.defaultHotkey)
        this.gui.Add("Button", "x+8 yp w70 h21", "Reset").OnEvent("Click", (*) => this.ResetHotkey())
        
        this.gui.Add("Text", "w350 xm y+6 c0x808080", "Note: Some combinations may not work if already used by the system.")
        
        ; === FEEDBACK OPTIONS SECTION ===
        this.gui.Add("Text", "w350 y+8 cBlue", "Feedback Options")
        
        ; Tooltip checkbox
        this.tooltipCheckbox := this.gui.Add("CheckBox", "w350 y+4 vShowTooltips", "Show mute status tooltips")
        this.tooltipCheckbox.Value := this.controller.showTooltips
        
        ; Beep checkbox
        this.beepCheckbox := this.gui.Add("CheckBox", "w350 y+2 vEnableBeeps", "Enable audio beeps (low=muted, high=unmuted)")
        this.beepCheckbox.Value := this.controller.enableBeeps
        
        ; === BUTTONS SECTION ===
        this.gui.Add("Button", "w80 y+10 Default", "Save").OnEvent("Click", (*) => this.SaveConfig())
        this.gui.Add("Button", "x+8 yp w80", "Cancel").OnEvent("Click", (*) => this.CancelConfig())
        
        ; Show with auto-sizing, then hide to get final size
        this.gui.Show("AutoSize")
        this.gui.Hide()
    }
    
    PopulateDeviceList() {
        microphones := this.audioManager.GetMicrophoneDevices()
        
        if (microphones.Length == 0) {
            this.deviceList.Add(["No microphones detected automatically"])
            this.gui.Add("Text", "w350 y+4 c0x808080", "Try running as administrator if your microphone isn't listed.")
        } else {
            for mic in microphones {
                this.deviceList.Add([mic])
            }
            
            ; Pre-select current device if it exists in the list
            if (this.controller.selectedDevice) {
                for index, mic in microphones {
                    if (mic == this.controller.selectedDevice) {
                        this.deviceList.Choose(index)
                        break
                    }
                }
            }
        }
    }
    
    ResetHotkey() {
        this.hotkeyControl.Value := this.controller.defaultHotkey
    }
    
    SaveConfig() {
        ; Get selected device
        selectedText := this.deviceList.Text
        if (!selectedText) {
            MsgBox("Please select a microphone from the list.", "No Selection", "Icon!")
            return
        }
        
        ; Get hotkey value
        newHotkey := this.hotkeyControl.Value
        if (!newHotkey || newHotkey == "") {
            newHotkey := this.controller.defaultHotkey
        }
        
        ; Get checkbox values
        showTooltips := this.tooltipCheckbox.Value
        enableBeeps := this.beepCheckbox.Value
        
        ; Save all configurations
        this.controller.SaveConfiguration(selectedText, newHotkey, showTooltips, enableBeeps)
    }
    
    CancelConfig() {
        this.controller.CancelConfiguration()
    }
    
    Close() {
        if (this.gui) {
            this.gui.Destroy()
        }
    }
}