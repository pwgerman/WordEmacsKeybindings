Attribute VB_Name = "NewMacros"
Sub EmacsCustomKeybind()
Attribute EmacsCustomKeybind.VB_Description = "Assigns emacs and other useful keybindings to Keyboard Customization."
Attribute EmacsCustomKeybind.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Emacs1"
'
' Emacs1 Macro
' Assigns emacs and other useful keybindings to Keyboard Customization.
'
    CustomizationContext = NormalTemplate
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyB, wdKeyOption, wdKeyControl), _
        KeyCategory:=wdKeyCategoryCommand, Command:="WordLeft"
    CustomizationContext = NormalTemplate
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyB, wdKeyShift, wdKeyOption, _
        wdKeyControl), KeyCategory:=wdKeyCategoryCommand, Command:= _
        "WordLeftExtend"
    CustomizationContext = NormalTemplate
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyP, wdKeyShift, wdKeyControl), _
        KeyCategory:=wdKeyCategoryCommand, Command:="LineUpExtend"
    CustomizationContext = NormalTemplate
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyN, wdKeyShift, wdKeyControl), _
        KeyCategory:=wdKeyCategoryCommand, Command:="LineDownExtend"
    CustomKeybind_DeleteChar
End Sub

Sub DeleteChar()
Attribute DeleteChar.VB_Description = "Forward delete one char to the right"
Attribute DeleteChar.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.EditBack1"
'
' EditBack1 Macro
' Forward delete one char to the right
'
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.TypeBackspace
End Sub
Sub CustomKeybind_DeleteChar()
Attribute CustomKeybind_DeleteChar.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Add_delete_macro_shortcut"
'
' Add_delete_macro_shortcut Macro
'
'
    CustomizationContext = NormalTemplate
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyD, wdKeyControl), KeyCategory:= _
        wdKeyCategoryMacro, Command:="DeleteChar"
End Sub
