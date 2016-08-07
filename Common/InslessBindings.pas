{                                                                           }
{ File:       INSlessBindings.pas                                           }
{ Function:   Key bindings for Delphi on a computer with a keyboard that    }
{             does not have an INS key.                                     }
{                                                                           }
{             Changes the following keybindings and menu commands:          }
{               Edit|Copy            Ctrl+C                                 }
{               Edit|Paste           Ctrl+V                                 }
{               Edit|Cut             Ctrl+X                                 }
{                                                                           }
{             Additionally, introduces the following keybindings:           }
{               Shift+Ctrl+Alt+C     Copy and append to clipboard           }
{               Shift+Ctrl+Alt+X     Cut and append to clipboard            }
{                                                                           }
{             This is a partial keybinding, so the other keys are not       }
{             affected.                                                     }
{                                                                           }
{ Language:   Delphi 2007 and above                                         }
{ Author:     Rudy Velthuis                                                 }
{ Version:    1.0                                                           }
{ Copyright:  © 2012 Rudy Velthuis                                          }
{                                                                           }
{ License:    Redistribution and use in source and binary forms, with or    }
{             without modification, are permitted provided that the         }
{             following conditions are met:                                 }
{                                                                           }
{             * Redistributions of source code must retain the above        }
{               copyright notice, this list of conditions and the following }
{               disclaimer.                                                 }
{             * Redistributions in binary form must reproduce the above     }
{               copyright notice, this list of conditions and the following }
{               disclaimer in the documentation and/or other materials      }
{               provided with the distribution.                             }
{                                                                           }
{ Disclaimer: THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDER "AS IS"     }
{             AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT     }
{             LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND     }
{             FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO        }
{             EVENT SHALL THE COPYRIGHT OWNER OR CONTRIBUTORS BE LIABLE     }
{             FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY,     }
{             OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO,      }
{             PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE,     }
{             DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED    }
{             AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT   }
{             LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE)        }
{             ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF   }
{             ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.                    }
{                                                                           }

unit INSlessBindings;

interface

procedure Register;

implementation

uses
  Windows, Classes, SysUtils, ToolsAPI, Menus, Forms, Dialogs, Controls, ActiveX;

type
  TINSlessBinding = class(TNotifierObject, IUnknown, IOTANotifier, IOTAKeyboardBinding)
  protected
    procedure ClipCopy(const Context: IOTAKeyContext; KeyCode: TShortcut; var BindingResult: TKeyBindingResult);
    procedure ClipCut(const Context: IOTAKeyContext; KeyCode: TShortcut; var BindingResult: TKeyBindingResult);
    procedure ClipPaste(const Context: IOTAKeyContext; KeyCode: TShortcut; var BindingResult: TKeyBindingResult);
    procedure ClipCopyAppend(const Context: IOTAKeyContext; KeyCode: TShortcut; var BindingResult: TKeyBindingResult);
    procedure ClipCutAppend(const Context: IOTAKeyContext; KeyCode: TShortcut; var BindingResult: TKeyBindingResult);
  public
    constructor Create;
    function GetBindingType: TBindingType;
    function GetDisplayName: string;
    function GetName: string;
    procedure BindKeyboard(const BindingServices: IOTAKeyBindingServices);
  end;

resourcestring
  SINSlessExtension = 'Cut, Copy and Paste without INS key';

// Do not localize.
const
  SBindingsName = 'Velthuis.INSKeylessBindings';

procedure Register;
begin
  (BorlandIDEServices as IOTAKeyBoardServices).AddKeyboardBinding(TINSlessBinding.Create as IOTAKeyboardBinding);
end;

{ TINSlessBinding }

{ Do not localize the following strings }

procedure TINSlessBinding.BindKeyboard(const BindingServices: IOTAKeyBindingServices);
begin
  BindingServices.AddKeyBinding([ShortCut(Ord('C'), [ssCtrl])], ClipCopy, nil, kfImplicitShift, '', 'EditCopyItem');
  BindingServices.AddKeyBinding([ShortCut(Ord('V'), [ssCtrl])], ClipPaste, nil, kfImplicitShift, '', 'EditPasteItem');
  BindingServices.AddKeyBinding([ShortCut(Ord('X'), [ssCtrl])], ClipCut, nil, kfImplicitShift, '', 'EditCutItem');
  BindingServices.AddKeyBinding([ShortCut(Ord('C'), [ssShift, ssCtrl, ssAlt])], ClipCopyAppend, nil);
  BindingServices.AddKeyBinding([ShortCut(Ord('X'), [ssShift, ssCtrl, ssAlt])], ClipCutAppend, nil);
  BindingServices.AddMenuCommand(mcClipCopy, ClipCopy, nil);
  BindingServices.AddMenuCommand(mcClipCut, ClipCut, nil);
  BindingServices.AddMenuCommand(mcClipPaste, ClipPaste, nil);
end;

procedure TINSlessBinding.ClipCopyAppend(const Context: IOTAKeyContext; KeyCode: TShortcut;
  var BindingResult: TKeyBindingResult);
begin
  Context.EditBuffer.EditBlock.Copy(True);
  BindingResult := krHandled;
end;

procedure TINSlessBinding.ClipCutAppend(const Context: IOTAKeyContext; KeyCode: TShortcut;
  var BindingResult: TKeyBindingResult);
begin
  Context.EditBuffer.EditBlock.Cut(True);
  BindingResult := krHandled;
end;

procedure TINSlessBinding.ClipCopy(const Context: IOTAKeyContext; KeyCode: TShortcut;
  var BindingResult: TKeyBindingResult);
begin
  Context.EditBuffer.EditBlock.Copy(False);
  BindingResult := krHandled;
end;

procedure TINSlessBinding.ClipCut(const Context: IOTAKeyContext; KeyCode: TShortcut;
  var BindingResult: TKeyBindingResult);
begin
  Context.EditBuffer.EditBlock.Cut(False);
  BindingResult := krHandled;
end;

procedure TINSlessBinding.ClipPaste(const Context: IOTAKeyContext; KeyCode: TShortcut;
  var BindingResult: TKeyBindingResult);
begin
  Context.EditBuffer.EditPosition.Paste;
  BindingResult := krHandled;
end;

constructor TINSlessBinding.Create;
begin
  inherited;
end;

function TINSlessBinding.GetBindingType: TBindingType;
begin
  Result := btPartial;
end;

function TINSlessBinding.GetDisplayName: string;
begin
  Result := SINSlessExtension;
end;

function TINSlessBinding.GetName: string;
begin
  Result := SBindingsName;
end;

end.
