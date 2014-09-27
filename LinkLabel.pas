{******************************************************************************}
{ Jazarsoft LinkLabel Component                                                }
{******************************************************************************}
{                                                                              }
{ VERSION      : 1.0                                                           }
{ AUTHOR       : James Azarja                                                  }
{ CREATED      : 10 July 2000                                                  }
{ WEBSITE      : http://jazarsoft.cjb.net/                                     }
{ SUPPORT      : support@jazarsoft.cjb.net                                     }
{ BUG-REPORT   : bugreport@jazarsoft.cjb.net                                   }
{ COMMENT      : comment@jazarsoft.cjb.net                                     }
{ LEGAL        : Copyright (C) 2000 Jazarsoft.                                 }
{                                                                              }
{******************************************************************************}
{ NOTE         :                                                               }
{                                                                              }
{ This code may be used and modified by anyone so long as  this header and     }
{ copyright  information remains intact.                                       }
{                                                                              }
{ The code is provided "as-is" and without warranty of any kind,               }
{ expressed, implied or otherwise, including and without limitation, any       }
{ warranty of merchantability or fitness for a  particular purpose.            }
{                                                                              }
{ In no event shall the author be liable for any special, incidental,          }
{ indirect or consequential damages whatsoever (including, without             }
{ limitation, damages for loss of profits, business interruption, loss         }
{ of information, or any other loss), whether or not advised of the            }
{ possibility of damage, and on any theory of liability, arising out of        }
{ or in connection with the use or inability to use this software.             }
{                                                                              }
{******************************************************************************}

unit LinkLabel;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls;


type
  tLinkType  = (ltfile,ltftp,ltgopher,lthttp,lthttps,ltmailto,ltnews,lttelnet,ltwais,ltNotes);
  TLinkLabel = class(TLabel)
  private
   FLinkType  : tLinkType;
   FHyperLink : String;

   Procedure SetLinkType(LinkType:tLinkType);
   Procedure SetHyperLink(HyperLink:String);
  protected
   procedure EvLButtonDown(var M : TWMMouse); message WM_LBUTTONDOWN;
  public
   constructor create(AOwner:TComponent);override;
   destructor destroy;override;
  published
   property LinkType : tlinkType read FLinkType Write SetLinkType;
   property HyperLink: String read Fhyperlink write SetHyperLink;
  end;

procedure Register;

implementation

Function ShellExecute(hWnd:HWND;lpOperation:Pchar;lpFile:Pchar;lpParameter:Pchar;
                      lpDirectory:Pchar;nShowCmd:Integer):Thandle; Stdcall;
External 'Shell32.Dll' name 'ShellExecuteA';

Procedure tLinkLabel.SetLinkType(LinkType:tLinkType);
Begin
 If FLinkType<>LinkType then FlinkType:=LinkType;
End;

Procedure tLinkLabel.SetHyperLink(HyperLink:String);
Begin
 FHyperLink:=HyperLink;
 if caption='' then
  Caption:=FHyperLink;
End;

Constructor tLinkLabel.create(AOwner:TComponent);
Begin
 inherited Create(AOwner);
 Parent:=(AOwner as tForm);
 LinkType:=ltHttp;
 Font.Style:=[fsUnderline];
 Font.Color:=clblue;
 Cursor:=crhandPoint;
 Caption:='';
End;

Destructor tLinkLabel.destroy;
Begin
 inherited destroy;
End;


procedure tLinkLabel.EvLButtonDown(var M : TWMMouse);
var
 commandline : string;
Begin
 if linktype=ltfile then commandline:='file://'+hyperlink else
 if linktype=ltftp then commandline:='ftp://'+hyperlink else
 if linktype=ltgopher then commandline:='gopher://'+hyperlink else
 if linktype=lthttp then commandline:='http://'+hyperlink else
 if linktype=lthttps then commandline:='https://'+hyperlink else
 if linktype=ltmailto then commandline:='mailto:'+hyperlink else
 if linktype=ltnews then commandline:='news:'+hyperlink else
 if linktype=lttelnet then commandline:='telnet:'+hyperlink else
 if linktype=ltwais then commandline:='wais:'+hyperlink;
 if linktype=ltNotes then commandline:='Notes://'+hyperlink;

 ShellExecute(Parent.Handle,'Open',pchar(commandline),Nil,nil,SW_SHOWNORMAL);
End;

procedure Register;
begin
  RegisterComponents('Jazarsoft', [TLinkLabel]);
end;

end.
