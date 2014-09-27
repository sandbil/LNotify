{==============================================================================|
| Project : Notes/Delphi class library                           | 3.10        |
|==============================================================================|
| Content:                                                                     |
|==============================================================================|
| The contents of this file are subject to the Mozilla Public License Ver. 1.0 |
| (the "License"); you may not use this file except in compliance with the     |
| License. You may obtain a copy of the License at http://www.mozilla.org/MPL/ |
|                                                                              |
| Software distributed under the License is distributed on an "AS IS" basis,   |
| WITHOUT WARRANTY OF ANY KIND, either express or implied. See the License for |
| the specific language governing rights and limitations under the License.    |
|==============================================================================|
| Initial Developers of the Original Code are:                                 |
|   Sergey Kolchin (Russia) skolchin@yahoo.com                                 |
|   Sergey Kucherov (Russia)                                                   |
|   Sergey Okorochkov (Russia)                                                 |
| All Rights Reserved.                                                         |
|   Last Modified:                                                             |
|     17.01.2002, Sergey Kolchin                                               |
|==============================================================================|
|==============================================================================|
| Include file for Delphi and N/D version defines                              |
|                                                                              |
| Delphi versions: D1 to D7, DELPHI_VERSION                                    |
| N/D versions: ND_310, NDLIB_VERSION                                          |
|==============================================================================|}

{$IFNDEF WIN32}
-- This unit is for Windows 32 environment
{$ENDIF}

// Delphi versions
{$IFDEF VER150}
  {$DEFINE D7}
  {$DEFINE D6}
  {$DEFINE D5}
  {$DEFINE D4}
{$ELSE}
  {$IFDEF VER140}
    {$DEFINE D6}
    {$DEFINE D5}
    {$DEFINE D4}
  {$ELSE}
    {$IFDEF VER130}
      {$DEFINE D5}
      {$DEFINE D4}
    {$ELSE}
      {$IFDEF VER120}
        {$DEFINE D4}
      {$ELSE}
        {$DEFINE D3}
      {$ENDIF}
    {$ENDIF}
  {$ENDIF}
{$ENDIF}

// N/D version
{$DEFINE ND_310}

