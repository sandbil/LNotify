---
title: Lotus Notes Notify
description: simple 
author: sandbil
tags: delphi, Lasarus, Lotus Notes, Lotus Domino

---
LNotify
=========
This programm notify you about a new records at the NSF databases (so email or a document).
It  monitor UnReadmarks in the your set NSF databases and show notification window with new record clickable captions.
After, Click on caption to open document in the  Lotus Notes client and see a new it. 
This used simple RC4 to encode  for remembering password.

This tested with version 8.5 

## Requirements
  Delphi 7
  Before compiling install component NextSuite.
  The client Lotus Notes (v.8.x.x) has to be installed and set.
    
## Usage

   Setting in the "LNotify.INI" file:

[Connect]
Server=enter your server Lotus Domino
SedFolder=enter your default folder on server (ex.- Referent)
Group= enter user group (not necessary)
User=
[Timer]
CheckTime(min)=15
TimeShowHint(sec)=5
[Mail]
MailFile=mail\administ.nsf
[Database]
NSF01=Referent\example.nsf

Server - your Domino server
SedFolder - default folder on Domino server where is your databases
Group - user's group (not necessary parameter)
User - it fill after checked CheckBox "Save password" 
 
[![screenshot1](/public/screenshot_th1.png)](/public/screenshot1.png)
[![screenshot2](/public/screenshot_th2.png)](/public/screenshot2.png)
