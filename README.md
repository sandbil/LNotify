---
title: Lotus Notes Notify
description: simple 
author: sandbil
tags: delphi, Lasarus, Lotus Notes, Lotus Domino

---
LNotify
=========
This program notify you about a new records at the NSF databases (so email or a document).  
It  monitor UnReadmarks in the your set NSF databases and show notification window with new record clickable captions.  
After, Click on caption to open document in the  Lotus Notes client and see a new it.  
This used simple RC4 to encode  for remembering password. 
This example program based on class library, for detail see file Ln\Readme.txt (Sergey Kolchin..)


This tested with version 8.5.   

## Requirements
  Delphi 7
  Before compiling install component NextSuite.
  The client Lotus Notes (v.8.x.x) has to be installed and set.  
  After building you have to copy LNotify.exe and LNotify.ini to folder with notes client. 

  If you want to have portable version, you must have following structure:
  
  ..\data\names.nsf  
  ..\icc\icclib\icclib.dll  
  ..\icc\osslib\libeay32.dll  
  icudt34l.dat  
  js32.dll  
  ndgts.dll  
  nlsccstr.dll  
  nnet.dll  
  nnotes.dll  
  nntcp.dll  
  nsdhelp.dll  
  nstrings.dll  
  ntcp.dll  
  nxmlproc.dll  
  LNotify.exe  
  LNotify.ini  
  notes.ini  
  ltsci3.tlb  
    
    
## Usage

   Setting in the "LNotify.INI" file:  

*  [Connect]  
*  Server=SEDSRV  
*  SedFolder=SedFolder  
*  Group= SedUsers  
*  User=  
*  [Timer]  
*  CheckTime(min)=15  
*  TimeShowHint(sec)=5  
*  [Mail]  
*  MailFile=mail\administ.nsf  
*  [Database]  
*  NSF01=Referent\example1.nsf  
*  NSF02=Referent\example2.nsf  
   
    
  
*  Server - your Domino server  
*  SedFolder - default folder on Domino server where is your databases  
*  Group - user's group (not necessary parameter)  
*  User - it fill after checked CheckBox "Save password"   
*  CheckTime(min)= time for checking (default every 15 minuts)  
*  TimeShowHint(sec)= time for showing notifing's window (default 5 second)   
*  MailFile= mail file, auto set after user authentification   
*  NSF01= your NSF file for monitoring UnReadmarks  
*  NSF02= ..  
  
   
[![screenshot1](/setting.png)](/setting.png)
[![screenshot2](/windowNotify.png)](/windowNotify.png)
