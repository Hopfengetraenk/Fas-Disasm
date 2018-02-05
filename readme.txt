Fas-Disassembler for Visuallisp 0.8
===================================

Autolisp is a programming language for AutoCAD.
Lisp sourcecode files have the extension LSP and 
the compiled Lisp-scripts FAS.
This program will decrypt the resource part of fas 
and fsl files (write it to disk) and disassemble the it.
With help of the disassembling you can see exactly :) 
what the program does, how thing are done and you 
can change some things with help of the offset and 
an Hexeditor.


Keys:
Enter,Space, Doubleclick on a line to 
	jump to a offset that is in the Disam
	Like "Goto 0213"

Backspace, Num'-' to 
	go back to old offset

´(key beside Backspace), Num'+' to 
	go forward



All internal lisp programs of Visuallisp are store in the 
resource section of vllib.dll (or vl.arx in Autocad LT).
Use "Resouce Hacker" or "Exescope" to dump 
these resources to disk and rename them to  *.fsl.

Well this is no decompiler so far, 
but the next step will be one... Hehe :)

Ok just for fun I'm writing the data from the decompiler colum to
some *.lsp. This far from being compilable code and I really advice
you to delete it and use the *.Txt instead which contains all
information.



When hexediting fas-files maybe this two commands come in handy:

<NOP> No Operation 
	Good for deleting (= Nop out) unwanted instructions
	Bytecode: 20h (Just Space) [or 62 or 63]
	Parameters: none

<jmp> short Jump to offset
	Good for jumping over unwanted instructions
	i.e. Fix Conditional Jump (=IF)
	Bytecode: 0Fh  [57h for FAR Jump]
	Size: 3 Byte [5 Byte]
	Parameters: 1
	Word 2 Bytes [Dword 4 Bytes] specify Bytes to Skip
	Example: '0F 0000' will point to the next
	 	 instruction after jmp

At the moment there is no documentation file for fas-commands. 
To get more information about the command please look at the function 
'InterpretStream()' in FasFile.cls and learn from the fas-disassembling.

---------------------------------------------
Version history:
0.9
  * Support for 'AutoCAD PROTECTED LISP file' *.lsp
  
0.8
  * Support for vlx-files (vlx-splitter)
  * forward backward buttons for navigation add
  
0.7
  * opcodesnames for fsl-disassembling improved
  * add loop recognission
  * decompilation column added
  
0.6
  * added Quickjump function for Hexworkshop
  * added Case insensitiv search Checkbox 
  
0.5
  * FasCommand disassembling improved
  * small bug's fixed
  
0.4
  * First public Version

0.3..0.1
  * internal alpha-version 
   (based on AutoLisp Resource Decrypter V0.9)