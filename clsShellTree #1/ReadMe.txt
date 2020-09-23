
Warning: If you wanted a production-ready folder browser, sorry to disappoint you.

This is deliberately left at an experimental stage
 - to provide a universal platform, from which different solutions can be derived
   ie include code in exe app, compile as seperate ocx , 
   multiselect checkbox-style folder chooser, etc
 - to delve further into shell's mysteries (explorer's right pane & IShellView)
 - learn & discover (my knowledge of shell programming consisted of calling ShellExecuteEx)

I use a stripped down version in the same app, I intentionally coded the 
ownerdrawn treeview for, and it's a pure pleasure, so this is despite warnings 
functional (but not a three liner). 

You are welcome to implement missing functionality, enhance or repackage it and post it as your! work. 
(especially network/FTP code, I work on a single PC and I don't use (IE)explorer for the web)

To Whom it may concern,

I started off with studying (aka pasting) Brad Martinez different samples 
(eh,what is a pidl?), modernized it a bit (wow,he did it years ago and still 
there's scarce (VB) code on the subject), introduced 'settlement' API's to VB 
(a shame,documented since 5 years and no published VB code), (re-)invented doing 
data transfer by invoking item's contextmenu, implemented D&D (if target is on 
different drive it's vbDropEffectMove etc) and felt really smart when I eh.. 
borrowed the localized! drag contextmenu from shell.
(need shell resources? -> Old Code/modDropContextMenu.bas)

I now humbly forget the time spent getting OpenNode() to work plus reacting to 
shell changes and achieved after some time a tree that looked and behaved more 
or less like the real (Explorer) thing. Considered it superior to anything I 
came across before ,wanted to post it, quit VB and maybe start coding with C#  
next winter. (no 1003. 'hide app from taskbar' and fewer MoMoYa's, I hope)

Well, I stumbled upon Shell Explorer's Cookbook (C++ but recommended! reading). 
Nikolaos describing using some obscure COM interfaces (us VB's are sitting in 
top of them) to let the shell do much of the gruntwork by forwarding calls to the 
respective folders. I again got hooked, OK nice in C++, is this possible in VB?

As you can see most of it is, but it took some turns, dead ends and inelegant 
solutions. Compare EdanMo's black magic pIDataObjectFromVB() to my 
reconstructing of the IDataobject (-> Old Code/GetIDataObjectFromVB.txt).
This is the only VB code around that reads CFSTR_SHELLIDLIST format,
so it still may be of value, however I damn fail writing the format.

PS: 
Although lately no GPF's occurred a decent(== better than mine) exceptionhandler is a MUST! 
(Misprogrammed shell namespace extensions, explorer quirks are outside of your reach)

PPS:
If functions declared in the typelib don't work as expected, first blame it on my additions.
 (-> ISHF_Ex.odl) , some parts need checking.

PPPS: 
I,only,premiere,no published code etc : Google depicts the corners of my (coding) world!


Waiting for your bug reports (Win2K untested)

OrlandoCurioso

