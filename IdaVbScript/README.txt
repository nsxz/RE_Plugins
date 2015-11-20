

put idavbscript.plw and idavbplugin.dll in IDA plugin
directory.

regsvr32 idavbplugin.dll

start ida 

edit menu -> plugins -> dzzies Ida Plugin


dont have time to write real docs for this
you will have to read the source for documentation 
if you cant figure it out from the UI.

the plw stub was compiled with VS 2008 express 
using hte ida 5.5 sdk

Note: the CMD socket might be a bit crashy its super old
I may remove it, I have disabled it from automatically starting up
I will switch any apps that use it over to using the OllySyncBridge
which uses the new and more stable IDASrvr plugin to proxy the 
network requests to IDA (nice small external easily debuggable better
backend tech)







 

