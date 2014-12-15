
httpsendhook.dll hooks the following wininet api calls:

HttpOpenRequest,InternetConnect,InternetReadFile,InternetCrackUrl,HttpSendRequest

It will log the key information used in these api at a point where it doesnt
matter if the call was made to a https url or not.

Webpages downloaded will be saved (in fragments) to c:\pages
Form posts will be saved to c:\posts
Many other messages such as servers connected to, pages requested etc will
be sent to the api_logger.exe interface if it is open. data logged to this
UI is prefixed by the decimal_processid.offset_called_from>. This data is also
logged to c:\wininet_hooks.txt too. (always in append mode delete old files)

api_logger.exe can inject this dll into the process you want. You can drag and drop
the target exe into the top text box and hit inject, or you can select an already
running process using hte select pid button. (then hit inject)

you dont need api_logger running, if you inject the dll another way such as adding
the hook dll to the appinit_dll setting.

this thing is written very fast and dirty, but it is a useful tool if you need it.

you need the vb runtimes installed for api_logger.exe, as well as the dll and ocx included
run the bat file to install them. if you are just using hte dll, it has no other requirements
and should be(?) fairly stable.

if you do use the appinit_dll setting, you will be logging every single process on the system
if it uses any of these api from the first boot of the machine.

