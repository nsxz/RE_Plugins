
So apparently my OllySync (contained within the Olly_VBScript plug-in)
no longer operates with my current build of IDAVbscript which contains the
cmd socket handler to accept the commands from the remote olly instance over UDP.

Actually my IDA VBScript must be out of sync. I must've lost the working version
somewhere along the way because the commands required were not even implemented.
Any rate it was just crashing IDA..

soo this small proxy app comes in to play.

it will listen on UDP 3333 for commands from OllySync

as you single step or hit breakpoints in olly (after enabling the plugin sync feature)
olly will send a udp jmp_rva command to the ip you have configured. if this tool is running
then it forward the command on to IDA using IDASrvr .

Note you will have to have \IDASrvr\ActiveX_clientLib\IDAClientLib.dll registered on your 
system. (from a 32bit elevated cmd.exe window or just install IDACompare)

this is the cleanest way to fix up the mess.