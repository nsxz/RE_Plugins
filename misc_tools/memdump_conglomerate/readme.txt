
reads a folder full of memory dumps and puts them 
all into a single dll husk so they will disassemble at the proper offsets.

drag and drop the folder into the text box to load it.

all files must be in the format

[pid]_[hexbaseaddr].mem 

pid is ignored, so you will have to manually sort them first

output file will be saved to scanned folder as: conglomerated_dmp.dll

the dll will not be able to loaded by the windows loader
since the entrypoint and imagebase are 0. this is so the sections
can be directly set to the va of the section which ida is fine with.

otherwise i would have to scan for the lowest one, set imgbase = that -1000
then adjust all the others to be relative to it. can be done..if I need it
it will be.





