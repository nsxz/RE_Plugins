
takes a memory dump and embeds it into a dummy dll husk
so that you can disassemble it at the target base address
without having to manually reset it everytime.

this will be part of sysanalyzer soon

drag and drop the file into the file text box to load it.

if it is in the format

[pid]_[hexbaseaddr].mem 

then the base address field will auto populate.
output file will be saved to input file parent 
directory as

[hexbase]_dmp.dll



