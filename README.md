# Zepto Ransomware Document Dump
This is a dump of the sample shared by John Troony i.e. the .docm file sent via email to victims. It is a downloader for the actual ransomware

https://github.com/JohnTroony/Zepto-sample
 
The VT submission can be found here: 
https://www.virustotal.com/en/file/8aabb3e77b4d3a75a6cc1c7a5ef7c2d9fc7231ef0d890a2c14c8adf13d3ccca4/analysis/

The dump includes the following : 
- Streams
- Macros
- Unknown Sectors (sectors that are not represented in the FAT)
- FAT values that point to sectors that do not exist.
- FAT structure
- Expected file size based on the number of FAT sectors and sector size

