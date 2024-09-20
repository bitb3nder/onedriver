# Onedriver

Post-exploitation script to access OneDrive with an access/refresh token. Supports basic file and folder enumeration such as `cd`, `ls`, `upload`, `download`, and `remove`. 

### LNK Backdooring

If desktop sync is enabled on Onedrive, arbitrary code execution could be acheived by backdooring an .lnk that is sync'd to the targets desktop. The `backdoor` module will download the target .lnk, modify it with a backdoored command, and upload the backdoored .lnk to the system. 
