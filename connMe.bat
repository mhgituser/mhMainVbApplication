@echo off

NET USE * /DELETE /YES
net use m: \\192.168.1.12\dummy /USER:WORKGROUP\admin password /PERSISTENT:NO

