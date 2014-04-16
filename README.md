fsrm-quotas-backup
==================

Quotas backup script for MS Windows File Server Resource Management

Usage:
.\quotas_backup.vbs - launch on old server (from which quotas are going to migrate)

.\quotas_create.bat - launch on new server (to which quotas are goinf to migrate)


Run quotas_backup.vbs on server from which you are willing to move quotas.
It will generate 3 files:
1) quota_templates.xml - contains quotas templates exported by "dirquota template export";
2) all_quotas.txt - contains quotas list generated by "dirquota q l";
3) quotas_create.bat - batch file with command to create quotas on a new server.

After that you should copy (or move) target folders to a new server and then
you need to edit path to target folders in file quotas_create.bat. Now you
should launch quotas_create.bat on new server and patiently wait for it to do
it's magic.