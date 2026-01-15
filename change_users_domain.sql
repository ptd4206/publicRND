set timing on
SET SERVEROUTPUT ON
SET DEFINE ON
set lines 120 pages 9999
SET NEWPAGE 0
SET SPACE 0
SET LINESIZE 200
SET PAGESIZE 0
SET ECHO OFF
SET FEEDBACK OFF
SET HEADING OFF
SET MARKUP HTML OFF SPOOL OFF
SET term off 
SET verify off
set echo off


var NEWDOM VARCHAR2(2);
var OLDDOM VARCHAR2(2);
var PACNAM VARCHAR2(16);
BEGIN
:NEWDOM := '&&1';
:OLDDOM := '&&2';
:PACNAM := '&&3';
END;
/

SET DEFINE OFF

BEGIN

	update cfg.cfg_repository set value=:NEWDOM where PATH_SEGMENT='imaginet\system\nodes\\:PACNAM\properties\current_domain';
	update medilink.secm_ext_pass set domain=:NEWDOM where domain=:OLDDOM;
	update medilink.secm_groups set domain=:NEWDOM where domain=:OLDDOM;
	update medilink.secm_sync_del_groups set domain=:NEWDOM where domain=:OLDDOM;
	update medilink.secm_sync_del_users set domain=:NEWDOM where domain=:OLDDOM;
	update medilink.secm_users set domain=:NEWDOM where domain=:OLDDOM;
    commit;     

END;
/

quit;
/

SPOOL OFF;
SET DEFINE ON
SET SERVEROUTPUT OFF
