Simple Windows remote control API service (python 3 + flask)
************************************************************

I like when EVA ICS (https://www.eva-ics.com) automatically turns on my
computer when I come to office and puts it to sleep when I go. I was too lazy
to program this correctly in past and simply used telnet server on machine to
let remote script login and execute sleep command. But I've upgraded my
computer to Windows 10 and found telnet server is no longer available.

Then, as default MS Windows RPC is very complex, I've decided to write a simple
python service to provide control API. The problem is it's not well documented
and need you to know a couple of tricks.

This is an example of 100% working python + flask service for Windows.

So, let's go, we have 2 files:

wincontrol.py
=============

If started with "app" param, launches flask application, otherwise starts to
win32serviceutil params handler.

Params:

* **app** - launch web application

**install**  - install a service (SimpleWinControlAPI / can be found as Windows
  Remote Control Simple API in service list)

**remove** - remove a service

(run *wincontrol.py ---help* for all available) 

wincontrol.yml
==============

Service configuration. Set here *access-key*, *listen host:port*, hosts acl and
list of available commands (note that they'll be executed by *SYSTEM*).

In real life, make sure the file has correct permissions (not readable by
users).

Installing required libraries
=============================

Execute:

    pip install pywin32 flask pyyaml netaddr

API
===

Has only a couple of functions:

* **GET /state** - returns *{ "ok": true }*

* **POST /command/\<cmd>** - executes command, returns 404 if command is not
  defined, *ok:true* if command is executed correctly or *ok:false* if exit
  code was not zero.

Access key should be set in **X-Auth-Key** request header. Remote IP should
match **hosts-allow** acl if defined.

E.g. by default, request

    curl -X POST -H x-auth-key:123 ip-of-windows-machine/command/sleep

will run

    rundll32.exe powrprof.dll,SetSuspendState 0,1,0

and put the computer to sleep. Bingo! )

Dealning with 1053 service error
================================

* Make sure python is in *PATH* for all users.

* Make sure *pywintypesXX.dll* file is in PATH. If no - copy it to the python
  root folder (e.g. if you have python 3.7 installed in *c:\python37* - copy
  *pywintypes37.dll* from *c:\python37\Lib\site-packages\pywin32_system3* to
  *c:\python37*

* If the service is started, but flask app failed with an exception, it will be
  written to *wincontrol-error.log* (in the folder where python script is
  located)

Nuts and bolts
==============

The main problem I found - it's never noted that when python script is started
as service, it has no *sys.stdout* and *sys.stderr* set. So when any library
want to exec *sys.stdout.write* or *sys.stdout.flush()*, it will immediately
cause an exception, because *sys.stdout* is actually set to *None*.

So I've put a simple dummy stdout/stderr to prevent this.

