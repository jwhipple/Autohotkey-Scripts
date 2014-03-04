Created by Joe Whipple for KEMBA Financial Credit Union on FEB 2009
Licensed under the Creative Commons License
Free to use/modify as long as this notice remains intact.

Hope this helps you like it helped me.



All the ahk files are AutoHotkey files. Tested and working with v1.1.11.01
www.autohotkey.com

The MainLogon.ahk is the only file needed to be compiled with Autohotkey. All other ahk's are either designed to be run stand-alone, or included 
in MainLogon.ahk. Compile by having AutoHotkey installed and right clicking the script.

Save all .ahk sources in a non-user accessible location for security when deploying.

netlogon.bat should be the file executed in your group policy logon script location/file setting.

READ THE SOURCE, UNDERSTAND WHAT IT DOES. I am not responsible for nuking your domain, etc. I use this script in production, but our enviroment is 
very different than yours. The outlook setup doesnt always work with things like exchange hosted somewhere else other than your domain.

The functions in JRWTools.ahk are very powerful and can be used for other AutoHotkey scripts. Look at whats there.

Use MakeAHash.ahk to make your admin password/username hashes, but change the salt to something else.

Recompile Search64BitReg.exe if you must (I included the source), but make SURE its compiled in 64 bit mode.

The sources do not have to be accessible to the compiled EXE. STORE THEM SEPRATE FROM USERSPACE.

And as always, I am happy to help you, but I expect some effort on your part. I will not admin for you or teach you AutoHotkey.

-Joe
