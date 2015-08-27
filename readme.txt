//=================================================================================================================
//=						 _          _    _   _                           _               
//=						| |    ___ | | _(_) | |    __ _ _   _ _ __   ___| |__   ___ _ __ 
//=						| |   / _ \| |/ / | | |   / _` | | | | '_ \ / __| '_ \ / _ \ '__|
//=						| |__| (_) |   <| | | |__| (_| | |_| | | | | (__| | | |  __/ |   
//=						|_____\___/|_|\_\_| |_____\__,_|\__,_|_| |_|\___|_| |_|\___|_|   
//=										           _            _ 
//= 									 _ __ ___ | | ___  ___ / |
//=										| '_ ` _ \| |/ _ \/ _ \| |
//=   									| | | | | | |  __/ (_) | |
//=  									|_| |_| |_|_|\___|\___/|_|                                                       
//=
//=== Description =================================================================================================
//=
//=	http://hercules.ws/board/topic/1070-loki-launcher/
//=
//=== Notes  ======================================================================================================
//= 
//= DO NOT DIFF Restore Login Window
//= 
//=== Shortcuts ===================================================================================================
//=	
//=	ESC, ALT+F4 = exit
//=	ENTER		= login
//=	F1			= replay
//=	F5			= call setup.exe
//=	CTRL+1		= chk keep
//=	
//=== Changelogs ==================================================================================================
//=	
//=	v0.1	Initial commit
//=	v0.2	MD5 pass support
//=	v0.3	File req checks
//=			Fix MD5
//=	v0.4	Login window now movable
//=	v0.5	Added RGB (255, 0, 255) transparency
//=			Auto hide when login, auto show when called exe has exit, idea from Neomind
//=			Form can now be triggered to be dragable when click body of that form
//=			Added replay button, as suggested by Yommy
//=			INI file uses the same name as exe, as suggested by Yommy
//=			Thanks to Igniz and Cjei for buttons and skins (images are not used mwahaha)
//=	v0.6	Added show in tray when playing, as suggested by EvilPuncker
//=	v0.7	Window Focus Fix
//= v0.73	Fixed Replay. Thanks to Densetsu
//= v1.00	Autoshow is now configurable
//= 		File checks are now triggered by button events, instead of start, as suggested by Ai4rei
//=			Auto create ini file when there is none
//=			Loginwindow autopositions based on skin loaded
//=			+ more faggotries
//=	v1.10	Removed Auto aka runonbgtray, Plus fixes, Added Admin Manifest, Added configurer
//=	v1.11	Focus on usertextbox when there is no user else passtextbox, Fixes
//=	
//=================================================================================================================