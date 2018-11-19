    Please make sure port 80 is forwarded to your Emby Server.
    
    Set "Advanced - External domain" address in Emby eg. "www.myembyserver.com" .
    
    If "Advanced - Custom certificate path" is not in use then a cert is created in the Emby Server directory using the filename "{yourwanaddress}.pfx". 
    It is required to add this cert manually to the "Advanced - Custom certificate path" setting in the Emby control panel.
    
    If "Advanced - Custom certificate path" is in use then the script will overwrite the old cert using the same filename.
