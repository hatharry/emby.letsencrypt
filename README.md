    Please make sure port 80 is forwarded to your Emby Server.
    
    Change "mailto:somebody@example.org" in the Script to your email address to receive notifications from Let's Encrypt.
    
    Set "Advanced - External domain" address in Emby eg. "www.myembyserver.com" .
    
    If "Advanced - Custom certificate path" is not in use then a cert is created in the Emby Server SSL directory using the filename "{yourwanaddress}.pfx". 
    It is required to add this cert manually to the "Advanced - Custom certificate path" setting in the Emby control panel.
    
    If "Advanced - Custom certificate path" is in use then the script will overwrite the old cert using the same filename.
