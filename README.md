# Web-MS-Teams-UserScript

Script to expose several actions and meetings control status via a named API for easier interaction and automation via external js, Browser Development Tools or Remote Debugging Protocol.

This version is currently aimed at Teams v2, which is found at https://teams.microsoft.com.mcas.ms/v2/

## How to use it 

1. Download or clone the repo

2. First you need to install [Tampermonkey](https://www.tampermonkey.net/) and allow userscripts execution.

3. Go to src folder and open teams-userscript.js

4. Edit the following line:

``` js
    unsafeWindow.myTeamsAPIName = API;
```

change `myTeamsAPIName` for something of your liking. I suggest you to use this format

`[Company]TeamsAPI`

or 

`[YourName]TeamsAPI`

so the result is something like:

``` js
    unsafeWindow.GitHubTeamsAPI = API;
```

``` js
    unsafeWindow.JohnDoeTeamsAPI = API;
```

> [!NOTE]
> If you have multiple teams accounts, you must isolate each one on a separate profile (isolated profile) if your browser allows it, so you can define a different API name for each account. Usually each profile must have Tampermonkey installed

5. Once you have applied your changes you can load the script via Tampermonkey (_[How do I install and uninstall Tampermonkey?](https://www.tampermonkey.net/faq.php?locale=en&q=Q100)_)

6. Go to Teams and Log in. If already logged reload the window.

7. API is now available. I suggest you to use your browser dev tools to test it.

Try:

``` js
    JohnDoeTeamsAPI.general.chats.goToCalendar()
```

this should take you to the calendar
