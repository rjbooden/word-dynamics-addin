# Release
This is the latest release build of the Word Dynamics Add-In

Please create an app registration and update the appsettings.json and manifest.xml accordingly. And create your own setting.json file. For details see **App registration** and **Configuration** the top level [readme.md](../readme.md) It is also posible to use the **CreateAppRegistration.ps1** PowerShell commands for creating the app registration *(uses Azure CLI)*. This PowerShell script can even do the configuration for you.

Optionally you can add additional locales files if needed, currently Dutch (*nl-nl.json*) and English (*en-us.json*) are supplied. These files need to be placed in the **wwwroot\locales** folder. The translation will be loaded for the language used for Office (**File > Options > Language**). If not found it will use the default language (*en-us*), if translations are missing the default (*hard coded*) translation will also be used.