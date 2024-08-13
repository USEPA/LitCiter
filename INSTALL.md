# How to install LitCiter

These instructions contain steps on how sideload the LitCiter addon. Based on [these](https://docs.microsoft.com/en-us/office/dev/add-ins/testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins) instructions.

If you do not have EPA network access and access to the O drive, you will need to put the add-in manifest on a network drive. Follow the instructions on the botton of this page for instructions.

1. In Word, go to `File > Options > Trust Center > Trust Center Settings > Trusted Add-in Catalogs`.
2. Add your network share path to the `Catalog Url` box and click `Add catalog`. If you have O drive access, this path is `\\aa\ord\ORD\DATA\Public\HERO\LitCiter`.
4. Check the `Show in Menu` box, and click `OK` to save your changes.
5. Close and reopen your Office Applications.
6. In Word, go to the `Insert` tab at the top and select `My Add-ins`. Navigate to the `Shared Folder` tab of the new dialog box.
7. Select `LitCiter` and click `Add`. LitCiter should now appear in Word's `Home` tab.

To run the addon, `Track Changes` needs to be disabled. Click the `Review` tab at the top of Word, then click the `Track Changes` button so it is not highlighted. You can re-enable it after using the addon.

To remove the new version of LitCiter, go back to the `Trusted Add-in Catalogs` box, check `Next time Office starts, clear all previously-started web add-ins cache`, and click `OK`. Close and reopen your Office applications. See [this](https://docs.microsoft.com/en-us/office/dev/add-ins/testing/clear-cache#clear-the-office-cache-on-windows) page for more details.

To remove the old version of LitCiter, open `Control Panel`, and click `Programs > Uninstall a Program` or `Programs and Features`. Find `LitCiter` in the list of applications and click `Uninstall`.

## Creating your own network drive
If you do not have access to the O drive, you can follow these instructions to add a shared folder locally.
1. Find a folder that you want to use to host LitCiter. Move save [manifest.xml](https://hero.epa.gov/static/litciter/manifest.xml) to this folder.
2. Right click on the folder, select properties, and go to the Sharing tab.
3. Under the Sharing tab, click `Share`
4. Within the Network access dialog window, add yourself. You'll need at least Read/Write permission to the folder. Then click `Share`.
5. Make note of the shared folder path (it should start with `\\`) and close the dialog.
