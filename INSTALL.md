# How to install LitCiter

These instructions contain steps on how sideload the LitCiter addon. Based on [these](https://docs.microsoft.com/en-us/office/dev/add-ins/testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins) instructions.

You must have access to the O drive for these instructions to work. If you do not, please reach out. You must also be connected to the EPA network (e.g. by VPN).

1. In Word, go to `File > Options > Trust Center > Trust Center Settings > Trusted Add-in Catalogs`
2. Add `\\aa\ord\ORD\DATA\Public\HERO\LitCiter` to the `Catalog Url` box and click `Add catalog`
3. Check the `Show in Menu` box, and click `OK` to save your changes
4. Close and reopen your Office Applications
5. In Word, go to the `Insert` tab at the top and select `My Add-ins`. Navigate to the `Shared Folder` tab of the new dialog box.
6. Select `LitCiter` and click `Add`. LitCiter should now appear in Word's `Home` tab.

To run the addon, `Track Changes` needs to be disabled. Click the `Review` tab at the top of Word, then click the `Track Changes` button so it is not highlighted. You can re-enable it after using the addon.

To remove the new version of LitCiter, go back to the `Trusted Add-in Catalogs` box, check `Next time Office starts, clear all previously-started web add-ins cache`, and click `OK`. Close and reopen your Office applications. See [this](https://docs.microsoft.com/en-us/office/dev/add-ins/testing/clear-cache#clear-the-office-cache-on-windows) page for more details.

To remove the old version of LitCiter, open `Control Panel`, and click `Programs > Uninstall a Program` or `Programs and Features`. Find `LitCiter` in the list of applications and click `Uninstall`.
