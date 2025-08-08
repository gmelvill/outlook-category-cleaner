\# Outlook Category Cleaner



An Outlook add-in that scans for unused categories and allows you to delete them.



\## Setup



1\. \*\*Create a Private Repo\*\*  

&nbsp;  On GitHub, create a new repo `outlook-category-cleaner` and upload this structure.



2\. \*\*Enable GitHub Pages\*\*  

&nbsp;  - Go to `Settings` → `Pages`.

&nbsp;  - Set source to `main` branch, root (`/`).

&nbsp;  - Save, and you’ll get a URL like:  

&nbsp;    `https://gmelvill.github.io/outlook-category-cleaner/`



3\. \*\*Update manifest.xml\*\*  

&nbsp;  Make sure the `<SourceLocation>` and `<IconUrl>` in `manifest.xml` point to your GitHub Pages URL.



4\. \*\*Sideload into Outlook\*\*  

&nbsp;  - Open \*\*New Outlook\*\*.

&nbsp;  - `Home` → `Add-ins` → `My Add-ins` → \*\*Add from file\*\*.

&nbsp;  - Select `manifest.xml` (download it from your repo).

&nbsp;  - Launch the add-in from the ribbon.



\## Permissions

This add-in uses `ReadWriteMailbox` permission and Microsoft Graph API to list, check, and delete categories.





