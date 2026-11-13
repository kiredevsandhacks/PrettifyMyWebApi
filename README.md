All of your rows can be quickly viewed and edited! Even when fields are not available on any form. Quickly view and edit Power Automate Flow definitions.

How to use:
- Click the extension button when viewing a form to open the current row in the Dataverse Web Api and prettify the output!
- When viewing a 'solution-aware' flow in make.powerautomate.com, use the extension to open it in the Dataverse Web Api. Click the Edit Flow button in the top right to edit your flow definition! This is very convenient for renaming actions/variables or as a last measure for working around bugs, when the default flow UI just doesn't want to cooperate (Power Automate developers will relate to this).
- When already viewing a row or query in the Dataverse Web Api, just use the extension to prettify!

You no longer need to open your solutions to view the values for your Option Sets/Choices. Simply open the row and prettify the output.

Browse to related rows for quick insights using the 'Preview' button on lookup columns.

Update columns of a row by using the 'Edit this row' feature. Update text columns, Option Sets, numbers, lookups, etc. Very convenient for modifying data that is not available on a form. You also have the option to impersonate other users, skip custom plugins and skip Power Automate Flows when updating a row.

Create new rows by using the 'Create new row' button.

Browse related rows by using the 'View relationships' button. When viewing related rows, use the 'Disassociate this row' button to disassociate the row from the parent/related row. 

Quickly copy the contents of a column to your clipboard by hovering and clicking the copy button.

You can prettify single rows, but also the results of your OData and FetchXml queries.


Firefox:
- Run ./firefox/build.sh to build the Firefox version. It writes dist/firefox and a zip in dist.
- Temporary install: open about:debugging, choose This Firefox, Load Temporary Add-on, and select dist/firefox/manifest.json. Firefox removes it again when you close the browser.
- Permanent install: sign the zip as an unlisted add-on on addons.mozilla.org, then open the signed .xpi in Firefox. Firefox Developer Edition and Nightly can also load the unsigned zip when xpinstall.signatures.required is set to false in about:config.
- Firefox needs access to your Dataverse site. When it asks, allow it, or enable it in about:addons under Permissions.
- Firefox shows application/json in its own JSON viewer and adds a strict Content-Security-Policy to that page. The Firefox build serves Dataverse Web Api pages (dynamics.com) as plain text, which gives the same page Chrome shows, and the extension works from there.
- Containers: the extension opens new tabs in the same Firefox container as the tab where you click the button, so you stay logged in.
