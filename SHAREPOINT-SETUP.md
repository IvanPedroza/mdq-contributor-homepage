# SharePoint setup

The homepage is a static HTML/CSS/JavaScript site. Copying its files into a SharePoint library does not install a modern SharePoint page or connect the tracker.

Repository settings checked October 5, 2026: `IvanPedroza/mdq-showcase-embed` is **public**, and GitHub Pages publishes `main:/` publicly. The complete homepage contains internal guidance, contacts, and a bundled Word guide. The complete homepage source is stored separately in the private `IvanPedroza/mdq-contributor-homepage` repository. A private source repository alone does not establish authenticated web hosting. Do not paste the GitHub repository URL into the Embed web part; it needs a deployed website URL.

## Native SharePoint page: internal content and live tracker

1. Open the program SharePoint site, create or edit a modern page, and add a **one-column section**. The local homepage now fills its available width, but it cannot change SharePoint's page layout. Full-width sections are unavailable on team sites; use the widest supported section. See [Microsoft's section and column guidance](https://support.microsoft.com/en-us/sharepoint/pages-in-sharepoint/add-sections-and-columns-on-a-sharepoint-modern-page).
2. Add an **Embed** web part for the existing public banner and paste:

   ```html
   <iframe src="https://ivanpedroza.github.io/mdq-showcase-embed/?motion=on" width="100%" height="398" title="MDQ Blog Program showcase" style="border:0" loading="lazy"></iframe>
   ```

   Keep **Resize to fit the page** off for the intended 398 px banner height. Remove `?motion=on` to respect the visitor's reduced-motion preference. If embedding is blocked, ask the site administrator to allow `ivanpedroza.github.io`. See [Microsoft's Embed web part guidance](https://support.microsoft.com/en-us/office/add-content-to-your-page-using-the-embed-web-part-721f3b2f-437f-45ef-ac4e-df29dba74de8).
3. Use **Text**, **Quick Links**, and **People** web parts for Start here, the proposal Form, workspace link, contacts, asset guidance, and publication steps. Keep internal documents in the site's document library; link the verified files or add a [Document Library web part](https://support.microsoft.com/en-au/sharepoint/web-parts-and-apps-in-sharepoint/use-the-document-library-web-part).
4. Add a **List** web part and select **Blog Publication Tracker** from this site. Select a saved view with the available title, author, stage, target date, next action, and workspace-link columns. Ensure contributors have access to the List and linked folders. This displays real records with SharePoint permissions; it replaces the prototype's disconnected table. See [Microsoft's List web part instructions](https://support.microsoft.com/en-us/sharepoint/lists/web-parts-and-apps-in-sharepoint/use-the-list-web-part).
5. Create or use a calendar view of the tracker based on its publication-date column and add a Quick Link to that view. Confirm the actual field names and resource files before labeling templates as approved.
6. Preview, test links and List access with a contributor account, then publish or republish. Native web parts need their own layout; the prototype's custom tabs and dialogs do not transfer automatically.

## Hosted homepage: preserve the custom interface

1. Choose an approved HTTPS host. Use authenticated hosting for the complete internal version, or prepare and approve the content for public release before publishing to public GitHub Pages. Preserve the existing banner URL when adding a separate homepage route.
2. Deploy the homepage and its relative dependencies together. Verify the deployed page directly, then add a SharePoint **Embed** web part using its actual URL. The host must permit iframe embedding; SharePoint may also require the administrator to allow its domain.
3. Set the iframe width to **100%**. Start with **1100 px** height and test scrolling and dialogs on desktop and mobile; **398 px is only the banner height**. The current homepage has no parent-page auto-height integration, so do not assume it expands to fit all content.
4. Add a native **List** web part outside the iframe for live status, or implement a separate approved authenticated integration. Embedding the page does not grant it SharePoint access. Its current table and calendar remain disconnected.
5. Remove local review controls and preview-only statements from the release copy. Confirm the guide link, taxonomy workbook, article template, and asset templates; those mappings are still unfinished. Test the deployed page inside SharePoint before republishing.
