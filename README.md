# MDQ contributor homepage

Release copy of the approved Microsoft Discovery contributor homepage. The banner and page use 100% of the available viewport width. The banner remains 398px high. Narrow screens stack the content panels.

## Status

Source is stored in the private `IvanPedroza/mdq-contributor-homepage` repository; web hosting is not deployed. Contains internal program guidance, professional contact addresses, internal links, the program process diagram, and the publication guide. Keep this complete package in a private repository and an approved authenticated host. The existing `IvanPedroza/mdq-showcase-embed` repository and its Pages site are public; adding these files there publishes their contents.

## Preview

Run `python3 -m http.server 4175 --bind 127.0.0.1` from this folder and open `http://127.0.0.1:4175/`.

Local annotation controls, preview notices, screenshots, and review notes are excluded from this release copy. The original local design remains unchanged apart from its approved layout edits.

## What works

Submission Form and tracker links, contact email links, asset guidance tabs, publication-stage tabs, searchable troubleshooting, resource dialogs, and the animated linked-story banner.

## Connections still needed

- The custom status/calendar panel is not connected to Microsoft Lists. The working tracker link opens the actual list. Use a native SharePoint List web part for live records, or implement an approved authenticated Microsoft integration separately.
- The four supplied document links remain accessible in relevant dialogs. Their titles and owners cannot be derived from the opaque URL IDs. The tag-taxonomy file, article-template file, and approved asset-template assignment still need confirmation.
- The publication-guide link opens the bundled known guide (v1.0, September 29, 2026). Replace it with a confirmed SharePoint guide URL when available.

## SharePoint

See `SHAREPOINT-SETUP.md`. This folder is a standalone website, not an uploadable SharePoint page or SPFx package. GitHub source storage alone does not create an authenticated hosting endpoint. A private GitHub repository also does not automatically make GitHub Pages private.
