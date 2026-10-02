# SUK Tabs

SUK Tabs is a standalone web part. It creates its own tab panels; SharePoint does
not expose a supported way for it to take existing web-part instances from the
page canvas and move them into tabs. Configure the content for each tab in the
web part's **Tab configuration (JSON)** property.

## Panel types

- `textLinks`: plain text plus optional links.
- `documentLibrary`: latest files from a document library, optionally filtered
  by an exact value in a SharePoint field.
- `sukComponent`: one of the built-in SUK React components, configured through
  that panel's `component` and `settings` values.

Set the top-level **Panel heading** and **See-all link text** properties for the
header. A tab can optionally provide `seeAllUrl` (and override its link label
with `seeAllText`).

Example with all three panel types:

```json
[
  {
    "id": "recent",
    "title": "Terkini",
    "type": "documentLibrary",
    "seeAllUrl": "/sites/SmartOffice/Shared%20Documents",
    "documentLibrary": {
      "libraryId": "00000000-0000-0000-0000-000000000000",
      "itemLimit": 5,
      "emptyMessage": "Tiada dokumen terkini."
    }
  },
  {
    "id": "guides",
    "title": "Garis Panduan",
    "type": "textLinks",
    "content": "Useful guidance and reference documents.",
    "links": [
      {
        "text": "Project management guide",
        "description": "Open the current guide",
        "url": "/sites/SmartOffice/Shared%20Documents/guide.pdf"
      }
    ]
  },
  {
    "id": "quick-links",
    "title": "Pautan",
    "type": "sukComponent",
    "component": "quickLinks",
    "settings": {
      "sectionTitle": "Quick access",
      "items": [
        {
          "text": "Staff portal",
          "linkUrl": "/sites/SmartOffice/SitePages/Home.aspx",
          "iconUrl": ""
        }
      ]
    }
  }
]
```

For `documentLibrary`, `libraryId` is the document library's GUID. `filterField`
is the field's internal name and `filterValue` is an exact match; the configured
item limit is capped at 100.

## Built-in SUK components

Set `component` to one of `announcements`, `bannerSlider`, `birthdaySlider`,
`calendar`, `customLayout`, `gallery`, `quickLinks`, or `welcomeUser`. The
`settings` object uses the same property names as the corresponding component.
Data-backed components accept their source configuration there:

- Announcements: `listId`, `titleField`, `descriptionField`, `dateField`, and
  optional `categoryField`.
- Calendar: SharePoint list mode with `listId` and event field mappings.
- Gallery: `libraryId`, optional `rootFolderUrl`/`folderUrl`, and display
  settings.
- Birthday slider: `sourceList` (list title), optional `mappings`,
  `birthdayRange`, and `maxItems`. A `people` array can be supplied instead of
  `sourceList`.
- Welcome user: `sourceList`, optional `mappings`, and `selections` containing
  `{ "itemId": 1, "days": 14, "startDate": "2026-10-02" }` records. A `people`
  array can be supplied instead of `sourceList`.
- Banner slider: configure a `slides` array.
- Quick links: configure an `items` array.
- Custom layout: configure `columnCount`, `column1`, and optional additional
  columns.

Each panel loads only while its tab is active. This web part does not move or
reuse an existing SharePoint page web-part instance; it renders the selected
SUK component from the explicit settings stored in its own tab configuration.
