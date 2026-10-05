# SUK Document Library

Browse a selected SharePoint document library with optional folder navigation,
breadcrumbs, and search scoped to that library. Choose to show documents,
folders, or both. Search is submitted with the button or Enter and matches the
file name or title; folder navigation remains available when search is cleared.
Search input state is captured before React schedules its state update to avoid
synthetic-event lifetime issues.

If a **Group results into tabs by field** mapping is configured, the web part
creates an All tab and one tab for each group value in the loaded results.
Configure the page size and move between pages using SharePoint's continuation
links. Also configure the sort field and direction, title and description
mappings, metadata visibility, folder recursion, search labels, and empty state.

Folder icons can use the standard yellow Windows-style folder or stable,
different colors per folder. Document icons follow familiar file-type colors
and labels for Word, Excel, PowerPoint, PDF, images, archives, and text files.

Five responsive styles are available: **List rows**, **Two-column grid**,
**Metadata table**, **Document cards**, and **Compact links**. Appearance can
follow the site theme or use the shared custom color, font, and border options.
