# SUK Activity Display

The Activity Display shows a selected SharePoint list as a responsive monthly
agenda. Choose a list, map its fields, and the web part loads that month's
activities with previous/next month navigation.

## Field mapping

- **Activity title** and **Start date and time** are required.
- **End date and time**, **Description**, **Location**, **Category**,
  **Activity link**, and **Organizer/contact** are optional.
- Field dropdowns are loaded from the selected list. Common title and start-date
  fields are preselected when available.
- URL fields can open an activity from its row. List view and New form URLs are
  derived from the list by default and can be overridden in Display settings.
- The event row shows the start date and time range, location, description,
  organizer, and category when those fields are mapped.

The list is queried for the visible month and results are ordered by the mapped
start date. Users must have read access to the selected list. New activity
creation follows the user's normal SharePoint list permissions.
