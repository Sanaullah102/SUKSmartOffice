# SUK News

SUK News displays responsive SharePoint news with either **SharePoint News
Pages** (the Site Pages library) or a standard SharePoint list as its source.
Select the source and map the title, publication date, summary, image, category,
and optional target link in the property pane. Field options are discovered from
the selected source. For News Pages, items are restricted to promoted news pages
when the library exposes the `PromotedState` field.

Five distinct responsive display styles are available:

- **Image-led card grid**: consistent, bordered cards in a multi-column grid.
- **Horizontal story rows**: landscape thumbnails beside wider story summaries.
- **Editorial columns**: text-forward columns with serif headlines and restrained
  dividers.
- **Featured lead story**: a large, overlaid lead image beside smaller supporting
  stories.
- **Compact headlines**: dense single-line rows for narrow sections.

The Filters and sorting group supports promoted-page selection, included and
excluded keywords, any/all keyword matching, keyword search scope, included and
excluded categories, a publication-date window, future-date handling, sorting,
and skipping matching items before the display limit. Keyword and category lists
accept comma- or line-separated values. Category rules match the full category
value without regard to letter case.

The Image display group controls image fitting (crop, contain, stretch, original
size, or scale down) and alignment. Fit/alignment settings apply to each style,
including the feature-story layout. The web part also supports a configurable
title, item limit, see-all link, empty-state message, and shared appearance
settings. Image mappings accept URL and SharePoint image-field values; missing
or unavailable images use a neutral placeholder.
