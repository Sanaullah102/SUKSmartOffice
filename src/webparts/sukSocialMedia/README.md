# SUK Social Media

Create a responsive social-media directory with any number of platform profiles.
Add or remove accounts in **Accounts and embeds (JSON)** in the web-part
properties. Each account requires `platform`, `label`, and `url`; optional
properties are `handle`, `description`, `group`, `embedUrl`, and `enabled`.
Setting `enabled` to `false` hides an account without removing its configuration.

Example:

```json
[
  {
    "platform": "facebook",
    "label": "Facebook Page",
    "url": "https://www.facebook.com/example",
    "handle": "@example",
    "description": "News and community updates",
    "group": "Social",
    "embedUrl": "https://www.facebook.com/plugins/page.php?href=https%3A%2F%2Fwww.facebook.com%2Fexample&tabs=timeline&width=500&height=420&small_header=true&adapt_container_width=true"
  },
  {
    "platform": "linkedin",
    "label": "LinkedIn",
    "url": "https://www.linkedin.com/company/example",
    "group": "Professional"
  },
  {
    "platform": "pinterest",
    "label": "Pinterest",
    "url": "https://www.pinterest.com/example",
    "group": "Social"
  },
  {
    "platform": "custom",
    "label": "Official portal",
    "url": "https://example.com",
    "handle": "example.com",
    "group": "Other"
  }
]
```

Recognized platform marks and colors include Facebook, Instagram, LinkedIn,
Pinterest, X/Twitter, YouTube, TikTok, Threads, WhatsApp, Telegram, Reddit,
Snapchat, Spotify, SoundCloud, Twitch, Vimeo, Discord, and WeChat. Use
`"platform": "custom"` or any other platform name for an additional link.

Five display styles are available: **Icon strip**, **Social profile cards**,
**Featured accounts**, **Embedded feeds**, and **Compact list**. Configure
account groups as tabs, columns, labels, descriptions, handles, target tab
behavior, embed height, and shared theme/custom appearance.

Embeds are optional and use the `embedUrl` supplied in each account. For safety,
only HTTPS embeds from supported platform hosts are rendered in a sandboxed
iframe. Social platforms may require an official plugin/embed URL; regular
profile URLs are shown as links rather than loaded as feeds. Third-party
availability and login requirements remain controlled by each platform.
