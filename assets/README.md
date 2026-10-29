The Security Bank wordmark and icon are copied from the bank's official support website, https://help.securitybank.com/.

- Wordmark: https://static.helpjuice.com/helpjuice_production/uploads/template/brand_image/6042/8952/logo.png
- Icon: https://static.helpjuice.com/helpjuice_production/uploads/upload/image/6042/454425/favicon.png
- Blue `#0168B3` and green `#97C93D` are used in the official support site's stylesheet: https://static.helpjuice.com/helpjuice_production/uploads/account/6042/main_css_6042-8224da8f44639f35e9bb56da69b3a9ecb596341a43ff409d116a64f2e97516b8.css

Assets were retrieved on October 29, 2026. The logo and bank name belong to Security Bank Corporation. They are bundled locally so the application does not depend on third-party image requests while running.

The current seasonal presentation uses a dark purple background and orange Halloween accents. `halloween-sidebar.svg` is an original decorative illustration; it does not alter the bank's logo.

The ghost gently sways, floats and blinks using SVG/CSS animations. Its shadow follows the motion. All three animations stop when the browser requests reduced motion. No JavaScript, remote animation service or extra dependency is needed.

The pumpkin and ghost scene sits near the bottom of the main page's empty left margin, using the existing `halloween-sidebar.svg` artwork. Its size follows the free space outside the centered 1280-pixel content container, up to 420 pixels wide. On narrow screens, the decorative scene disappears rather than overlapping forms or sidebar instructions. It stays visible in the main background when the sidebar is collapsed.

`halloween-webs.svg` adds soft corner cobwebs and little silk highlights to the page background. `halloween-spiderman.svg` is a locally drawn, transparent vector decoration featuring a cute upside-down Spider-Man in the shared header. CSS gently sways him from the web thread and disables that animation when reduced motion is requested. Both decorations are bundled with the app, need no external requests, and cannot intercept clicks. The header rearranges for narrow screens so the artwork, bank logo and workspace name remain readable.
