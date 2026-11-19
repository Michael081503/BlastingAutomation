The Security Bank wordmark and icon are copied from the bank's official support website, https://help.securitybank.com/.

- Wordmark: https://static.helpjuice.com/helpjuice_production/uploads/template/brand_image/6042/8952/logo.png
- Icon: https://static.helpjuice.com/helpjuice_production/uploads/upload/image/6042/454425/favicon.png
- Blue `#0168B3` and green `#97C93D` are used in the official support site's stylesheet: https://static.helpjuice.com/helpjuice_production/uploads/account/6042/main_css_6042-8224da8f44639f35e9bb56da69b3a9ecb596341a43ff409d116a64f2e97516b8.css

Assets were retrieved on October 29, 2026. The logo and bank name belong to Security Bank Corporation. They are bundled locally so the application does not depend on third-party image requests while running.

The current seasonal presentation uses a deep evergreen background, cranberry red buttons, warm gold accents and falling snow. Streamlit's native theme in `.streamlit/config.toml` uses the same palette as `security-bank.css`.

`christmas-spiderman.svg` restores the hanging Spider-Man in the shared header, wearing a red Santa hat with white trim and a pom-pom. He gently sways from his web; the hat stays attached and fits inside the illustration at every screen size. `christmas-garland.svg` adds evergreen garlands and red/gold baubles across the header. `christmas-winter.svg` places a snowy forest and snowman in the main page's free left margin on wide screens. The scene shrinks with the available margin and disappears on narrow screens. These are local vector illustrations.

Two repeating CSS snow layers move at different speeds behind the workspace. Snow uses no JavaScript, remote requests or extra dependencies, and its layers cannot capture clicks or file drops. When the browser requests reduced motion, the snow becomes a static background and Spider-Man stops swaying.

`christmas-snowflake.svg` supplies the gold snowflake icon on **Clear all files**. The header rearranges on small screens to keep the logo, seasonal badge and controls readable. The Christmas tree and earlier Halloween illustrations remain in the repository as unused assets.
