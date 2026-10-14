# One-Click Color for Mac PowerPoint

**一键改色** · **0.1.2** · **MIT**

A focused PowerPoint VBA add-in for **Mac desktop PowerPoint**. Replace RGB or theme colors across an open presentation, a slide selection, or specific slide numbers, without saving and reopening the file.

This project started with a practical need: update recurring colors in shapes, text, tables and charts while editing a presentation. The initiator could not find a lightweight Mac tool matching that workflow. Other add-ins covered parts of the task but also included features they did not need. This describes the initiator's experience, not an exhaustive comparison of available products.

## Download

[Mac installation bundle](https://github.com/WenkoLee/mac-powerpoint-batch-color/releases/download/v0.1.2/OneClickColor_MacPowerPoint_0_1_2_Mac.zip) · [Release and checksums](https://github.com/WenkoLee/mac-powerpoint-batch-color/releases/tag/v0.1.2)

## Features

- Chinese modeless panel with old/new color previews.
- HEX or theme-slot source, object color sampling (including selected group children), and theme targets.
- Direct exact replacement without scanning; optional nearby RGB candidate scanning.
- Theme-to-theme migration retains tint; fixed RGB targets clear tint.
- Defaults to all slides when the panel opens; selected slides and explicit ranges remain available.
- Accessible shape, text, group, table and chart colors.
- Plain-text clipboard handling and input validation.
- Dedicated undo for the most recent add-in operation.

The current UI and detailed user documentation are in Chinese. See the [Chinese README](README.md), [installation guide](docs/INSTALL.md), [user guide](docs/USER_GUIDE.md) and [compatibility notes](docs/COMPATIBILITY.md).

![Chinese color picker panel](docs/images/panel.png)

The 0.1.0-beta.2 color engine passed 99 application-level checks. Version 0.1.2 changes only the initial page scope to all slides; its UI verification is recorded separately. See the acceptance record for manual interaction checks and unverified combinations.

## Build offline

Python 3.10+ is required. The tested MIT-licensed pyOpenVBA source snapshot is bundled in `vendor/`.

```sh
python3 scripts/build.py
python3 -m unittest discover -s tests -p 'test_*.py' -v
python3 scripts/audit.py
```

The installable `.ppam` and checksums are written to `release/`. See [development instructions](docs/DEVELOPMENT.md).

See the compatibility notes before using the add-in. Application-level results and unverified items are listed in the [acceptance record](docs/acceptance/ACCEPTANCE.md).

## Limits

Mac desktop PowerPoint only; Windows, web and mobile PowerPoint are not supported/tested. HEX sources match current RGB values, including theme-linked objects; theme codes match the source slot, including tinted variants. Mac chart APIs may report theme colors as RGB, so theme-source chart matching falls back to the slot’s base RGB and may also match fixed chart colors with the same value. Image pixels, individual gradient stops, SmartArt internals and some chart properties require manual handling. Use the add-in's dedicated undo; PowerPoint's standard Command+Z may not fully reverse VBA changes.

## License

[MIT](LICENSE). Preserve the third-party notices listed in [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md). This independent project is not affiliated with or endorsed by Microsoft.
