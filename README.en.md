# Mac PowerPoint Batch Color Replacer

**Mac PowerPoint 批量改色** · **0.1.0-beta.1 — Beta** · **MIT**

A focused PowerPoint VBA add-in for **Mac desktop PowerPoint**. Replace fixed RGB colors across an open presentation, a slide selection, or specific slide numbers, without saving and reopening the file.

This project started with a practical need: update recurring colors in shapes, text, tables and charts while editing a presentation. The initiator could not find a lightweight Mac tool matching that workflow. Other add-ins covered parts of the task but also included features they did not need. This describes the initiator's experience, not an exhaustive comparison of available products.

## Download

[Mac installation bundle](https://github.com/WenkoLee/mac-powerpoint-batch-color/releases/download/v0.1.0-beta.1/MacPowerPointBatchColor_0_1_0_beta_1_Mac.zip) · [Beta release and checksums](https://github.com/WenkoLee/mac-powerpoint-batch-color/releases/tag/v0.1.0-beta.1)

## Features

- Chinese modeless panel with old/new color previews.
- HEX input, object color sampling, theme color selection and nearby-color scanning.
- All slides, selected slides, or explicit ranges such as `2,4-6`.
- Accessible shape, text, group, table and chart colors.
- Plain-text clipboard handling and input validation.
- Dedicated undo for the most recent add-in operation.

The current UI and detailed user documentation are in Chinese. See the [Chinese README](README.md), [installation guide](docs/INSTALL.md), [user guide](docs/USER_GUIDE.md) and [compatibility notes](docs/COMPATIBILITY.md).

![Chinese color picker panel](docs/images/panel.png)

The acceptance harness passed 46 application-level checks. See the acceptance record for manual interaction checks and unverified combinations.

## Build offline

Python 3.10+ is required. The tested MIT-licensed pyOpenVBA source snapshot is bundled in `vendor/`.

```sh
python3 scripts/build.py
python3 -m unittest discover -s tests -p 'test_*.py' -v
python3 scripts/audit.py
```

The installable `.ppam` and checksums are written to `release/`. See [development instructions](docs/DEVELOPMENT.md).

This is a beta prerelease. Application-level results and unverified items are listed in the [acceptance record](docs/acceptance/ACCEPTANCE.md).

## Limits

Mac desktop PowerPoint only; Windows, web and mobile PowerPoint are not supported/tested. Existing theme-linked shape/text colors are skipped. Mac chart APIs may report theme colors as RGB; review chart candidates before replacing them. Image pixels, individual gradient stops, SmartArt internals and some chart properties require manual handling. Use the add-in's dedicated undo; PowerPoint's standard Command+Z may not fully reverse VBA changes.

## License

[MIT](LICENSE). Preserve the third-party notices listed in [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md). This independent project is not affiliated with or endorsed by Microsoft.
