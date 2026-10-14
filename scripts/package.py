#!/usr/bin/env python3
"""Create local review archives without contacting GitHub."""
from pathlib import Path
import hashlib
import json
import re
from zipfile import ZipFile, ZipInfo, ZIP_DEFLATED
from audit import audit, public_files

ROOT = Path(__file__).resolve().parents[1]

def archive(path, files, prefix, extra=None):
    with ZipFile(path, 'w') as out:
        for file in files:
            info = ZipInfo(prefix + '/' + file.relative_to(ROOT).as_posix(), (2026, 1, 1, 0, 0, 0))
            info.compress_type = ZIP_DEFLATED
            info.external_attr = 0o644 << 16
            out.writestr(info, file.read_bytes())
        for name, value in (extra or {}).items():
            out.writestr(prefix + '/' + name, value)

def main():
    result = audit()
    if result['issues']:
        raise SystemExit(json.dumps(result, ensure_ascii=False, indent=2))
    version = (ROOT / 'VERSION').read_text().strip()
    stem = 'MacPowerPointBatchColor_' + re.sub(r'[^0-9A-Za-z]', '_', version)
    target = ROOT / 'dist'
    target.mkdir(exist_ok=True)
    files = public_files()
    archive(target / (stem + '_Source.zip'), files, 'mac-powerpoint-batch-color')
    user_files = [p for p in files if p.relative_to(ROOT).parts[0] in {'release', 'examples', 'licenses'}
                  or p.name in {'LICENSE', 'THIRD_PARTY_NOTICES.md'}
                  or p.relative_to(ROOT).as_posix().startswith(('docs/images/', 'docs/acceptance/'))
                  or p.relative_to(ROOT).as_posix() in {'docs/INSTALL.md', 'docs/USER_GUIDE.md', 'docs/FAQ.md', 'docs/COMPATIBILITY.md'}]
    intro = '''# Mac PowerPoint 批量改色 · 测试版

Mac PowerPoint Batch Color Replacer · VERSION

1. 先阅读 [安装说明](docs/INSTALL.md)。
2. 安装包在 `release/`，请先解压并放到固定位置后加载。
3. 打开 [自制示例](examples/BatchColorDemo.pptx) 试用。
4. 详细操作见 [用户指南](docs/USER_GUIDE.md)。
5. 回退颜色请使用插件的“撤销上次改色”。

[兼容性与限制](docs/COMPATIBILITY.md) · [验收记录](docs/acceptance/ACCEPTANCE.md) · [MIT 许可证](LICENSE)

本包没有客户资料，不自动保存文件，不上传演示文稿。当前为未签名测试版；请阅读安装说明中的宏保护设置。独立源码另见 Source 压缩包。
'''.replace('VERSION', version)
    archive(target / (stem + '_Mac.zip'), user_files, stem, {'README.md': intro})
    sums = [hashlib.sha256(p.read_bytes()).hexdigest() + '  ' + p.name for p in sorted(target.glob(stem + '*.zip'))]
    sums += [hashlib.sha256(p.read_bytes()).hexdigest() + '  ' + p.name for p in sorted((ROOT / 'release').glob(stem + '.ppam'))]
    (target / 'SHA256SUMS.txt').write_text('\n'.join(sums) + '\n')
    for p in sorted(target.glob(stem + '*.zip')):
        print(p.name, p.stat().st_size)

if __name__ == '__main__':
    main()
