#!/usr/bin/env python3
"""
Build standalone HTML reports by inlining all ./data/*.js script files.
Run from the secureshare-phase3 folder.
"""
import os
import re

BASE = r'c:\Users\m347\Documents\Power BI Desktop\SecureShare Report\secureshare-phase3'
DIST = os.path.join(BASE, 'dist')

REPORTS = [
    ('Operational Reporting', 'Operational Reporting.html'),
    ('SecureShare Deployment Stats', 'SecureShare Deployment Stats.html'),
    ('Leavers Report', 'Leavers Report.html'),
    ('Finance Performance', 'Finance Performance.html'),
    ('EFT Confirmation', 'EFT Confirmation.html'),
    ('VRDD Appropriations', 'VRDD Appropriations.html'),
    ('Mothly Engagement', 'Monthly Engagement.html'),
]

def inline_scripts(html_path):
    html_dir = os.path.dirname(html_path)
    base_dir = os.path.abspath(html_dir)
    with open(html_path, 'r', encoding='utf-8') as f:
        content = f.read()

    replaced = 0
    total_inlined_kb = 0

    def replace_script_tag(m):
        nonlocal replaced, total_inlined_kb
        src = m.group(1)
        if not (
            src.startswith('./data/')
            or src.startswith('data/')
            or src.startswith('../Operational Reporting/data/')
        ):
            return m.group(0)  # leave CDN scripts alone

        if src.startswith('../Operational Reporting/data/'):
            data_path = os.path.normpath(os.path.join(base_dir, src.replace('/', os.sep)))
        else:
            data_path = os.path.join(html_dir, src.replace('/', os.sep))
        if not os.path.exists(data_path):
            print(f'  WARNING: {src} not found, skipping')
            return m.group(0)
        size = os.path.getsize(data_path)
        total_inlined_kb += size / 1024
        print(f'  + Inlining {src} ({size/1024:.0f} KB)')
        with open(data_path, 'r', encoding='utf-8') as df:
            data_content = df.read()
        replaced += 1
        return f'<script>\n{data_content}\n</script>'

    new_content = re.sub(
        r'<script\s+src=["\']([^"\']+)["\'](?:\s*/>|>\s*</script>)',
        replace_script_tag,
        content
    )

    return new_content, replaced, total_inlined_kb


os.makedirs(DIST, exist_ok=True)

for folder, filename in REPORTS:
    html_path = os.path.join(BASE, folder, filename)
    if not os.path.exists(html_path):
        print(f'SKIP (not found): {folder}/{filename}')
        continue

    print(f'\nBuilding: {folder}/{filename}')
    new_html, count, kb = inline_scripts(html_path)

    dist_folder = os.path.join(DIST, folder)
    os.makedirs(dist_folder, exist_ok=True)

    out_path = os.path.join(dist_folder, filename)
    with open(out_path, 'w', encoding='utf-8') as f:
        f.write(new_html)

    out_size = os.path.getsize(out_path)
    print(f'  => {count} data script(s) inlined ({kb:.0f} KB of data)')
    print(f'  => Output: {out_path} ({out_size/1024:.0f} KB)')

    # Copy assets folder
    assets_src = os.path.join(BASE, folder, 'assets')
    if os.path.exists(assets_src):
        import shutil
        assets_dst = os.path.join(dist_folder, 'assets')
        if os.path.exists(assets_dst):
            shutil.rmtree(assets_dst)
        shutil.copytree(assets_src, assets_dst)
        print(f'  => assets folder copied')

# Copy landing page
lp_src = os.path.join(BASE, 'Landing page.html')
if os.path.exists(lp_src):
    import shutil
    shutil.copy2(lp_src, os.path.join(DIST, 'Landing page.html'))
    root_assets = os.path.join(BASE, 'assets')
    if os.path.exists(root_assets):
        dst = os.path.join(DIST, 'assets')
        if os.path.exists(dst):
            shutil.rmtree(dst)
        shutil.copytree(root_assets, dst)
    print(f'\nLanding page copied to dist/')

print(f'\nDone. Standalone reports in: {DIST}')
print('Copy the dist/ folder contents to OneDrive.')
