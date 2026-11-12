// Sets the app icon and version info on the Windows exe without needing wine.
const fs = require('fs'); const path = require('path');
exports.default = async function (ctx) {
  if (ctx.electronPlatformName !== 'win32') return;
  const ResEdit = await import('resedit');
  const exePath = path.join(ctx.appOutDir, ctx.packager.appInfo.productFilename + '.exe');
  const exe = ResEdit.NtExecutable.from(fs.readFileSync(exePath), { ignoreCert: true });
  const res = ResEdit.NtExecutableResource.from(exe);
  const ico = ResEdit.Data.IconFile.from(fs.readFileSync(path.join(__dirname, 'icon.ico')));
  const groups = ResEdit.Resource.IconGroupEntry.fromEntries(res.entries);
  const id = groups.length ? groups[0].id : 1; const lang = groups.length ? groups[0].lang : 1033;
  ResEdit.Resource.IconGroupEntry.replaceIconsForResource(res.entries, id, lang, ico.icons.map(i => i.data));
  const vi = ResEdit.Resource.VersionInfo.fromEntries(res.entries)[0];
  if (vi) {
    const v = ctx.packager.appInfo.version.split('.').map(Number);
    vi.setFileVersion(v[0], v[1], v[2], 0, lang); vi.setProductVersion(v[0], v[1], v[2], 0, lang);
    vi.setStringValues({ lang, codepage: 1200 }, { ProductName: 'YourOfficeSuite', FileDescription: 'YourOfficeSuite', CompanyName: 'YourOfficeSuite', OriginalFilename: 'YourOfficeSuite.exe', InternalName: 'YourOfficeSuite', LegalCopyright: 'Free software under AGPL-3.0' });
    vi.outputToResourceEntries(res.entries);
  }
  res.outputResource(exe);
  fs.writeFileSync(exePath, Buffer.from(exe.generate()));
};
