# Tracker Manager — Office Add-in

A mock Office Add-in task pane for the Project Document Tracker.
All UI is functional. Excel integration (actual sheet manipulation) is stubbed out
and ready for Office.js code to be wired in.

---

## Files

| File | Purpose |
|------|---------|
| `manifest.xml` | Tells Excel what the add-in is and where to find it |
| `taskpane.html` | The task pane UI (all three tabs) |
| `commands.html` | Required placeholder for ribbon commands |
| `assets/` | Icon files (see below) |

---

## Step 1 — Replace the placeholder URL

Open `manifest.xml` and replace every instance of:

```
YOUR-GITHUB-USERNAME
```

with your actual GitHub username. For example if your username is `jsmith`:

```
https://jsmith.github.io/tracker-addin/taskpane.html
```

---

## Step 2 — Host on GitHub Pages

1. Create a new GitHub repository called `tracker-addin`
2. Upload all files to the root of the repo:
   - `manifest.xml`
   - `taskpane.html`
   - `commands.html`
   - `assets/icon-16.png`, `icon-32.png`, `icon-64.png`, `icon-80.png`
3. Go to **Settings → Pages → Source → main branch / root**
4. Click Save. Your files will be live at `https://YOUR-USERNAME.github.io/tracker-addin/` within ~60 seconds.

> **Icons:** You need four small PNG icons (16×16, 32×32, 64×64, 80×80 px).
> Any simple icon will do for a mockup — even a solid coloured square saved as PNG.
> Free tool: https://favicon.io/favicon-generator/

---

## Step 3 — Sideload into Excel (desktop)

### Windows
1. Open Excel desktop
2. Go to **File → Options → Trust Center → Trust Center Settings → Trusted Add-in Catalogs**
3. Add a new catalog path pointing to a **shared network folder** (e.g. `\\MyPC\AddinShare`)
4. Copy `manifest.xml` into that folder
5. Restart Excel
6. Go to **Insert → My Add-ins → Shared Folder tab** → select **Tracker Manager** → Add

### Mac
1. Open Finder → Go → Go to Folder → paste:
   ```
   /Users/<your-username>/Library/Containers/com.microsoft.Excel/Data/Documents/wef
   ```
2. Copy `manifest.xml` into that folder
3. In Excel: **Insert → My Add-ins → Developer Add-ins** → Tracker Manager

---

## Step 4 — Sideload into Excel Online (SharePoint)

1. Open your workbook in Excel Online via SharePoint
2. Go to **Insert → Add-ins → Upload My Add-in**
3. Browse to `manifest.xml` and click Upload
4. The **"Open Tracker Manager"** button appears in the **Home** tab of the ribbon

> Note: Excel Online sideloading requires admin permissions in some Microsoft 365 tenants.
> If you see a permissions error, ask your M365 admin to enable developer sideloading,
> or deploy via the Microsoft 365 admin centre instead.

---

## Adding real functionality (future)

All buttons show toast notifications as placeholders. To wire up real Excel calls:

1. Open `taskpane.html`
2. Find the `mockAction()` function and `setView()` function
3. Replace `showToast(...)` calls with `Office.js` API calls, e.g.:

```javascript
// Example: read the used range
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getItem("Project Tracker");
  const range = sheet.getUsedRange();
  range.load("values");
  await context.sync();
  console.log(range.values);
});
```

Full Office.js docs: https://learn.microsoft.com/en-us/office/dev/add-ins/excel/

---

*Tracker Manager v1.0 — mockup build*
