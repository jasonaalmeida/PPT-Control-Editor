# Deployment Options

## Important: Local `C:\` paths do NOT work

Office Add-ins on Windows require the manifest to live on a **network share** (UNC path) — the Trusted Add-in Catalog field only accepts paths like `\\PC-NAME\ShareName`, never local `C:\...` paths. This is a Microsoft platform limitation, not a bug.

That gives you two realistic distribution paths:

- **Network-share sideload** — every recipient shares a folder on their own PC (technical)
- **Web-hosted files + shared manifest** — you host files once, recipients only need the manifest file (easier for them)

For anything beyond a single technical user, **web-hosted is the better path.**

---

## Option 1: Network-Share Sideload (single user or small technical team)

**Recipient steps** (see [README.md](README.md) for the full walkthrough):

1. Extract ZIP to a folder
2. Right-click folder → Properties → Sharing → Share (add themselves, Read permission)
3. Note the UNC path (`\\PC-NAME\FolderName`)
4. PowerPoint → File → Options → Trust Center → Trusted Add-in Catalogs → paste UNC path → Add catalog → Show in Menu
5. Restart PowerPoint
6. Home → Add-ins → More Add-ins → SHARED FOLDER tab → "Shape Radius"

**Pros:** No hosting needed, works offline once installed  
**Cons:** Every recipient must set up file sharing on their PC. Some corporate machines block sharing entirely.

---

## Option 2: Web-Hosted (recommended for wider sharing)

You host the HTML/JS/CSS files at a public HTTPS URL. Recipients only sideload the manifest (a single file) — much simpler.

### A. Host the files

**GitHub Pages (free):**
```bash
cd C:\ProjectsAndCode\Personal\PPT-Control-Editor
git init
git add src assets
git commit -m "Initial commit"
gh repo create PPT-Control-Editor --public --source=. --push
```
Then enable Pages: repo Settings → Pages → Source: `main` branch, `/` root.
Your files land at: `https://YOUR-USERNAME.github.io/PPT-Control-Editor/src/taskpane/taskpane.html`

**Alternatives:** Netlify, Vercel, Azure Static Web Apps (all free tiers, drag-drop upload).

### B. Update manifest.xml

Replace every `https://localhost:3000` with your hosted base URL:

```xml
<SourceLocation DefaultValue="https://YOUR-USERNAME.github.io/PPT-Control-Editor/src/taskpane/taskpane.html" />
<!-- and the icon URLs, AppDomain, etc. -->
```

### C. Distribute

**For recipients on the same M365 tenant as an admin:**
- Admin uploads manifest via Microsoft 365 Admin Center → Settings → Integrated apps → Upload custom apps
- Deploys to selected users/groups
- Add-in appears automatically in Insert → My Add-ins → **Admin managed**
- **Recipients do nothing** — zero setup on their end

**For recipients without an M365 admin (still uses network share, but only for the manifest):**
- Share the single `manifest.xml` file with them
- They still need to put it in a network-shared folder and add it as a Trusted Catalog (same steps as Option 1)
- But it's a 2 KB file, not a full add-in — no code to keep in sync

**Pros:** You can update the code without redistributing anything. Centralized deployment = zero user setup.  
**Cons:** Requires public hosting. Users need internet.

---

## Option 3: Microsoft 365 Centralized Deployment

**For:** Organizations where an M365 admin can deploy to many users

**Steps:**
1. Host the add-in publicly (Option 2 above)
2. Admin uploads manifest at admin.microsoft.com → Settings → Integrated apps → Upload custom apps → Provide manifest
3. Assign to users, groups, or entire tenant
4. Add-in appears automatically for all assigned users under Insert → My Add-ins → Admin Managed

**Pros:** Zero user action, revocable, updatable  
**Cons:** Needs tenant admin access

---

## Option 4: Microsoft AppSource (Public Marketplace)

**For:** Public distribution to any PowerPoint user in the world

**Steps:**
1. Create a Microsoft Partner Center account
2. Host the add-in on a production-quality HTTPS endpoint
3. Submit for certification (security review, UI/accessibility checks)
4. If approved, publicly listed in the Office Store — searchable inside PowerPoint

**Pros:** Global reach, no per-user install steps  
**Cons:** 2–4 week review, ongoing compliance, must maintain the hosting

---

## Quick Recommendation Table

| Audience                    | Best option                                     |
| --------------------------- | ----------------------------------------------- |
| Just yourself               | Option 1 (network share on your PC)             |
| 2–10 technical friends      | Option 1 (each sets up their own share)         |
| 10–1000 users in your org   | **Option 3 (centralized deployment)**           |
| External users, no admin    | **Option 2 (web-hosted + shared manifest)**     |
| The general public          | Option 4 (AppSource)                            |
