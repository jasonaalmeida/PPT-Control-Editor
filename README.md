# Shape Radius — PowerPoint Add-in

Set numeric radius/adjustment values on one or many selected shapes in PowerPoint.

## What it does

- **Set corner radius** on rounded rectangles by typing a percent (0–100%)
- **Apply to multiple shapes** at once
- **Read current values** from selected shapes (auto or manual)
- **Works with other shapes** (arrows, stars, callouts) using raw 0–1 mode

## Requirements

- **PowerPoint Desktop on Windows** (Microsoft 365)
- **Office Insider/Beta channel** (for the preview `Shape.adjustments` API)
  - Stable Office builds do not support this API yet
  - Switch via: File → Account → Office Insider → Beta Channel

## Installation

Office Add-ins can't be loaded from a plain `C:\...` folder. The trusted catalog needs a **network share** (UNC path like `\\PC-NAME\ShareName`). You have two options:

### Option A — Share the folder locally (single user)

**1. Extract this ZIP** to a permanent folder (e.g., `C:\PowerPointAddins\ShapeRadius`)

**2. Share the folder as a network share:**
   - Right-click the extracted folder → **Properties** → **Sharing** tab
   - Click **Share...**
   - Add **yourself** (or "Everyone") with at least **Read** permission
   - Click **Share** → note the network path shown, e.g., `\\YOUR-PC-NAME\ShapeRadius`
   - Click **Done** → **Close**

**3. Trust the network path in PowerPoint:**
   - Open PowerPoint
   - **File → Options → Trust Center → Trust Center Settings**
   - **Trusted Add-in Catalogs** tab
   - In **Catalog Url**, paste the **UNC path** from step 2 (e.g., `\\YOUR-PC-NAME\ShapeRadius`)
   - Click **Add catalog**
   - ✅ Check **"Show in Menu"**
   - Click **OK** → **OK**

**4. Restart PowerPoint**

**5. Insert the add-in:**
   - On the **Home** tab, click the **Add-ins** button
   - In the flyout, click **More Add-ins** (or **Advanced**) at the bottom
   - In the **Office Add-ins** dialog, click the **SHARED FOLDER** tab at the top
   - Select **"Shape Radius"** → **Add**

   *(On older PowerPoint versions the path is Insert → Get Add-ins → SHARED FOLDER.)*

### Option B — Web-hosted (recommended for sharing with others)

Instead of network shares, host the add-in files on a public HTTPS URL and share just the manifest. See [DEPLOY.md](DEPLOY.md) for full steps.

## Usage

1. **Select one or more shapes** on the slide
2. In the taskpane, choose the input mode:
   - **Radius (%)** — for rounded rectangles, 0–100%
   - **0–1 scale** — for arrows, stars, callouts (raw adjustment values)
3. Type a number
4. Click **Apply to selected**

**Auto-read:** With "Auto-read from selected shape" checked (default), the input auto-populates when you select a shape. Uncheck to enter values freely without auto-updates.

**Manual read:** Click **Read from first selected** to pull the current value.

## Troubleshooting

**"I don't see 'SHARED FOLDER' in the Add-ins dialog"**
→ You added the catalog but PowerPoint wasn't restarted. Close all Office apps and reopen.

**"The add-in doesn't appear under SHARED FOLDER"**
→ Verify the folder is actually shared on the network (Properties → Sharing shows it as shared). Verify the manifest.xml is directly in that folder. The catalog URL must be the UNC path (`\\...`), not a local `C:\` path.

**"ApiNotFound" error when clicking Apply**
→ You need Office Insider Beta channel. The `Shape.adjustments` API is in preview.

**"Taskpane is blank"**
→ Don't move or delete the extracted folder. PowerPoint reads files from it every time the add-in runs.

**"Access denied" or the shared folder isn't reachable**
→ Windows Defender Firewall may be blocking network share access. Ensure "File and Printer Sharing" is enabled for your network profile (Private).

## Technical Details

- **Manifest:** `manifest.xml`
- **Code:** `src/taskpane/`
- **API:** PowerPoint JavaScript API 1.10 (preview)
- **License:** MIT

Built with the Office Add-ins platform. No external dependencies.
