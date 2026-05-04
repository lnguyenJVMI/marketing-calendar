# Free / Low-Cost Marketing Calendar + Broadcast Schedule

This starter kit reproduces the core outcome shown in the shared page: a **single standalone HTML calendar** with a full marketing calendar tab and a dedicated broadcast schedule tab. The recommended version costs **$0 in software and hosting** if your organization already has Asana and Microsoft 365/SharePoint, and it can remain close to $0 by using free static hosting.

## What the shared example is doing

The shared task appears to combine three ingredients into one browser-viewable page. First, it reads dated marketing work from **Asana**. Second, it reads a **Communications Calendar Excel workbook**, including dated broadcast/TV show records and a Tapings sheet. Third, it generates a static file called `marketing-calendar-standalone.html` with two tabs: **Marketing Calendar** for all items and **Broadcast Schedule** for TV/broadcast items only.

| Component | Free or low-cost replacement | Notes |
|---|---:|---|
| Calendar builder | Python script in this kit | Runs locally, on a laptop, or in GitHub Actions. |
| Asana data | CSV export or Asana API personal access token | CSV export is simplest. API sync is better later, but the token must remain secret. |
| Communications data | Excel workbook from SharePoint/OneDrive | The script reads `.xlsx` directly. |
| Web hosting | Azure Static Web Apps Free or GitHub Pages | Both support static HTML hosting. Azure is more Microsoft-native. |
| SharePoint display | SharePoint Embed web part | A site collection admin may need to allow/whitelist the hosting domain. |
| Automation | Manual export first; optional GitHub Actions later | Start manual to avoid paid connectors. Add automation only after the format is stable. |

## Recommended zero-cost architecture

The lowest-friction approach is to generate a **static HTML file** instead of building a database-backed application. Static HTML is cheap, safe, easy to host, and easy to embed. The pipeline is:

> Asana CSV export + Communications Excel workbook → Python generator → `marketing-calendar-standalone.html` → Azure Static Web Apps Free or GitHub Pages → SharePoint Embed web part.

This keeps the moving parts minimal. It also avoids paid automation platforms, paid middleware, and custom SharePoint Framework development.

## How to run the starter kit

Install Python dependencies if needed. The sandbox already includes `pandas` and `openpyxl`, but on a local machine you may need:

```bash
pip install pandas openpyxl
```

Create sample files and generate a demo calendar:

```bash
python3 make_sample_data.py
python3 generate_calendar.py \
  --asana-csv sample_asana_tasks.csv \
  --communications-xlsx sample_communications_calendar.xlsx \
  --output marketing-calendar-standalone.html \
  --title "Marketing Calendar"
```

Then open `marketing-calendar-standalone.html` in a browser. Replace the sample files with your real exports when ready.

## Data requirements

The script is intentionally flexible. It searches for common column names rather than requiring one exact schema.

| Data type | Columns it tries to detect | Examples |
|---|---|---|
| Date | `Due Date`, `Date`, `Air Date`, `Start Date`, `Publish Date`, `Send Date` | `2026-05-11` |
| Title | `Name`, `Task Name`, `Title`, `Subject`, `Episode`, `Show` | `Pentecost and Prophecy` |
| Category | `Type`, `Category`, `Channel`, `Section`, `Project`, `Medium` | `Broadcast/TV` |
| Owner | `Assignee`, `Owner`, `Lead`, `Responsible` | `TV Team` |
| Status | `Status`, `Completed`, `Progress` | `Upcoming` |
| Notes | `Notes`, `Description`, `Details`, `Summary` | `Weekly Monday show` |

The **Broadcast Schedule** tab includes records marked as `Broadcast/TV` or records containing words such as `broadcast`, `TV`, `television`, `show`, `episode`, `air`, `aired`, or `taping`. A sheet named `Tapings` is used to create taping date chips.

## Hosting for free

### Option A: Azure Static Web Apps Free

This is the best fit if SharePoint is the final destination. Microsoft states that Azure Static Web Apps has a **Free plan** that provides free web hosting, SSL, and custom domain support.[^azure-static] Upload or connect a GitHub repository containing `marketing-calendar-standalone.html`.

### Option B: GitHub Pages

GitHub Pages publishes static HTML, CSS, and JavaScript files directly from a repository. GitHub documents that Pages is available for public repositories with GitHub Free and GitHub Free for organizations.[^github-pages] This is usually the simplest no-cost host, but the calendar URL will be public unless you are on a paid/private plan or use another private hosting approach.

## SharePoint embedding note

SharePoint site collection administrators control whether contributors can embed external websites. Microsoft documents that admins can block iframes, allow iframes from any domain, or allow iframes only from specified domains through **HTML Field Security**.[^sharepoint-embed] In practice, this means your SharePoint admin may need to whitelist your Azure Static Web Apps domain, GitHub Pages domain, or custom domain.

## Suggested rollout plan

| Phase | Goal | Cost impact |
|---|---|---:|
| 1 | Use CSV/XLSX exports and generate the HTML manually | $0 |
| 2 | Host the HTML on Azure Static Web Apps Free or GitHub Pages | $0 |
| 3 | Ask SharePoint admin to whitelist the hosting domain | Usually $0 |
| 4 | Add scheduled rebuilds with GitHub Actions or a local scheduled task | Usually $0 |
| 5 | Only if necessary, add API sync from Asana | Usually $0, but requires secure token handling |

## When not to use this approach

This static approach is ideal for **read-only dashboards**. It is not ideal if users need to edit tasks directly inside the calendar, write back to Asana, apply SharePoint item-level permissions inside the HTML file, or hide confidential data on a public URL. In those cases, use Azure Static Web Apps with authentication, SharePoint Lists/Power Apps, or a custom Microsoft 365 application.

## Sources

[^sharepoint-embed]: Microsoft Support, [Allow or restrict the ability to embed content on SharePoint pages](https://support.microsoft.com/en-us/office/allow-or-restrict-the-ability-to-embed-content-on-sharepoint-pages-e7baf83f-09d0-4bd1-9058-4aa483ee137b).
[^azure-static]: Microsoft Azure, [Static Web Apps pricing](https://azure.microsoft.com/en-us/pricing/details/app-service/static/).
[^github-pages]: GitHub Docs, [What is GitHub Pages?](https://docs.github.com/en/pages/getting-started-with-github-pages/about-github-pages).
