# Synthesia Template Builder

Build premium, empty slide templates for training projects and download them as PowerPoint, ready to upload to Synthesia.

## What it does

- **1 · Prompt tab:** type a prompt such as `Maternity leave project, create 8 template slides`. The app reads the project name, the slide count (3–20) and where to leave space for the avatar (`avatar on the left`, `no avatar`).
- **2 · Template tab:** the slides are built with a design unique to the project (palette, motif, heading style and corners). You can change each slide's layout, add or remove slides, and shuffle the design.
- **Fill slides from a document:** upload a Word document (.docx), PowerPoint deck (.pptx), PDF or text file. Headings become slide titles, the text fills the blocks, and each section goes into the speaker notes so it can be used as the avatar script in Synthesia.
- **Download PowerPoint:** saves a .pptx file with real shapes and text boxes that you can edit.

Everything runs in the browser. Projects are saved in the browser's local storage, and uploaded files are never sent to a server.

## Run it on your computer

You need [Node.js](https://nodejs.org) 20 or later.

```bash
npm install
npm run dev
```

Open the address shown in the terminal (usually http://localhost:5173).

## Publish it with GitHub Pages

1. Create a repository on GitHub and upload all the files in this folder, including the hidden `.github` folder. Don't upload `node_modules` or `dist` folders.
2. In the repository, go to **Settings → Pages** and set **Source** to **GitHub Actions**.
3. Push to the `main` branch (or run the **Deploy to GitHub Pages** workflow from the **Actions** tab).
4. When the workflow finishes, the site is live at `https://<your-user-name>.github.io/<repository-name>/`.

## Build a static copy

```bash
npm run build
```

The finished site is in the `dist` folder and can be hosted on any static web host.

## Project structure

| Path | What's in it |
| --- | --- |
| `src/pages/home.tsx` | The whole screen: projects list, Prompt and Template tabs |
| `src/data/design.ts` | Palettes, motifs, the 14 slide layouts and text fitting |
| `src/data/fill.ts` | Rules that place document sections into slide layouts |
| `src/data/projects.ts` | Project model, prompt reading and browser storage |
| `src/lib/pptx.ts` | PowerPoint (.pptx) writer, including speaker notes |
| `src/lib/docimport.ts` | Reads .docx, .pptx, .pdf, .txt and .md files |

## Limits

- Layouts are chosen by fixed rules, not AI, so documents with clear headings give the best results.
- Scanned PDFs, images, and old .doc or .ppt files can't be read. Save them as .docx or .pptx first.
