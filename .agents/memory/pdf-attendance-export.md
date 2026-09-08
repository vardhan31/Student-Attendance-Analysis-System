---
name: Browser PDF attendance handling
description: The attendance app keeps PDF parsing local and uses a small in-browser PDF writer for exports.
---

The project environment may reject otherwise standard PDF-export packages at the package firewall, so browser-side PDF reports can be generated without a third-party writer when the report only needs text and simple vector chart bars. PDF.js workers should be imported with Vite's `?url` handling so the deployed asset has a real `GlobalWorkerOptions.workerSrc`.

**Why:** A blocked package download should not force backend processing or external services into this local-first app.

**How to apply:** Keep PDF input extraction in pdfjs-dist, configure the matching worker as a bundled URL, and treat multi-page attendance PDFs as one document with a reusable first-page header mapping. Preserve the local export path unless the product requirements expand to complex typography, images, or pagination.