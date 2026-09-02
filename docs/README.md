# FURM Site — Structure & How to Publish

## Folder layout
```
furm-web/
├── index.html         Home page (Meet the Board, fall lineup, join the board)
├── linktree.html      Link tree (reads links.json)
├── cubcare.html       CubCare mentor/mentee program page
├── roles.html         E-Board role descriptions
├── links.json         Controls the link tree buttons
├── assets/            All images (logo, photos, QR, banner)
├── events/            Event pages (kickoff, med-school, destress, spring)
├── slides/            PowerPoints + the Python generators
└── docs/              This README + the mailing-list form spec
```

## Things that are DONE but not yet shown (ready to publish)

### Event pages
The pages already exist in `events/`. To link them on the home page, open `index.html`,
find the comment that says `READY TO PUBLISH` on the matching fall card, and uncomment the
`<a class="siLink" ...>Learn more →</a>` line.

To show them on the link tree, open `links.json` and move the entry from the
`_ready_to_publish.links` list up into the main `links` list.

### Spring 2027 lineup
- Full page: `events/spring.html` (already built)
- Home-page section: in `index.html`, find the big `READY TO PUBLISH: Spring 2027` comment
  block and uncomment the whole `<section id="spring">`.

### New board members
In `index.html`, inside the Meet the Board section, there's a `MEMBER STENCIL` comment block.
Copy it, add the person's photo to `assets/`, fill in the details, point the role link to the
right `roles.html` anchor, and uncomment.

## Editing the link tree
Edit `links.json` only — no HTML needed. Each entry is `{ "label", "url", "note" }`.

## Regenerating slides (optional)
From inside `slides/`:
```
../.pptxvenv/bin/python make_slides.py    # landscape concept deck
../.pptxvenv/bin/python make_banner.py    # newsletter banner
```
DO NOT run `make_flyers.py` — the final flyer (`furm_flyers.pptx`) was hand-edited in
PowerPoint and regenerating would overwrite those edits.

## QR codes
The QR images point to a placeholder URL until the site is hosted. Once it's live,
regenerate the QR to point at the real site URL, then re-drop it into the slides.
