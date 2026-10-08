# Slide Improver

A starter Microsoft 365 PowerPoint task-pane add-in. The current version provides the panel UI and checks that it can communicate with the open presentation. AI suggestions and applying edits are planned next.

## Project files

- `manifest.xml` tells PowerPoint how to load the task pane.
- `taskpane.html` contains the panel interface.
- `taskpane.css` styles the interface.
- `taskpane.js` connects the panel to PowerPoint using Office.js.

## Development status

The add-in reads the selected slide and shows basic local checks. ChatGPT sign-in and plan usage are not connected yet. This is still a development version, not ready for other people to install.

## Local setup

1. Install Node.js 20 or newer.
2. In this project folder, run `npm install`.
3. Run `npm run install-dev-cert` once to install a trusted local HTTPS certificate.
4. Run `npm run dev`. The local add-in site will be available at `https://localhost:3000`.
5. In PowerPoint, sideload `manifest.xml` to load the task pane.

The planned AI connection is Sign in with ChatGPT, so each user can authorize their own eligible ChatGPT plan. Distribution and eligibility requirements need to be settled before implementing that connection.

## Planned first release

1. Read the active slide's text and layout.
2. Let the user connect ChatGPT and request specific improvement suggestions using their plan.
3. Let the user review suggestions and approve edits.
4. Apply supported text and layout changes through the PowerPoint JavaScript API.
