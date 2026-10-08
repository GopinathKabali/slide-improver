# Slide Improver

A starter Microsoft 365 PowerPoint task-pane add-in. The current version provides the panel UI and checks that it can communicate with the open presentation. AI suggestions and applying edits are planned next.

## Project files

- `manifest.xml` tells PowerPoint how to load the task pane.
- `taskpane.html` contains the panel interface.
- `taskpane.css` styles the interface.
- `taskpane.js` connects the panel to PowerPoint using Office.js.

## Development status

This starter is not ready for other people to install yet. It needs a local HTTPS web server for development, then a securely hosted web app and AI service for deployment. Do not put AI provider keys in browser code; the later service will keep them on the server.

## Planned first release

1. Read the active slide's text and layout.
2. Ask an AI service for specific improvement suggestions.
3. Let the user review suggestions and approve edits.
4. Apply supported text and layout changes through the PowerPoint JavaScript API.
