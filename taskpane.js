Office.onReady(() => {
  const button = document.getElementById("review-button");
  const request = document.getElementById("request");
  const status = document.getElementById("status");
  const suggestions = document.getElementById("suggestions");
  const suggestionText = document.getElementById("suggestion-text");

  button.addEventListener("click", async () => {
    button.disabled = true;
    status.textContent = "Checking the presentation…";

    try {
      let slideCount = 0;
      await PowerPoint.run(async (context) => {
        const slides = context.presentation.slides;
        slides.load("items");
        await context.sync();
        slideCount = slides.items.length;
      });

      const focus = request.value.trim();
      suggestionText.textContent = slideCount
        ? `The add-in is connected to PowerPoint and found ${slideCount} slide${slideCount === 1 ? "" : "s"}.${focus ? ` Your focus: “${focus}”` : " Add a request above to guide the review."}`
        : "No slides were found in this presentation.";
      suggestions.hidden = false;
      status.textContent = "Connection confirmed. The AI review will be added next.";
    } catch (error) {
      status.textContent = `Could not connect to PowerPoint: ${error.message || error}`;
    } finally {
      button.disabled = false;
    }
  });
});
