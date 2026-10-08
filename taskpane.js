Office.onReady(() => {
  const button = document.getElementById("review-button");
  const status = document.getElementById("status");
  const suggestions = document.getElementById("suggestions");
  const slideSummary = document.getElementById("slide-summary");
  const suggestionList = document.getElementById("suggestion-list");

  button.addEventListener("click", async () => {
    button.disabled = true;
    status.textContent = "Reviewing the selected slide…";

    try {
      let slideReview;
      await PowerPoint.run(async (context) => {
        const selectedSlides = context.presentation.getSelectedSlides();
        selectedSlides.load("items");
        await context.sync();

        if (!selectedSlides.items.length) {
          throw new Error("Select a slide in PowerPoint, then try again.");
        }

        const slide = selectedSlides.items[0];
        const shapes = slide.shapes;
        shapes.load("items");
        await context.sync();

        shapes.items.forEach((shape) => shape.load("type,left,top,width,height"));
        await context.sync();

        const textShapes = shapes.items.filter((shape) =>
          shape.type === PowerPoint.ShapeType.textBox || shape.type === PowerPoint.ShapeType.placeholder
        );
        textShapes.forEach((shape) => shape.textFrame.textRange.load("text"));
        await context.sync();

        const blocks = textShapes
          .map((shape) => ({
            text: shape.textFrame.textRange.text.trim(),
            left: shape.left,
            top: shape.top,
            width: shape.width,
            height: shape.height,
          }))
          .filter((block) => block.text);

        slideReview = { shapeCount: shapes.items.length, blocks };
      });

      const suggestionsForSlide = makeSuggestions(slideReview);
      slideSummary.textContent = `${slideReview.blocks.length} text area${slideReview.blocks.length === 1 ? "" : "s"} and ${slideReview.shapeCount} total object${slideReview.shapeCount === 1 ? "" : "s"} found on the selected slide.`;
      suggestionList.replaceChildren(...suggestionsForSlide.map((message) => {
        const item = document.createElement("li");
        item.textContent = message;
        return item;
      }));
      suggestions.hidden = false;
      status.textContent = "Slide checks are ready. No slide content was changed.";
    } catch (error) {
      status.textContent = `Review could not be completed: ${error.message || error}`;
    } finally {
      button.disabled = false;
    }
  });
});

function makeSuggestions(review) {
  const messages = [];
  const words = review.blocks.flatMap((block) => block.text.split(/\s+/).filter(Boolean));
  const longBlocks = review.blocks.filter((block) => block.text.split(/\s+/).filter(Boolean).length > 40);
  const textObjects = review.blocks.length;

  if (!textObjects) {
    messages.push("No editable text was found. The slide may use images or grouped objects for its text.");
  } else {
    const firstLine = review.blocks[0].text.split(/\r?\n/).find((line) => line.trim()) || "";
    if (firstLine.length > 90) {
      messages.push("The first text area is long for a title. Consider shortening it to make the main point easier to scan.");
    } else {
      messages.push("Check that the first text area states the slide’s main takeaway as a short heading.");
    }

    if (words.length > 90) {
      messages.push(`This slide has about ${words.length} words. Consider moving detail to speaker notes or splitting the content across slides.`);
    }

    if (longBlocks.length) {
      messages.push(`${longBlocks.length} text area${longBlocks.length === 1 ? " has" : "s have"} more than 40 words. Break dense paragraphs into shorter bullets or a simple visual.`);
    }

    if (textObjects > 8) {
      messages.push(`There are ${textObjects} separate text areas. Group related points and align them to a consistent grid.`);
    }

    if (messages.length === 1) {
      messages.push("The slide has a manageable amount of text. Check that spacing, font sizes, and alignment are consistent.");
    }
  }

  return messages;
}
