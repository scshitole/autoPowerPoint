const pptxgen = require("pptxgenjs");

// Create a new presentation
const ppt = new pptxgen();

// Add a slide with content
const slide = ppt.addSlide();
slide.addText("Hello, PowerPoint!");

// Save the presentation to a file
ppt.writeFile({
  fileName: "example.pptx",
}, (error) => {
  if (error) {
    console.log("Error saving PowerPoint file:", error);
    return;
  }
  console.log("PowerPoint slide created successfully!");
});

