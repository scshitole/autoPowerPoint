const fs = require("fs");
const pptxgen = require("pptxgenjs");

// Create a new presentation
const ppt = new pptxgen();

// Read the text elements from a file
const filePath = "text_file.txt";
const texts = fs.readFileSync(filePath, "utf8").split("\n").map((text) => text.trim());

// Set the initial coordinates and height for the first text element
let x = 1;
let y = 1;
const height = 0.2;
const maxTextsPerSlide = 4;

// Add the text elements to slides
let slide = ppt.addSlide();
let textCounter = 0;
texts.forEach((text) => {
  const textOpts = {
    x, // X coordinate of the text box
    y, // Y coordinate of the text box
    w: 8, // Width of the text box
    h: height, // Height of the text box
    color: "000000", // Text color
    align: "left", // Text alignment
  };

  slide.addText(text, textOpts);
  textCounter++;

  if (textCounter % maxTextsPerSlide === 0) {
    // Create a new slide after every four text elements
    slide = ppt.addSlide();
    x = 1;
    y = 1;
  } else {
    // Increment the y value for the next text element within the set
    y += height + 0.5; // Adjust the increment as needed
  }
});

// Save the presentation to a file
ppt.writeFile({
  fileName: "example.pptx",
}, (error) => {
  if (error) {
    console.log("Error saving PowerPoint file:", error);
    return;
  }
  console.log("PowerPoint slides created successfully!");
});
