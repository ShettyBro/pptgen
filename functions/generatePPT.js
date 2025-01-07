require("dotenv").config();
const { Configuration, OpenAIApi } = require("openai");
const PptxGenJS = require("pptxgenjs");

exports.handler = async (event) => {
  try {
    // Parse user input
    const { topic, numSlides } = JSON.parse(event.body);
    if (!topic || !numSlides) {
      return {
        statusCode: 400,
        body: JSON.stringify({ error: "Topic and number of slides are required" }),
      };
    }

    // Initialize OpenAI
    const openai = new OpenAIApi(
      new Configuration({ apiKey: process.env.OPENAI_API_KEY })
    );

    // Generate text content for slides
    const textResponse = await openai.createCompletion({
      model: "text-davinci-003",
      prompt: `Create a ${numSlides}-slide presentation outline for the topic: ${topic}. Each slide should include a title and content.`,
      max_tokens: 500,
    });

    const slideContents = textResponse.data.choices[0].text
      .trim()
      .split(/\n\n+/) // Split into sections
      .map((chunk) => chunk.trim())
      .filter((chunk) => chunk); // Filter out empty sections

    if (slideContents.length < numSlides) {
      return {
        statusCode: 400,
        body: JSON.stringify({ error: "Insufficient slide content generated." }),
      };
    }

    // Generate images for slides using OpenAI DALL·E
    const imagePromises = Array.from({ length: numSlides }, (_, i) =>
      openai
        .createImage({
          prompt: `A visually appealing and relevant image for slide ${i + 1} on the topic "${topic}"`,
          n: 1,
          size: "512x512", // Smaller size to reduce file size
        })
        .then((response) => response.data.data[0].url)
        .catch((error) => {
          console.error(`Error generating image for slide ${i + 1}:`, error);
          return null; // Use null as fallback
        })
    );
    const images = await Promise.all(imagePromises);

    // Create a PPTX presentation
    const pptx = new PptxGenJS();
    slideContents.forEach((content, index) => {
      const slide = pptx.addSlide();
      slide.addText(`Slide ${index + 1}`, { x: 0.5, y: 0.5, fontSize: 24, bold: true });
      slide.addText(content, { x: 0.5, y: 1.5, fontSize: 18 });

      if (images[index]) {
        slide.addImage({
          data: images[index],
          x: 0.5,
          y: 2.5,
          w: 6,
          h: 3,
        });
      }
    });

    // Write the PPTX to buffer
    const pptBuffer = await pptx.write("arraybuffer");

    // Return the PPTX file as a response
    return {
      statusCode: 200,
      headers: {
        "Content-Type": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        "Content-Disposition": `attachment; filename="${topic.replace(/\s+/g, "_")}.pptx"`,
      },
      body: Buffer.from(pptBuffer).toString("base64"),
      isBase64Encoded: true,
    };
  } catch (error) {
    console.error("Error in generatePPT function:", error);
    return {
      statusCode: 500,
      body: JSON.stringify({ error: "Internal Server Error" }),
    };
  }
};
