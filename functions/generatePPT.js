require("dotenv").config();
const OpenAI = require("openai");
const PptxGenJS = require("pptxgenjs");
const axios = require("axios");

exports.handler = async (event) => {
  try {
    console.log("Function started");

    // Log the incoming request
    console.log("Request body:", event.body);

    // Parse user input with error handling
    let topic, numSlides, webhookUrl;
    try {
      const parsedBody = JSON.parse(event.body);
      topic = parsedBody.topic;
      numSlides = parsedBody.numSlides;
      webhookUrl = "https://hook.eu2.make.com/6v32ociz9ki7rw9mahy6anfl30eds4yh"; // Replace with your Make.com webhook URL
      console.log("Parsed input - Topic:", topic, "Slides:", numSlides);
    } catch (parseError) {
      console.error("Error parsing request body:", parseError);
      return {
        statusCode: 400,
        body: JSON.stringify({
          error: "Invalid request body",
          details: parseError.message,
        }),
      };
    }

    if (!topic || !numSlides) {
      console.error("Missing required fields");
      return {
        statusCode: 400,
        body: JSON.stringify({ error: "Topic and number of slides are required" }),
      };
    }

    // Initialize OpenAI
    const openai = new OpenAI({
      apiKey: process.env.OPENAI_API_KEY,
    });

    // Generate text content for slides
    console.log("Generating slide content...");
    const textResponse = await openai.completions.create({
      model: "gpt-3.5-turbo-instruct",
      prompt: `Create a ${numSlides}-slide presentation outline for the topic: ${topic}. Each slide should include a title and content.`,
      max_tokens: 500,
    });

    const slideContents = textResponse.choices[0].text
      .trim()
      .split(/\n\n+/)
      .map((chunk) => chunk.trim())
      .filter((chunk) => chunk);

    console.log(`Generated ${slideContents.length} slides content`);

    if (slideContents.length < numSlides) {
      console.error("Insufficient content generated");
      return {
        statusCode: 400,
        body: JSON.stringify({ error: "Insufficient slide content generated." }),
      };
    }

    // Generate images for slides
    console.log("Generating images...");
    const images = [];
    for (let i = 0; i < numSlides; i++) {
      try {
        const imageResponse = await openai.images.generate({
          prompt: `A visually appealing and relevant image for slide ${i + 1} on the topic "${topic}"`,
          n: 1,
          size: "512x512",
        });
        images.push(imageResponse.data[0].url);
        console.log(`Generated image for slide ${i + 1}`);
      } catch (imageError) {
        console.error(`Error generating image for slide ${i + 1}:`, imageError);
        images.push(null);
      }
    }

    // Create PowerPoint
    console.log("Creating PowerPoint...");
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

    const pptBuffer = await pptx.write("arraybuffer");
    console.log("PowerPoint created successfully");

    // Send the generated file to Make.com webhook
    console.log("Sending file to Make.com webhook...");
    try {
      const response = await axios.post(webhookUrl, pptBuffer, {
        headers: {
          "Content-Type": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
          "Content-Disposition": `attachment; filename="${topic.replace(/\s+/g, "_")}.pptx"`,
        },
      });
      console.log("File sent successfully to webhook:", response.status);
    } catch (webhookError) {
      console.error("Error sending file to webhook:", webhookError);
      return {
        statusCode: 500,
        body: JSON.stringify({
          error: "Error sending file to webhook",
          details: webhookError.message,
        }),
      };
    }

    return {
      statusCode: 200,
      body: JSON.stringify({ message: "PowerPoint generated and sent to Make.com successfully" }),
    };
  } catch (error) {
    console.error("Unexpected error:", error);
    return {
      statusCode: 500,
      body: JSON.stringify({
        error: "Internal Server Error",
        details: error.message,
        stack: error.stack,
      }),
    };
  }
};
