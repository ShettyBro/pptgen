require("dotenv").config();
const OpenAI = require("openai");
const PptxGenJS = require("pptxgenjs");

exports.handler = async (event) => {
  try {
    console.log("Function started");
    
    // Log the incoming request
    console.log("Request body:", event.body);
    
    // Parse user input with error handling
    let topic, numSlides;
    try {
      const parsedBody = JSON.parse(event.body);
      topic = parsedBody.topic;
      numSlides = parsedBody.numSlides;
      console.log("Parsed input - Topic:", topic, "Slides:", numSlides);
    } catch (parseError) {
      console.error("Error parsing request body:", parseError);
      return {
        statusCode: 400,
        body: JSON.stringify({ 
          error: "Invalid request body",
          details: parseError.message 
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

    // Log OpenAI key format (safely)
    const apiKey = process.env.OPENAI_API_KEY;
    console.log("API Key format check:", 
      apiKey ? 
      `Key starts with: ${apiKey.substring(0, 7)}... (length: ${apiKey.length})` : 
      "No API key found");

    // Initialize OpenAI with error handling
    let openai;
    try {
      openai = new OpenAI({
        apiKey: process.env.OPENAI_API_KEY,
      });
      console.log("OpenAI client initialized");
    } catch (openaiError) {
      console.error("Error initializing OpenAI:", openaiError);
      return {
        statusCode: 500,
        body: JSON.stringify({ 
          error: "Error initializing OpenAI client",
          details: openaiError.message 
        }),
      };
    }

    // Generate text content for slides
    console.log("Generating slide content...");
    let textResponse;
    try {
      textResponse = await openai.completions.create({
        model: "gpt-3.5-turbo-instruct",
        prompt: `Create a ${numSlides}-slide presentation outline for the topic: ${topic}. Each slide should include a title and content.`,
        max_tokens: 500,
      });
      console.log("Slide content generated successfully");
    } catch (textError) {
      console.error("Error generating text content:", textError);
      return {
        statusCode: 500,
        body: JSON.stringify({ 
          error: "Error generating slide content",
          details: textError.message 
        }),
      };
    }

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

    // Generate images with error handling
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

    // Create PowerPoint with error handling
    console.log("Creating PowerPoint...");
    let pptBuffer;
    try {
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

      pptBuffer = await pptx.write("arraybuffer");
      console.log("PowerPoint created successfully");
    } catch (pptError) {
      console.error("Error creating PowerPoint:", pptError);
      return {
        statusCode: 500,
        body: JSON.stringify({ 
          error: "Error creating PowerPoint",
          details: pptError.message 
        }),
      };
    }

    console.log("Function completed successfully");
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
    console.error("Unexpected error in generatePPT function:", error);
    return {
      statusCode: 500,
      body: JSON.stringify({ 
        error: "Internal Server Error",
        details: error.message,
        stack: error.stack 
      }),
    };
  }
};