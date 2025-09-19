import 'dotenv/config';
import express from 'express';
import path from 'path';
import { GoogleGenerativeAI } from '@google/generative-ai';
import multer from 'multer';
import fs from 'fs/promises'; // For promise-based functions
import fsSync from 'fs'; // For synchronous functions like existsSync
import { v4 as uuidv4 } from 'uuid';
import fetch from 'node-fetch';
import doc2txt from './doc2txt.js';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const storage = multer.diskStorage({
  destination: function (req, file, cb) {
    cb(null, './uploads/');
  },
  filename: function (req, file, cb) {
    const fileExtension = path.extname(file.originalname);
    const newFileName = file.fieldname + '-' + Date.now() + fileExtension;
    cb(null, newFileName);
  }
});

const upload = multer({ 
    storage: storage,
    limits: {
        fileSize: 10 * 1024 * 1024, // 10MB limit
    },
    fileFilter: function (req, file, cb) {
        const allowedTypes = ['.pdf', '.doc', '.docx', '.txt'];
        const fileExtension = path.extname(file.originalname).toLowerCase();
        
        if (allowedTypes.includes(fileExtension)) {
            cb(null, true);
        } else {
            cb(new Error('Invalid file type. Only PDF, DOC, DOCX, and TXT files are allowed.'));
        }
    }
});

class App {
    constructor() {
        console.log('🚀 Initializing AI Presentation Generator...');
        this.app = express();
        this.port = process.env.PORT || 3001;
        this.unsplashAccessKey = process.env.UNSPLASH_ACCESS_KEY;
        this.configure();
        this.routes();
        this.setupErrorHandling();
        this.lastSearchTime = 0;
        this.THROTTLE_MS = 5000; // 5 seconds between image requests
    }

    configure() {
        console.log('⚙️ Configuring server...');
        this.app.set('view engine', 'ejs');
        this.app.set('views', path.join(__dirname, 'views'));
        this.app.use(express.json({ limit: '50mb' }));
        this.app.use(express.urlencoded({ extended: true }));
        this.app.use('/static', express.static(path.join(__dirname, 'static')));
        
        // Create uploads directory if it doesn't exist
        if (!fsSync.existsSync('./uploads')) {
            fsSync.mkdirSync('./uploads');
            console.log('📁 Created uploads directory');
        }

        // CORS middleware for API requests
        this.app.use((req, res, next) => {
            res.header('Access-Control-Allow-Origin', '*');
            res.header('Access-Control-Allow-Headers', 'Origin, X-Requested-With, Content-Type, Accept');
            next();
        });
    }

    routes() {
        console.log('🛣️ Setting up routes...');
        
        // Main routes
        this.app.post('/api/generate', upload.single('file'), this.generateSlides.bind(this));
        this.app.get('/api/image', this.getImage.bind(this));
        this.app.get('/api/health', this.healthCheck.bind(this));
        
        // ✅ NEW ENDPOINT: To serve the temporary images
        this.app.get('/api/images/:folder/:filename', this.serveImage.bind(this));
        
        // Home page
        this.app.get('/', (req, res) => {
            res.render('index', { 
                title: 'AI Presentation Generator',
                version: '2.0',
                features: [
                    'AI-Powered Content Generation',
                    'Beautiful Modern Templates', 
                    'Automatic Image Integration',
                    'Professional PPTX Export'
                ]
            });
        });

        // 404 handler
        this.app.use((req, res) => {
            res.status(404).json({ error: 'Endpoint not found' });
        });
    }

    setupErrorHandling() {
        // Global error handler
        this.app.use((error, req, res, next) => {
            console.error('❌ Server Error:', error);
            
            if (error instanceof multer.MulterError) {
                if (error.code === 'LIMIT_FILE_SIZE') {
                    return res.status(400).json({ error: 'File too large. Maximum size is 10MB.' });
                }
            }
            
            res.status(500).json({ 
                error: 'Internal server error',
                message: process.env.NODE_ENV === 'development' ? error.message : 'Something went wrong'
            });
        });
    }

    async healthCheck(req, res) {
        const health = {
            status: 'healthy',
            timestamp: new Date().toISOString(),
            services: {
                server: 'running',
                gemini: process.env.GEMINI_API_KEY ? 'configured' : 'not configured',
                unsplash: this.unsplashAccessKey ? 'configured' : 'not configured',
                uploads: fsSync.existsSync('./uploads') ? 'available' : 'missing'
            }
        };
        
        res.json(health);
    }

async generateSlides(req, res) {
    const startTime = Date.now();
    const { topic, aiModel } = req.body;
    let sourceText = '';
    let imagePayloads = [];

    console.log(`🎯 Generating presentation for: "${topic}" using ${aiModel}`);

    if (!topic || topic.trim().length === 0) {
        return res.status(400).json({ error: 'Topic is required and cannot be empty.' });
    }

    if (topic.length > 200) {
        return res.status(400).json({ error: 'Topic is too long. Please keep it under 200 characters.' });
    }

    if (req.file) {
        try {
            console.log(`📄 Processing uploaded file: ${req.file.originalname}`);
            const extractionResult = await doc2txt.extractTextFromFile(req.file.path);
            sourceText = extractionResult.text;
            const extractedImages = extractionResult.images;
            
            await fs.unlink(req.file.path); 
            if (sourceText.length > 10000) {
                sourceText = sourceText.substring(0, 10000) + '...';
                console.log('📝 Source text truncated to 10,000 characters');
            }

            if (extractedImages && extractedImages.length > 0) {
                const tempDirId = `request-${uuidv4()}`;
                const tempDirPath = path.join(__dirname, 'temp_images', tempDirId);
                await fs.mkdir(tempDirPath, { recursive: true });

                for (const [index, imageObj] of extractedImages.entries()) {
                    const imageBase64 = imageObj.base64; 

                    const mimeMatch = imageBase64.match(/^data:image\/(jpeg|png);base64,/);
                    if (mimeMatch) {
                        const fileType = mimeMatch[1];
                        const extension = fileType === 'jpeg' ? '.jpg' : '.png';
                        const base64Data = imageBase64.replace(/^data:image\/(jpeg|png);base64,/, "");
                        const buffer = Buffer.from(base64Data, 'base64');
                        const filename = `image_${index + 1}${extension}`;
                        const imagePath = path.join(tempDirPath, filename);
                        
                        await fs.writeFile(imagePath, buffer);
                        
                        const relativeUrl = `/api/images/${tempDirId}/${filename}`;
                        
                        imagePayloads.push({
                            url: relativeUrl,
                            width: imageObj.width,
                            height: imageObj.height
                        });
                    }
                }
                console.log(`✅ Saved ${imagePayloads.length} images to temp directory`);
            }
        } catch (error) {
            console.error('❌ Error parsing file:', error);
            throw error;
        }
    }

    try {
        let slidesData;
        let slides;

        if (aiModel === 'Gemini' && process.env.GEMINI_API_KEY) {
            console.log('🤖 Using Gemini AI...');
            slidesData = await this.generateSlidesWithGemini(topic, sourceText);
            // Gemini returns { "slides": [...] }
            slides = slidesData.slides;
        } else if (aiModel === 'LMStudio') {
            console.log('🤖 Using LM Studio...');
            // LM Studio returns a direct array of slides
            slides = await this.generateSlidesWithLMStudio(topic, sourceText);
        } else {
            return res.status(400).json({
                error: 'Invalid AI model selected or API key not configured.'
            });
        }

        // ✅ Centralize the slide validation
        if (!slides || !Array.isArray(slides)) {
            throw new Error('Invalid response format from AI service');
        }

        // ✅ Consolidate all slide data into a single object for consistency
        const finalSlidesData = {
            slides,
            metadata: {
                topic,
                aiModel,
                slideCount: slides.length,
                generationTime: Date.now() - startTime,
                hasSourceDocument: !!sourceText,
                hasImages: imagePayloads.length > 0
            }
        };

        // 🖼️ Step 2: Consolidate images into multi-image slides
        if (imagePayloads.length > 0) {
            imagePayloads.sort((a, b) => (b.width * b.height) - (a.width * a.height));
            const imageSlides = this.createImageSlidesFromPayloads(imagePayloads);
            finalSlidesData.slides.push(...imageSlides);
            console.log(`🖼️ Appended ${imageSlides.length} image slides.`);
        }

        if (finalSlidesData.slides.length === 0) {
            throw new Error('No slides generated');
        }

        console.log(`✅ Generated ${finalSlidesData.slides.length} slides in ${Date.now() - startTime}ms`);
        res.json(finalSlidesData);

    } catch (error) {
        console.error('❌ Error generating slides:', error);
        let errorMessage = 'Failed to generate slides.';
        if (error.message.includes('API key')) {
            errorMessage = 'AI service not configured properly. Please check API keys.';
        } else if (error.message.includes('quota')) {
            errorMessage = 'AI service quota exceeded. Please try again later.';
        } else if (error.message.includes('network') || error.message.includes('fetch') || error.message.includes('ECONNREFUSED')) {
            errorMessage = 'Network error or AI service is not running.';
        } else if (error.message.includes('parse') || error.message.includes('Invalid response')) {
            errorMessage = 'Failed to parse AI response. The response may be invalid.';
        }

        res.status(500).json({
            error: errorMessage,
            details: process.env.NODE_ENV === 'development' ? error.message : undefined
        });
    }
}


// ✅ NEW HELPER FUNCTION: To consolidate images into slides
// This is an example of a simple packing algorithm.
createImageSlidesFromPayloads(imagePayloads) {
    const slides = [];
    let currentSlide = null;
    let currentImageCount = 0;
    const MAX_IMAGES_PER_SLIDE = 4; // A simple heuristic
    
    // Create the first slide
    if (imagePayloads.length > 0) {
        currentSlide = {
            title: 'Visuals from Source Document',
            images: [],
            isImageSlide: true // A flag for the client to recognize this slide type
        };
        slides.push(currentSlide);
    }
    
    imagePayloads.forEach(image => {
        // If the current slide is full, create a new one
        if (currentImageCount >= MAX_IMAGES_PER_SLIDE) {
            currentSlide = {
                title: 'Additional Visuals',
                images: [],
                isImageSlide: true
            };
            slides.push(currentSlide);
            currentImageCount = 0;
        }
        
        // Add the image to the current slide
        currentSlide.images.push(image);
        currentImageCount++;
    });

    return slides;
}

    // ✅ NEW METHOD: To serve the temporary images
    async serveImage(req, res) {
        const { folder, filename } = req.params;
        const filePath = path.join(__dirname, 'temp_images', folder, filename);

        try {
            // Check if file exists and is in the correct directory to prevent path traversal attacks
            if (!filePath.startsWith(path.join(__dirname, 'temp_images'))) {
                return res.status(403).json({ error: 'Forbidden' });
            }
            
            // Use stream to send the file back to the client
            res.sendFile(filePath);
        } catch (error) {
            console.error('❌ Error serving image:', error);
            res.status(404).json({ error: 'Image not found.' });
        }
    }


    async generateSlidesWithGemini(topic, sourceText) {
        const genAI = new GoogleGenerativeAI(process.env.GEMINI_API_KEY);
        const model = genAI.getGenerativeModel({ 
            model: "gemini-1.5-pro-latest",
            generationConfig: {
                temperature: 0.7,
                topP: 0.8,
                topK: 40,
                maxOutputTokens: 4096,
            }
        });

        const prompt = this.createPrompt('gemini-prompt.txt', topic, sourceText);
        
        console.log('📤 Sending request to Gemini...');
        const result = await model.generateContent(prompt);
        const response = await result.response;
        const text = await response.text();
        
        console.log('📥 Received response from Gemini');
        const cleanedText = this.cleanJson(text);
        
        try {
            return JSON.parse(cleanedText);
        } catch (parseError) {
            console.error('❌ JSON parsing failed:', parseError);
            console.error('Raw response:', text);
            throw new Error('Failed to parse AI response. Please try again.');
        }
    }

    async generateSlidesWithLMStudio(topic, sourceText) {
        console.log('📤 Sending request to LM Studio...');
        
        const prompt = this.createPrompt('lm-studio-prompt.txt', topic, sourceText);
        
        try {
            const response = await fetch('http://localhost:1234/v1/chat/completions', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                },
                body: JSON.stringify({
                    messages: [
                        { 
                            role: 'system', 
                            content: 'You are a professional presentation designer. Your task is to generate a set of slides. **You must return a valid JSON array of slides and nothing else. Do not include any conversational text, explanations, or code block formatting like ```json```.**' 
                        },
                        { role: 'user', content: prompt },
                    ],
                    temperature: 0.1,
                    max_tokens: 8192,
                    stream: false,
                }),
                timeout: 1200000 // 30 second timeout
            });

            if (!response.ok) {
                throw new Error(`LM Studio server error: ${response.status} ${response.statusText}`);
            }

            const data = await response.json();
            console.log('📥 Received response from LM Studio');
            
            if (!data.choices || !data.choices[0] || !data.choices[0].message) {
                throw new Error('Invalid response format from LM Studio');
            }

            const cleanedText = this.cleanJson(data.choices[0].message.content);
            return JSON.parse(cleanedText);
            
        } catch (error) {
            if (error.code === 'ECONNREFUSED') {
                throw new Error('LM Studio server is not running. Please start LM Studio and try again.');
            }
            throw error;
        }
    }

    createPrompt(fileName, topic, sourceText) {
        const promptTemplate = fsSync.readFileSync(path.join(__dirname, fileName), 'utf-8');
        return promptTemplate
            .replace('{topic}', topic)
            .replace('{sourceText}', sourceText ? `Base your content on this source material: ${sourceText}` : '');
    }

    async getImage(req, res) {
        const { q: query } = req.query;
        
        if (!query) {
            return res.status(400).json({ error: 'Missing query parameter.' });
        }

        if (!this.unsplashAccessKey) {
            console.warn('⚠️ Unsplash API key not configured');
            return res.status(503).json({ error: 'Image service not configured.' });
        }

        // Rate limiting
        const now = Date.now();
        const timeSinceLast = now - this.lastSearchTime;

        if (timeSinceLast < this.THROTTLE_MS) {
            const waitTime = this.THROTTLE_MS - timeSinceLast;
            return res.status(429).json({ 
                error: `Too many requests. Try again in ${Math.ceil(waitTime / 1000)} seconds.`,
                retryAfter: Math.ceil(waitTime / 1000)
            });
        }

        this.lastSearchTime = now;

        try {
            console.log(`🖼️ Searching for image: "${query}"`);
            
            const unsplashResponse = await fetch(
                `https://api.unsplash.com/search/photos?query=${encodeURIComponent(query)}&per_page=1&orientation=landscape&content_filter=high`,
                {
                    headers: {
                        Authorization: `Client-ID ${this.unsplashAccessKey}`,
                        'Accept-Version': 'v1'
                    },
                    timeout: 10000 // 10 second timeout
                }
            );

            if (!unsplashResponse.ok) {
                throw new Error(`Unsplash API error: ${unsplashResponse.status}`);
            }

            const data = await unsplashResponse.json();
            
            const imageUrl = data.results && data.results.length > 0
                ? data.results[0].urls.regular
                : null;

            if (imageUrl) {
                console.log('✅ Image found and returned');
                res.json({ 
                    imageUrl,
                    attribution: data.results[0].user ? {
                        name: data.results[0].user.name,
                        username: data.results[0].user.username,
                        link: data.results[0].user.links.html
                    } : null
                });
            } else {
                console.log('❌ No image found for query');
                res.status(404).json({ error: 'No suitable image found for this query.' });
            }
        } catch (error) {
            console.error('❌ Error fetching image from Unsplash:', error);
            res.status(500).json({ 
                error: 'Failed to fetch image.',
                details: process.env.NODE_ENV === 'development' ? error.message : undefined
            });
        }
    }

cleanJson(text) {
        // Find the index of the first valid JSON character: '{' or '['
        const jsonStart = text.indexOf('{');
        const arrayStart = text.indexOf('[');
        let start = -1;
        if (jsonStart !== -1 && (arrayStart === -1 || jsonStart < arrayStart)) {
            start = jsonStart;
        } else if (arrayStart !== -1) {
            start = arrayStart;
        }

        // Find the index of the last valid JSON character: '}' or ']'
        const jsonEnd = text.lastIndexOf('}');
        const arrayEnd = text.lastIndexOf(']');
        let end = -1;
        if (jsonEnd !== -1 && (arrayEnd === -1 || jsonEnd > arrayEnd)) {
            end = jsonEnd;
        } else if (arrayEnd !== -1) {
            end = arrayEnd;
        }

        // If a valid JSON-like structure is found
        if (start !== -1 && end !== -1 && end > start) {
            let jsonString = text.substring(start, end + 1);

            // Remove any trailing commas that might break JSON.parse
            jsonString = jsonString.replace(/,(\s*[}\]])/g, '$1');

            // Log and discard any content after the valid JSON
            const postJsonContent = text.substring(end + 1);
            if (postJsonContent.trim().length > 0) {
                console.warn("⚠️ Discarding trailing non-JSON content:", postJsonContent.trim());
            }
            
            // Remove any characters that could break the parsing
            jsonString = jsonString.replace(/(\r\n|\n|\r)/gm,"");

            return jsonString.trim();
        }

        console.error("❌ No valid JSON structure found in AI response.");
        return '';
    }

    start() {
        this.app.listen(this.port, () => {
            console.log(`🚀 Server running on http://localhost:${this.port}`);
            console.log(`📊 Health check: http://localhost:${this.port}/api/health`);
            console.log('🎨 Ready to create amazing presentations!');
        });

        // Graceful shutdown
        process.on('SIGTERM', () => {
            console.log('🛑 Received SIGTERM, shutting down gracefully...');
            process.exit(0);
        });

        process.on('SIGINT', () => {
            console.log('🛑 Received SIGINT, shutting down gracefully...');
            process.exit(0);
        });
    }
}

const app = new App();
app.start();
