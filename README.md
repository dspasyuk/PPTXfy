# PPTXfy

**Automatically Generate Beautiful PowerPoint Presentations Using AI**

PPTXfy is a modern web application that transforms topics and documents into professional PowerPoint presentations using AI. Upload your documents, enter a topic, and let AI create stunning slides with automatic image integration.

## Features

- **AI-Powered Generation**: Supports multiple AI models including LM Studio and Google Gemini
- **Document Processing**: Extract content from PDF, DOC, DOCX, and TXT files
- **Image Extraction**: Automatically extracts and displays images from uploaded documents
- **Smart Image Search**: Integrates with Unsplash for relevant slide images
- **PPTX Export**: Download your presentations as PowerPoint files
- **Real-time Preview**: See your slides as they're generated

## Prerequisites

- Node.js (v16 or higher)
- LM Studio installed and running (for local AI processing)
- API keys for external services (optional):
  - Google Gemini API key
  - Unsplash API key

## Installation

1. Clone the repository:
```bash
git clone https://github.com/yourusername/pptxfy.git
cd pptxfy
```

2. Install dependencies:
```bash
npm install
```

3. Create a `.env` file in the root directory:
```env
PORT=3001
GEMINI_API_KEY=your_gemini_api_key_here
UNSPLASH_ACCESS_KEY=your_unsplash_access_key_here
NODE_ENV=development
```

## Usage

### Starting the Server

1. Start LM Studio and load your preferred model
2. Launch the PPTXfy server:
```bash
npm start
```

3. Open your browser and navigate to `http://localhost:3001`

### Creating Presentations

1. **Enter a Topic**: Describe what you want your presentation to be about
2. **Select AI Model**: Choose between LM Studio (local) or Gemini (cloud)
3. **Choose Template**: Pick from Default, Dark, Light, or Blue themes
4. **Upload Document** (optional): Add PDF, DOC, DOCX, or TXT files for content extraction
5. **Generate**: Click "Generate Slides" and watch your presentation come to life
6. **Download**: Export your presentation as a PPTX file

## Supported File Types

- **PDF**: Text and image extraction
- **DOC/DOCX**: Microsoft Word documents
- **TXT**: Plain text files

Maximum file size: 10MB

## AI Models

### LM Studio (Recommended for Privacy)
- Runs locally on your machine
- No data sent to external services
- Requires LM Studio to be running on `localhost:1234`
- Supports any compatible language model

### Google Gemini
- Cloud-based processing
- Requires API key
- Fast and reliable
- Advanced content understanding

## API Endpoints

- `POST /api/generate` - Generate slides from topic and optional file
- `GET /api/image?q={query}` - Search for images via Unsplash
- `GET /api/images/{folder}/{filename}` - Serve extracted document images
- `GET /api/health` - Check service status

## Templates

Choose from four professionally designed templates:

- **Default**: Purple gradient with modern styling
- **Dark**: Dark theme with blue accents
- **Light**: Clean light theme with blue highlights
- **Blue**: Blue-themed design with elegant layouts

## Project Structure

```
pptxfy/
├── server.js              # Main server file
├── main.js               # Frontend JavaScript
├── doc2txt.js            # Document processing utilities
├── views/
│   └── index.ejs         # Main HTML template
├── static/
│   ├── css/              # Stylesheets
│   └── js/               # Client-side libraries
├── uploads/              # Temporary file storage
├── temp_images/          # Extracted document images
├── gemini-prompt.txt     # AI prompt for Gemini
└── lm-studio-prompt.txt  # AI prompt for LM Studio
```

## Configuration

### Environment Variables

- `PORT`: Server port (default: 3001)
- `GEMINI_API_KEY`: Google Gemini API key (optional)
- `UNSPLASH_ACCESS_KEY`: Unsplash API key for image search (optional)
- `NODE_ENV`: Environment mode (development/production)

### LM Studio Setup

1. Download and install [LM Studio](https://lmstudio.ai/)
2. Load a compatible language model (recommend: Code Llama, Mistral, or similar)
3. Start the local server on port 1234
4. Ensure the model supports JSON output formatting

## Error Handling

PPTXfy includes comprehensive error handling for:

- Invalid file types or sizes
- AI service connectivity issues
- API rate limiting
- Malformed AI responses
- Network timeouts

## Security Features

- File type validation
- Path traversal protection
- CORS configuration
- Input sanitization
- File size limits

## Performance Optimizations

- Image compression and resizing
- Intelligent slide layouts
- Efficient document processing
- Rate limiting for external APIs
- Memory management for large files

## Troubleshooting

### Common Issues

**LM Studio Connection Error**
```
Error: LM Studio server is not running
```
- Ensure LM Studio is running on localhost:1234
- Check that a model is loaded and ready
- Verify no firewall is blocking the connection

**File Upload Issues**
```
Error: Invalid file type
```
- Only PDF, DOC, DOCX, and TXT files are supported
- Maximum file size is 10MB
- Ensure file is not corrupted

**Image Search Not Working**
```
Error: Image service not configured
```
- Add your Unsplash API key to the .env file
- Verify the API key is valid and active

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
<img width="2437" height="1313" alt="image" src="https://github.com/user-attachments/assets/c79bca88-423f-4574-94be-a989c1bcfccd" />
<img width="2437" height="1313" alt="image" src="https://github.com/user-attachments/assets/4e571b81-3f15-45df-9572-0504da263cde" />


