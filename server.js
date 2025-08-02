const express = require('express');
const cors = require('cors');
const fetch = require('node-fetch');
const path = require('path');
require('dotenv').config();

const app = express();
const PORT = process.env.PORT || 8080;

// Middleware
app.use(cors());
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public'), {
  setHeaders: function (res, filePath) {
    if (filePath.endsWith('.css')) {
      res.setHeader('Content-Type', 'text/css');
    }
  }
}));

// OpenAI API configuration
const OPENAI_API_KEY = process.env.OPENAI_API_KEY;
const OPENAI_API_URL = 'https://api.openai.com/v1/chat/completions';

// Health check endpoint
app.get('/health', (req, res) => {
    res.json({ status: 'ok', message: 'XSLGPT server is running' });
});

// Main formula generation endpoint
app.post('/api/generate', async (req, res) => {
    try {
        const { prompt } = req.body;

        // Validate input
        if (!prompt || typeof prompt !== 'string' || prompt.trim().length === 0) {
            return res.status(400).json({
                success: false,
                error: 'Prompt is required and must be a non-empty string'
            });
        }

        // Check if API key is configured
        if (!OPENAI_API_KEY) {
            console.error('OpenAI API key not configured');
            return res.status(500).json({
                success: false,
                error: 'Server configuration error: OpenAI API key not set'
            });
        }

        // Call OpenAI API
        const openaiResponse = await fetch(OPENAI_API_URL, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'Authorization': `Bearer ${OPENAI_API_KEY}`
            },
            body: JSON.stringify({
                model: 'gpt-3.5-turbo',
                messages: [
                    {
                        role: 'system',
                        content: `You are an Excel formula expert. When asked for an Excel formula, respond with ONLY:
1. A valid Excel formula that starts with "="
2. A brief explanation (1-2 sentences) on a new line

Example response:
=SUM(A1:A10)
This formula adds up all values in cells A1 through A10.

Always ensure the formula is valid Excel syntax and starts with "=".`
                    },
                    {
                        role: 'user',
                        content: prompt.trim()
                    }
                ],
                max_tokens: 200,
                temperature: 0.3
            })
        });

        if (!openaiResponse.ok) {
            const errorData = await openaiResponse.json();
            console.error('OpenAI API error:', errorData);
            return res.status(openaiResponse.status).json({
                success: false,
                error: errorData.error?.message || `OpenAI API error: ${openaiResponse.status}`
            });
        }

        const data = await openaiResponse.json();
        const content = data.choices[0].message.content.trim();

        // Parse the response to extract formula and explanation
        const lines = content.split('\n');
        const formula = lines[0];
        const explanation = lines.slice(1).join('\n').trim();

        // Validate that the response starts with "="
        if (!formula.startsWith('=')) {
            return res.status(400).json({
                success: false,
                error: 'Generated response does not contain a valid Excel formula'
            });
        }

        // Log successful generation (without sensitive data)
        console.log(`Formula generated successfully for prompt: "${prompt.substring(0, 50)}..."`);

        res.json({
            success: true,
            formula: formula,
            explanation: explanation
        });

    } catch (error) {
        console.error('Server error:', error);
        res.status(500).json({
            success: false,
            error: 'Internal server error'
        });
    }
});

// Serve the manifest file
app.get('/manifest.xml', (req, res) => {
    res.sendFile(path.join(__dirname, 'manifest.xml'));
});

// Handle favicon requests
app.get('/favicon.ico', (req, res) => {
    res.setHeader('Content-Type', 'image/x-icon');
    res.status(200).send(''); // Empty favicon
});

// Default route
app.get('/', (req, res) => {
    res.json({
        message: 'XSLGPT Excel Add-in Server',
        version: '1.0.0',
        endpoints: {
            health: '/health',
            generate: '/api/generate'
        }
    });
});

// Start server
app.listen(PORT, () => {
    console.log(`🚀 XSLGPT server running on port ${PORT}`);
    console.log(`📊 Health check: http://localhost:${PORT}/health`);
    console.log(`🔧 API endpoint: http://localhost:${PORT}/api/generate`);
    
    if (!OPENAI_API_KEY) {
        console.warn('⚠️  WARNING: OPENAI_API_KEY not set in environment variables');
        console.warn('   Please set OPENAI_API_KEY in your .env file or environment');
    } else {
        console.log('✅ OpenAI API key configured');
    }
});

module.exports = app; 