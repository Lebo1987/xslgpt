# XSLGPT - Excel Formula Assistant

An AI-powered Excel Add-in that generates Excel formulas using natural language prompts via OpenAI's GPT-3.5-Turbo.

## 🏗️ Architecture

This project uses a **backend proxy server** architecture for enhanced security and ease of deployment:

```
Excel Add-in UI ↔ Backend Server ↔ OpenAI API
```

**Benefits:**
- ✅ **No API keys for end users** - Everything is handled securely on the server
- ✅ **Centralized management** - Single API key for all users
- ✅ **Better security** - API keys never leave your server
- ✅ **Easy deployment** - One server serves all users

## 🚀 Quick Start

### Prerequisites

1. **Node.js** (version 14 or higher)
2. **OpenAI API Key** - Get one from [OpenAI Platform](https://platform.openai.com/api-keys)
3. **Excel Desktop Application** (Windows or Mac)

### Installation

1. **Clone or download this project**
2. **Install dependencies**:
   ```bash
   npm install
   ```

3. **Configure your OpenAI API key**:
   ```bash
   # Copy the example environment file
   cp env.example .env
   
   # Edit .env and add your OpenAI API key
   OPENAI_API_KEY=sk-your-actual-api-key-here
   ```

4. **Start the server**:
   ```bash
   npm start
   ```

5. **Load the add-in in Excel**:
   - Open Excel
   - Go to **Insert > Add-ins > My Add-ins**
   - Click **"Upload My Add-in"**
   - Select the `manifest.xml` file from this project

## 📁 Project Structure

```
xslgpt-excel/
├── server.js              # Backend Express server
├── package.json           # Node.js dependencies
├── env.example           # Environment variables template
├── manifest.xml          # Office Add-in manifest
├── taskpane.html         # Main HTML interface
├── taskpane.js           # Frontend JavaScript
├── config.js             # Configuration management
├── styles.css            # Styling
└── README.md            # This file
```

## 🔧 Development

### Local Development

```bash
# Install dependencies
npm install

# Start development server with auto-reload
npm run dev

# Start production server
npm start
```

### Environment Variables

Create a `.env` file based on `env.example`:

```env
# Required: Your OpenAI API key
OPENAI_API_KEY=sk-your-openai-api-key-here

# Optional: Server port (defaults to 3000)
PORT=3000

# Optional: Environment
NODE_ENV=development
```

## 🌐 Deployment

### Option 1: Heroku

1. **Create a Heroku app**:
   ```bash
   heroku create your-xslgpt-app
   ```

2. **Set environment variables**:
   ```bash
   heroku config:set OPENAI_API_KEY=sk-your-api-key
   ```

3. **Deploy**:
   ```bash
   git push heroku main
   ```

4. **Update manifest.xml** with your Heroku URL:
   ```xml
   <SourceLocation DefaultValue="https://your-app.herokuapp.com/taskpane.html"/>
   ```

### Option 2: Vercel

1. **Install Vercel CLI**:
   ```bash
   npm i -g vercel
   ```

2. **Deploy**:
   ```bash
   vercel
   ```

3. **Set environment variables** in Vercel dashboard

### Option 3: Railway

1. **Connect your GitHub repo** to Railway
2. **Set environment variables** in Railway dashboard
3. **Deploy automatically** on git push

### Option 4: DigitalOcean App Platform

1. **Connect your GitHub repo** to DigitalOcean
2. **Set environment variables** in the dashboard
3. **Deploy with one click**

## 📖 Usage

1. **Open the XSLGPT add-in** from the Excel ribbon
2. **Type your question** in the input box (e.g., "Calculate the sum of cells A1 to A10")
3. **Click "Generate Formula"** or press Ctrl+Enter
4. **The formula will be automatically inserted** into your currently selected cell
5. **View your prompt history** below the input box

## 💡 Example Prompts

- "Calculate the average of values in column B"
- "Find the maximum value in range A1:D10"
- "Count how many cells in column C contain text"
- "Calculate the percentage of A1 divided by B1"
- "Create a formula to round the value in cell A1 to 2 decimal places"
- "Calculate the date difference between A1 and B1"

## 🔒 Security Features

- **API Key Protection**: OpenAI API key is stored only on the server
- **Input Validation**: All user inputs are validated before processing
- **Error Handling**: Comprehensive error handling and user feedback
- **CORS Protection**: Proper CORS configuration for security
- **Rate Limiting**: Built-in protection against abuse (can be enhanced)

## 🛠️ Customization

### Backend Customization

- **Modify AI prompts**: Edit the system message in `server.js`
- **Add rate limiting**: Install and configure `express-rate-limit`
- **Add authentication**: Implement user authentication if needed
- **Add logging**: Integrate with logging services like Winston

### Frontend Customization

- **Change styling**: Modify `styles.css`
- **Add features**: Extend `taskpane.js`
- **Update branding**: Modify `manifest.xml` and HTML

## 🐛 Troubleshooting

### Common Issues

1. **"Server not starting"**
   - Check if port 3000 is available
   - Verify Node.js version (14+)
   - Ensure all dependencies are installed

2. **"API key not working"**
   - Verify your OpenAI API key is correct
   - Check that you have sufficient credits
   - Review server logs for errors

3. **"Add-in not loading"**
   - Ensure server is running on correct port
   - Check manifest.xml URLs match your server
   - Try accessing `http://localhost:3000/health` in browser

4. **"Formula not inserting"**
   - Make sure you have a cell selected in Excel
   - Check browser console for errors
   - Verify add-in permissions

### Debug Mode

```bash
# Start server with debug logging
DEBUG=* npm start
```

## 📊 API Endpoints

- `GET /health` - Server health check
- `POST /api/generate` - Generate Excel formula
- `GET /` - Server info
- `GET /manifest.xml` - Excel Add-in manifest

## 🔄 Updates

To update the add-in for all users:

1. **Deploy server changes** to your hosting platform
2. **Update manifest.xml** if server URL changes
3. **Users may need to reload** the add-in in Excel

## 📞 Support

For issues or questions:
1. Check the troubleshooting section above
2. Review server logs for errors
3. Test the `/health` endpoint
4. Verify OpenAI API key and credits

## 📄 License

This project is provided as-is for educational and commercial use.

---

**Note**: This add-in requires an active internet connection to communicate with the backend server. Make sure your server is accessible from your users' networks. 