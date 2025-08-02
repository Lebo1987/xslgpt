# 🚀 XSLGPT Excel Add-in - Production Distribution Package

A complete, production-ready Excel Add-in that uses AI to generate Excel formulas from natural language descriptions.

## 📦 What's Included

This package contains everything needed to deploy XSLGPT as an independent Excel Add-in:

### Frontend Files (`public/` folder)
- `taskpane.html` - Main add-in interface
- `taskpane.js` - Add-in logic and Excel integration
- `taskpane.css` - Modern, responsive styling
- `config.js` - Configuration management

### Backend Files
- `server-production.js` - Production-ready Node.js server
- `package-production.json` - Dependencies for backend
- `env.production` - Environment variables template

### Distribution Files
- `manifest-production.xml` - Excel Add-in manifest
- `download-page.html` - Professional download page
- `DEPLOYMENT.md` - Complete deployment guide

## 🎯 Features

- **🤖 AI-Powered:** Uses OpenAI GPT-3.5 to understand natural language
- **⚡ Instant Results:** Generate Excel formulas instantly with explanations
- **📊 Smart Insertion:** Automatically inserts formulas into selected cells
- **📅 Date Formatting:** Automatically applies proper date formatting for date formulas
- **📝 History Tracking:** Keeps track of previous prompts and formulas
- **🎨 Modern UI:** Clean, professional interface with responsive design
- **🔒 Secure:** API key stored securely on backend only

## 🚀 Quick Start

### For End Users (Excel Users)

1. **Download the manifest file** from the download page
2. **Open Excel** and go to Insert → Add-ins → Upload My Add-in
3. **Select the manifest file** and click Upload
4. **Start using XSLGPT** by clicking the XSLGPT button in the Home tab

### For Developers/Deployers

1. **Follow the deployment guide** in `DEPLOYMENT.md`
2. **Update domain URLs** in the manifest and server files
3. **Deploy backend** to your preferred hosting service
4. **Upload frontend files** to your web server
5. **Test the deployment** using the provided checklist

## 📋 Prerequisites

### For Deployment
- Web server with HTTPS support
- Node.js (v14 or higher)
- OpenAI API key
- Domain name

### For End Users
- Microsoft Excel (Desktop or Online)
- Internet connection
- No technical knowledge required

## 🔧 Installation Instructions

### Step 1: Download
Visit the download page and click "Download Manifest File"

### Step 2: Install in Excel
1. Open Microsoft Excel
2. Go to the **Insert** tab
3. Click **Add-ins** → **Upload My Add-in**
4. Browse to the downloaded manifest file
5. Click **Upload**

### Step 3: Use XSLGPT
1. Look for the **XSLGPT** button in the **Home** tab
2. Click it to open the formula assistant
3. Type your request in natural language
4. Click "Generate Formula"
5. Click "Insert Formula" to add it to your selected cell

## 💡 Usage Examples

| Request | Generated Formula | Explanation |
|---------|------------------|-------------|
| "Calculate the sum of cells A1 to A10" | `=SUM(A1:A10)` | Adds all values in the specified range |
| "Show today's date" | `=TODAY()` | Returns today's date with proper formatting |
| "Find the average of values in column B" | `=AVERAGE(B:B)` | Calculates the average of all values in column B |
| "Count how many cells in A1:A20 contain text" | `=COUNTIF(A1:A20,"*")` | Counts cells containing any text |

## 🏗️ Architecture

```
Excel Add-in (Frontend)
    ↓ HTTPS
Web Server (Static Files)
    ↓ HTTPS
Backend API (Node.js)
    ↓ HTTPS
OpenAI API (GPT-3.5)
```

## 🔒 Security Features

- **HTTPS Required:** All communications use secure HTTPS
- **API Key Protection:** OpenAI API key stored only on backend
- **CORS Configuration:** Proper cross-origin resource sharing setup
- **Input Validation:** All user inputs are validated and sanitized
- **Error Handling:** Comprehensive error handling and logging

## 📊 Performance

- **Fast Response:** Average response time < 2 seconds
- **Caching:** Static files cached for optimal performance
- **Compression:** Gzip compression enabled for faster loading
- **CDN Ready:** Files structured for easy CDN integration

## 🛠️ Customization

### Changing the Domain
Update these files with your domain:
- `manifest-production.xml` - Replace `yourdomain.com` with your domain
- `server-production.js` - Update CORS origins
- `download-page.html` - Update download links

### Styling
- Modify `public/taskpane.css` to change the appearance
- Update colors, fonts, and layout as needed

### Backend Configuration
- Edit `env.production` to configure environment variables
- Modify `server-production.js` for custom API behavior

## 🚨 Troubleshooting

### Common Issues

**Add-in not loading:**
- Verify HTTPS is enabled
- Check that manifest.xml is accessible
- Ensure all frontend files are served correctly

**API errors:**
- Verify OpenAI API key is correct
- Check network connectivity
- Review server logs for errors

**CORS errors:**
- Ensure CORS headers are properly configured
- Verify domain is in allowed origins list

### Getting Help

1. Check the server logs for error messages
2. Verify API connectivity with health check
3. Test frontend files by visiting URLs directly
4. Use Excel's add-in debugging tools

## 📞 Support

For deployment issues:
1. Review `DEPLOYMENT.md` for detailed instructions
2. Check the troubleshooting section above
3. Verify all prerequisites are met

For end user issues:
1. Ensure Excel is up to date
2. Check internet connectivity
3. Try reinstalling the add-in

## 📄 License

This project is licensed under the MIT License - see the LICENSE file for details.

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Test thoroughly
5. Submit a pull request

## 📈 Roadmap

- [ ] Support for more Excel functions
- [ ] Formula optimization suggestions
- [ ] Multi-language support
- [ ] Advanced formatting options
- [ ] Integration with Excel templates

---

**Made with ❤️ for Excel users everywhere**

*Powered by OpenAI GPT-3.5* 