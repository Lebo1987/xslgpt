# 🚀 XSLGPT Excel Add-in - Deployment Guide

This guide will help you deploy the XSLGPT Excel Add-in to any HTTPS web server for independent distribution.

## 📋 Prerequisites

- A web server with HTTPS support
- Node.js (for the backend server)
- An OpenAI API key
- Domain name (e.g., `yourdomain.com`)

## 🗂️ File Structure

```
yourdomain.com/
├── xslgpt/
│   ├── taskpane.html          # Main add-in interface
│   ├── taskpane.js            # Add-in logic
│   ├── taskpane.css           # Styles
│   ├── config.js              # Configuration
│   ├── manifest-production.xml # Add-in manifest
│   └── assets/                # Icons (optional)
│       ├── icon-16.png
│       ├── icon-32.png
│       └── icon-80.png
├── api/                       # Backend API
│   ├── server-production.js
│   ├── package.json
│   └── .env
└── download.html              # Download page
```

## 🔧 Step-by-Step Deployment

### 1. Prepare Your Domain

Replace all instances of `yourdomain.com` in the following files with your actual domain:

- `manifest-production.xml`
- `server-production.js`
- `download-page.html`

### 2. Set Up the Backend Server

#### Option A: Deploy to a Node.js Hosting Service (Recommended)

1. **Create a new directory for the backend:**
   ```bash
   mkdir xslgpt-api
   cd xslgpt-api
   ```

2. **Copy the backend files:**
   ```bash
   cp server-production.js server.js
   cp package.json .
   cp .env.example .env
   ```

3. **Install dependencies:**
   ```bash
   npm install
   ```

4. **Configure environment variables:**
   ```bash
   # Edit .env file
   OPENAI_API_KEY=your_openai_api_key_here
   PORT=3000
   ```

5. **Deploy to your preferred hosting service:**
   - **Vercel:** `vercel --prod`
   - **Netlify:** `netlify deploy --prod`
   - **Railway:** `railway up`
   - **Heroku:** `git push heroku main`

#### Option B: Deploy to a VPS/Cloud Server

1. **Upload files to your server:**
   ```bash
   scp -r xslgpt-api/ user@your-server:/var/www/xslgpt-api/
   ```

2. **Install dependencies:**
   ```bash
   ssh user@your-server
   cd /var/www/xslgpt-api
   npm install
   ```

3. **Set up environment variables:**
   ```bash
   nano .env
   # Add your OpenAI API key
   ```

4. **Set up PM2 for process management:**
   ```bash
   npm install -g pm2
   pm2 start server.js --name xslgpt-api
   pm2 startup
   pm2 save
   ```

5. **Configure Nginx reverse proxy:**
   ```nginx
   server {
       listen 443 ssl;
       server_name yourdomain.com;
       
       ssl_certificate /path/to/your/certificate.crt;
       ssl_certificate_key /path/to/your/private.key;
       
       location /api/ {
           proxy_pass http://localhost:3000;
           proxy_http_version 1.1;
           proxy_set_header Upgrade $http_upgrade;
           proxy_set_header Connection 'upgrade';
           proxy_set_header Host $host;
           proxy_set_header X-Real-IP $remote_addr;
           proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
           proxy_set_header X-Forwarded-Proto $scheme;
           proxy_cache_bypass $http_upgrade;
       }
       
       location / {
           root /var/www/xslgpt;
           try_files $uri $uri/ =404;
       }
   }
   ```

### 3. Deploy Frontend Files

1. **Create the frontend directory on your web server:**
   ```bash
   mkdir -p /var/www/xslgpt
   ```

2. **Upload frontend files:**
   ```bash
   scp public/taskpane.html user@your-server:/var/www/xslgpt/
   scp public/taskpane.js user@your-server:/var/www/xslgpt/
   scp public/taskpane.css user@your-server:/var/www/xslgpt/
   scp public/config.js user@your-server:/var/www/xslgpt/
   scp manifest-production.xml user@your-server:/var/www/xslgpt/manifest.xml
   scp download-page.html user@your-server:/var/www/xslgpt/index.html
   ```

3. **Set proper permissions:**
   ```bash
   ssh user@your-server
   chmod 644 /var/www/xslgpt/*
   ```

### 4. Configure Web Server

#### For Apache:
```apache
<VirtualHost *:443>
    ServerName yourdomain.com
    DocumentRoot /var/www/xslgpt
    
    SSLEngine on
    SSLCertificateFile /path/to/your/certificate.crt
    SSLCertificateKeyFile /path/to/your/private.key
    
    <Directory /var/www/xslgpt>
        AllowOverride All
        Require all granted
    </Directory>
    
    # Handle CORS
    Header always set Access-Control-Allow-Origin "*"
    Header always set Access-Control-Allow-Methods "GET, POST, OPTIONS"
    Header always set Access-Control-Allow-Headers "Content-Type"
</VirtualHost>
```

#### For Nginx:
```nginx
server {
    listen 443 ssl http2;
    server_name yourdomain.com;
    root /var/www/xslgpt;
    index index.html;
    
    ssl_certificate /path/to/your/certificate.crt;
    ssl_certificate_key /path/to/your/private.key;
    
    # Handle CORS
    add_header Access-Control-Allow-Origin "*" always;
    add_header Access-Control-Allow-Methods "GET, POST, OPTIONS" always;
    add_header Access-Control-Allow-Headers "Content-Type" always;
    
    # Handle preflight requests
    if ($request_method = 'OPTIONS') {
        add_header Access-Control-Allow-Origin "*";
        add_header Access-Control-Allow-Methods "GET, POST, OPTIONS";
        add_header Access-Control-Allow-Headers "Content-Type";
        add_header Content-Length 0;
        add_header Content-Type text/plain;
        return 204;
    }
    
    location / {
        try_files $uri $uri/ =404;
    }
    
    # Serve static files with proper MIME types
    location ~* \.(css|js|html)$ {
        expires 1y;
        add_header Cache-Control "public, immutable";
    }
}
```

### 5. Test Your Deployment

1. **Test the backend API:**
   ```bash
   curl https://yourdomain.com/api/health
   ```

2. **Test the frontend:**
   - Visit `https://yourdomain.com/xslgpt/`
   - Should show the download page

3. **Test the manifest:**
   - Visit `https://yourdomain.com/xslgpt/manifest.xml`
   - Should download the manifest file

### 6. Update URLs in Files

Before deploying, update these URLs in your files:

#### In `manifest-production.xml`:
```xml
<SourceLocation DefaultValue="https://yourdomain.com/xslgpt/taskpane.html"/>
<bt:Url id="Taskpane.Url" DefaultValue="https://yourdomain.com/xslgpt/taskpane.html"/>
```

#### In `server-production.js`:
```javascript
app.use(cors({
    origin: ['https://yourdomain.com', 'https://www.yourdomain.com'],
    credentials: true
}));
```

## 🔒 Security Considerations

1. **HTTPS Required:** Excel add-ins require HTTPS for production use
2. **CORS Configuration:** Ensure proper CORS headers are set
3. **API Key Security:** Keep your OpenAI API key secure and never expose it in client-side code
4. **Rate Limiting:** Consider implementing rate limiting on your API endpoints

## 📊 Monitoring

1. **Set up logging:**
   ```bash
   pm2 logs xslgpt-api
   ```

2. **Monitor API usage:**
   ```bash
   curl https://yourdomain.com/api/health
   ```

3. **Check server status:**
   ```bash
   pm2 status
   ```

## 🚨 Troubleshooting

### Common Issues:

1. **CORS Errors:**
   - Ensure CORS headers are properly configured
   - Check that your domain is in the allowed origins

2. **HTTPS Issues:**
   - Verify SSL certificate is valid
   - Check that all URLs use HTTPS

3. **API Connection Errors:**
   - Verify OpenAI API key is correct
   - Check network connectivity
   - Review server logs for errors

4. **Add-in Not Loading:**
   - Verify manifest.xml is accessible
   - Check that all frontend files are served correctly
   - Ensure MIME types are set properly

## 📞 Support

If you encounter issues:

1. Check the server logs: `pm2 logs xslgpt-api`
2. Verify API connectivity: `curl https://yourdomain.com/api/health`
3. Test frontend files: Visit each URL directly in browser
4. Check Excel add-in debugging tools for client-side errors

## ✅ Deployment Checklist

- [ ] Domain configured with HTTPS
- [ ] Backend API deployed and accessible
- [ ] Frontend files uploaded to web server
- [ ] Manifest file accessible at `https://yourdomain.com/xslgpt/manifest.xml`
- [ ] Download page working at `https://yourdomain.com/xslgpt/`
- [ ] CORS headers properly configured
- [ ] OpenAI API key configured and working
- [ ] All URLs updated to use your domain
- [ ] SSL certificate valid and trusted
- [ ] Server monitoring and logging set up

Your XSLGPT Excel Add-in is now ready for production distribution! 🎉 