# Social Story Creator - Deployment Guide

## 🚀 Quick Deploy Options (Recommended)

### Option 1: Vercel (Easiest - Free)
1. Go to [vercel.com](https://vercel.com) and sign up with GitHub
2. Upload all files to a new GitHub repository
3. Connect your repo to Vercel
4. Deploy automatically - your app will be live in minutes!
5. **Custom domain**: Free with Vercel (yourapp.vercel.app)

### Option 2: Netlify (Also Easy - Free)
1. Go to [netlify.com](https://netlify.com) and sign up
2. Drag & drop the files directly to Netlify's deploy area, OR
3. Connect your GitHub repository
4. Automatic deployments on every push
5. **Custom domain**: Free with Netlify (yourapp.netlify.app)

### Option 3: GitHub Pages (Free)
1. Upload files to a GitHub repository
2. Go to Settings > Pages
3. Select source as "Deploy from a branch"
4. Select your main branch
5. Your app will be available at `username.github.io/repository-name`

## 📁 Files You Need to Upload

For **simple deployment** (HTML only):
- `index.html` (main app file)
- `vercel.json` (for Vercel)
- `netlify.toml` (for Netlify)

For **full React development**:
- All files above PLUS:
- `package.json`
- `social-story-creator.jsx` (React component)

## 🔑 OpenAI API Key Setup

### For Users:
1. Get API key at [platform.openai.com/api-keys](https://platform.openai.com/api-keys)
2. Create new secret key
3. Enter key in the app (stored locally in browser)
4. **Cost**: ~$0.01-0.05 per character generation, ~$0.04 per image

### Security Notes:
- API keys are stored in user's localStorage (browser only)
- Keys never sent to your servers
- Each user pays for their own API usage
- Consider adding usage warnings/limits

## 🛠 Advanced Setup (Optional)

### Full React Development Environment:
```bash
# If you want to develop locally
npm install
npm start

# Build for production
npm run build
```

### Custom Domain:
1. **Vercel**: Project Settings > Domains
2. **Netlify**: Site Settings > Domain Management
3. **GitHub Pages**: Repository Settings > Pages > Custom Domain

### Environment Variables (if needed):
- Create `.env` file for any server-side configs
- For client-side: Use React's `REACT_APP_` prefix

## 📊 Monitoring & Analytics

### Add Google Analytics (Optional):
Add to `<head>` of index.html:
```html
<!-- Google tag (gtag.js) -->
<script async src="https://www.googletagmanager.com/gtag/js?id=GA_TRACKING_ID"></script>
<script>
  window.dataLayer = window.dataLayer || [];
  function gtag(){dataLayer.push(arguments);}
  gtag('js', new Date());
  gtag('config', 'GA_TRACKING_ID');
</script>
```

## 🔒 Security Considerations

1. **API Key Safety**:
   - Keys stored client-side only
   - No server-side key storage needed
   - Users responsible for their own keys

2. **CORS Issues**:
   - OpenAI API supports CORS for direct calls
   - No proxy server needed

3. **Rate Limiting**:
   - OpenAI has built-in rate limits
   - Consider adding user-facing warnings

## 🎯 Success Metrics

After deployment, monitor:
- User signups/API key additions
- Image generations per day
- Character creations
- Story completions

## 🚨 Cost Estimates for Users

**OpenAI API Pricing** (as of 2024):
- **Character Generation** (GPT-4): ~$0.01-0.05 per character
- **Image Generation** (DALL-E 3): ~$0.04 per image
- **Monthly estimate**: $5-20 for active users

## 🆘 Troubleshooting

### Common Issues:
1. **API Key errors**: Check key format (starts with 'sk-')
2. **CORS errors**: Ensure using HTTPS in production
3. **Image generation fails**: Check OpenAI account credits
4. **Slow loading**: Consider CDN for static assets

### Support Resources:
- OpenAI API docs: [platform.openai.com/docs](https://platform.openai.com/docs)
- React docs: [reactjs.org](https://reactjs.org)
- Deployment platform specific support

## 🎉 You're Ready!

Choose your deployment method and your Social Story Creator will be live for the world to use! 

**Recommended for beginners**: Start with Vercel - it's the fastest way to get online with zero configuration.
