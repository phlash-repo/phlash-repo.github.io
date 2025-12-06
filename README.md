# Social Story Creator 📚

Create engaging visual social stories for children using AI-powered character generation and image creation.

## 🌟 Features

- **Character Builder**: Create detailed characters with AI-generated descriptions
- **Scene Creator**: Build story scenes with customizable settings
- **AI Image Generation**: Generate illustrations using DALL-E 3
- **Story Management**: Organize scenes into complete stories
- **User-Friendly**: No technical knowledge required
- **Privacy-First**: API keys stored locally, never on servers

## 🚀 Live Demo

[View Demo](your-deployed-url-here) (Replace with your actual deployment URL)

## 🎯 Perfect For

- **Parents**: Creating personalized social stories for children
- **Teachers**: Educational storytelling and behavioral guidance
- **Therapists**: Social skills development tools
- **Special Needs**: Autism support and communication aids

## 📋 How to Use

### 1. Get Your OpenAI API Key
- Visit [platform.openai.com/api-keys](https://platform.openai.com/api-keys)
- Create a new secret key
- Enter it in the app (stored locally in your browser)

### 2. Create Characters
- Add character name, age, and gender (required)
- Leave optional fields blank for AI generation
- AI fills in appearance, clothing, and personality details

### 3. Build Scenes
- Select characters to include
- Describe the main action
- Set background and mood
- Add character expressions
- Generate images with DALL-E

### 4. Complete Your Story
- Review all scenes in the Story tab
- Download or share your finished social story

## 💰 Costs

**OpenAI API Usage** (user pays directly):
- Character generation: ~$0.01-0.05 per character
- Image generation: ~$0.04 per image
- Typical story (5 scenes): ~$0.25-0.50

## 🔧 Technical Details

### Built With
- **React 18**: Modern UI framework
- **Tailwind CSS**: Responsive styling
- **OpenAI API**: GPT-4 for text, DALL-E 3 for images
- **Vanilla Deployment**: No build process required

### Browser Support
- Chrome/Edge (latest)
- Firefox (latest)
- Safari (latest)
- Mobile responsive

## 🚀 Deploy Your Own

See [DEPLOYMENT_GUIDE.md](./DEPLOYMENT_GUIDE.md) for detailed instructions.

**Quick Options:**
- **Vercel**: Drop files and deploy (recommended)
- **Netlify**: Drag & drop deployment
- **GitHub Pages**: Free hosting option

## 🔒 Privacy & Security

- **API keys**: Stored in browser localStorage only
- **No data collection**: Stories created locally
- **User responsibility**: Each user manages their own API costs
- **No server required**: Pure client-side application

## 🎨 Customization

### Modify Character Templates
Edit the character generation prompt in the `generateCharacterDescription` function to customize AI outputs.

### Add New Frame Styles
Update the `frameStyle` options in the story settings to add new visual styles.

### Extend Scene Options
Add new camera angles, lighting moods, or art styles in the scene builder dropdowns.

## 🐛 Troubleshooting

### Common Issues

**"API Error: 401"**
- Check your OpenAI API key is correct
- Ensure you have credits in your OpenAI account

**"API Error: 429"**
- You've hit rate limits, wait a few minutes
- Consider upgrading your OpenAI plan

**Images not generating**
- Verify DALL-E 3 is available in your region
- Check your OpenAI account has image generation enabled

**Characters not generating properly**
- Ensure you filled in required fields (name, age, gender)
- Check your internet connection
- Verify GPT-4 access in your OpenAI account

## 🤝 Contributing

Want to improve the Social Story Creator? Here are ways to help:

1. **Report bugs**: Create an issue with reproduction steps
2. **Suggest features**: Ideas for new functionality
3. **Submit pull requests**: Code improvements welcome
4. **Share feedback**: How can we make it better?

### Development Setup
```bash
# Clone the repository
git clone [your-repo-url]

# Install dependencies (for React development)
npm install

# Start development server
npm start

# Build for production
npm run build
```

## 📄 License

This project is open source. Feel free to use, modify, and distribute.

## 🙏 Credits

- **OpenAI**: GPT-4 and DALL-E 3 APIs
- **React Team**: React framework
- **Tailwind CSS**: Styling framework
- **Community**: Feedback and feature requests

## 📞 Support

- **Documentation**: Check this README and deployment guide
- **Issues**: Create a GitHub issue for bugs
- **OpenAI API**: [platform.openai.com/docs](https://platform.openai.com/docs)
- **Community**: Share in relevant forums and communities

---

**Made with ❤️ for families, educators, and therapists creating meaningful stories for children.**

### Example Story Workflow

1. **Create Characters**:
   - "Emma, 5, Female" → AI generates: blonde hair, blue dress, friendly smile
   - "Sam, 7, Male" → AI generates: curly brown hair, red shirt, confident pose

2. **Build Scene 1**:
   - Title: "Emma feels nervous about her first day"
   - Action: "Emma stands at the school entrance looking worried"
   - Background: "Colorful elementary school playground"
   - Generate image

3. **Build Scene 2**:
   - Title: "Sam offers to help Emma"
   - Action: "Sam approaches Emma with a welcoming smile"
   - Characters: Both Emma and Sam
   - Generate image

4. **Complete Story**: Review all scenes, adjust as needed, share with child

Start creating meaningful social stories today! 🌈
