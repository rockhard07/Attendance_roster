# Deployment Guide for Streamlit Cloud

## Quick Start

### Option 1: Deploy to Streamlit Cloud (Recommended)

#### Prerequisites
- GitHub account
- Streamlit account (free tier available at https://share.streamlit.io)

#### Steps:
1. Push code to GitHub
2. Go to https://share.streamlit.io
3. Click "New app"
4. Connect your GitHub repository
5. Select branch: `main`
6. Set main file path: `app.py`
7. Click "Deploy"

**Your app will be live at**: `https://<your-username>-<repo-name>.streamlit.app`

### Option 2: Deploy to Heroku

#### Prerequisites
- Heroku account
- Heroku CLI installed

#### Steps:
```bash
heroku login
heroku create your-app-name
git push heroku main
heroku open
```

### Option 3: Deploy to AWS/Azure/Google Cloud

Each cloud platform has different deployment methods. Contact support for platform-specific guidance.

## Environment Setup

### Local Testing
```bash
pip install -r requirements.txt
streamlit run app.py
```

### Streamlit Cloud Configuration
The app uses the `.streamlit/config.toml` file for configuration. No environment variables needed for basic deployment.

## File Structure for Deployment
```
git/
├── app.py
├── pdf_extractor.py
├── requirements.txt
├── README.md
├── DEPLOYMENT.md (this file)
├── .streamlit/
│   └── config.toml
└── .gitignore
```

## Key Points

✅ All dependencies listed in `requirements.txt`  
✅ No external API keys required  
✅ No database needed  
✅ Streamlit Cloud handles Python version management  
✅ Max file upload: 200MB (configurable)  

## Troubleshooting

### App won't deploy
- Check that `app.py` is in the root of the selected directory
- Ensure all imports are in `requirements.txt`
- Check Streamlit Cloud logs for errors

### Slow performance
- Streamlit Cloud free tier has limited resources
- Consider upgrading to paid tier for better performance
- Optimize PDF processing for large files

### File size limits
- Default: 200MB per file
- Configure in `.streamlit/config.toml` with `maxUploadSize`

## Updates & Maintenance

To update the deployed app:
1. Make changes locally
2. Test with `streamlit run app.py`
3. Commit and push to GitHub
4. Streamlit Cloud auto-redeploys on push (if connected)

## Support

For issues:
- Check Streamlit documentation: https://docs.streamlit.io
- Review logs in Streamlit Cloud dashboard
- Consult Deployment.md in the main folder

---

**Version**: 1.0  
**Last Updated**: January 2026
