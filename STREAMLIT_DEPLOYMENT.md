# 🚀 Streamlit Deployment Guide

This guide will help you deploy your CostHead Report Generator as a Streamlit web application.

## 📋 Prerequisites

- Python 3.8 or higher
- pip (Python package installer)
- Git (for version control)

## 🛠️ Local Development Setup

### 1. Install Dependencies

```bash
# Install required packages
pip install -r requirements.txt

# Or install Streamlit directly
pip install streamlit pandas openpyxl python-docx reportlab
```

### 2. Run the Application Locally

```bash
# Method 1: Using the launch script
python run_streamlit.py

# Method 2: Direct Streamlit command
streamlit run streamlit_app.py

# Method 3: With custom port
streamlit run streamlit_app.py --server.port 8502
```

### 3. Access the Application

- Open your web browser
- Navigate to: `http://localhost:8501`
- The app will automatically reload when you make changes to the code

## 🌐 Cloud Deployment Options

### Option 1: Streamlit Cloud (Recommended - Free)

1. **Push to GitHub:**
   ```bash
   git init
   git add .
   git commit -m "Initial commit"
   git branch -M main
   git remote add origin https://github.com/yourusername/your-repo.git
   git push -u origin main
   ```

2. **Deploy on Streamlit Cloud:**
   - Go to [share.streamlit.io](https://share.streamlit.io)
   - Sign in with GitHub
   - Click "New app"
   - Select your repository
   - Set the main file path to `streamlit_app.py`
   - Click "Deploy"

### Option 2: Heroku

1. **Create a Procfile:**
   ```bash
   echo "web: streamlit run streamlit_app.py --server.port=$PORT --server.address=0.0.0.0" > Procfile
   ```

2. **Create runtime.txt:**
   ```bash
   echo "python-3.11.0" > runtime.txt
   ```

3. **Deploy to Heroku:**
   ```bash
   heroku create your-app-name
   git push heroku main
   ```

### Option 3: AWS EC2

1. **Launch EC2 instance**
2. **Install dependencies:**
   ```bash
   sudo apt update
   sudo apt install python3-pip
   pip3 install -r requirements.txt
   ```

3. **Run with nohup:**
   ```bash
   nohup streamlit run streamlit_app.py --server.port 80 --server.address 0.0.0.0 &
   ```

### Option 4: Docker Deployment

1. **Create Dockerfile:**
   ```dockerfile
   FROM python:3.11-slim

   WORKDIR /app
   COPY requirements.txt .
   RUN pip install -r requirements.txt

   COPY . .

   EXPOSE 8501

   CMD ["streamlit", "run", "streamlit_app.py", "--server.address", "0.0.0.0"]
   ```

2. **Build and run:**
   ```bash
   docker build -t costhead-app .
   docker run -p 8501:8501 costhead-app
   ```

## 🔧 Configuration Options

### Environment Variables

Create a `.env` file for configuration:

```bash
# .env
STREAMLIT_SERVER_PORT=8501
STREAMLIT_SERVER_ADDRESS=0.0.0.0
STREAMLIT_BROWSER_GATHER_USAGE_STATS=false
```

### Streamlit Configuration

Create `.streamlit/config.toml`:

```toml
[server]
port = 8501
address = "0.0.0.0"
maxUploadSize = 200

[browser]
gatherUsageStats = false

[theme]
primaryColor = "#1f538d"
backgroundColor = "#ffffff"
secondaryBackgroundColor = "#f0f2f6"
textColor = "#262730"
```

## 📊 Features of the Streamlit App

### ✅ What's Included:
- **File Upload Interface**: Drag-and-drop Excel file uploads
- **Data Preview**: Real-time preview of uploaded data
- **Auto Column Detection**: Automatically detects required columns
- **Report Generation**: Full CostHead report generation
- **Download Functionality**: ZIP file download with all reports
- **Progress Indicators**: Loading spinners and progress bars
- **Error Handling**: Comprehensive error messages
- **Responsive Design**: Works on desktop and mobile

### 🎨 UI Features:
- Modern, clean interface
- Sidebar navigation
- Tabbed data preview
- Status indicators
- Success/error notifications
- Download buttons

## 🚀 Quick Start Commands

```bash
# Clone and setup
git clone <your-repo-url>
cd excel-project

# Install dependencies
pip install -r requirements.txt

# Run locally
python run_streamlit.py

# Or run directly
streamlit run streamlit_app.py
```

## 🔍 Troubleshooting

### Common Issues:

1. **Port already in use:**
   ```bash
   streamlit run streamlit_app.py --server.port 8502
   ```

2. **Permission denied:**
   ```bash
   chmod +x run_streamlit.py
   ```

3. **Module not found:**
   ```bash
   pip install -r requirements.txt
   ```

4. **File upload issues:**
   - Check file size limits
   - Ensure Excel files are not corrupted
   - Try with smaller test files first

### Performance Tips:

- Use smaller Excel files for testing
- Clear browser cache if issues persist
- Check server logs for detailed error messages

## 📝 Usage Instructions

1. **Upload Files:**
   - Activities (Master) Excel file
   - GIN (Good Issue Note) Excel file
   - CostHead mapping Excel file

2. **Configure Settings:**
   - Select match mode (contains/exact)
   - Review data preview

3. **Generate Report:**
   - Click "Generate CostHead Report"
   - Wait for processing
   - Download the ZIP file

4. **Review Results:**
   - Check the generated reports
   - Verify data accuracy
   - Download additional files if needed

## 🆘 Support

If you encounter issues:

1. Check the browser console for errors
2. Review the Streamlit logs
3. Ensure all dependencies are installed
4. Verify file formats are correct
5. Check available disk space

## 🔄 Updates and Maintenance

To update the application:

1. Make changes to the code
2. Test locally first
3. Commit changes to Git
4. Redeploy to your chosen platform
5. Verify the deployment works

---

**🎉 Congratulations!** Your CostHead Report Generator is now ready for web deployment!
