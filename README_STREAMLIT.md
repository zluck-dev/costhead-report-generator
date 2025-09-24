# 📊 CostHead Report Generator - Streamlit Web App

A modern web application for generating CostHead reports from Excel data, built with Streamlit.

## 🚀 Quick Start

### 1. Setup Environment

```bash
# Create virtual environment
python3 -m venv streamlit_env

# Activate virtual environment
source streamlit_env/bin/activate  # On macOS/Linux
# or
streamlit_env\Scripts\activate    # On Windows

# Install dependencies
pip install -r requirements.txt
```

### 2. Run the Application

```bash
# Method 1: Using the launch script
python run_streamlit.py

# Method 2: Direct Streamlit command
streamlit run streamlit_app.py

# Method 3: With custom port
streamlit run streamlit_app.py --server.port 8502
```

### 3. Access the App

Open your web browser and go to: **http://localhost:8501**

## 📋 Features

### ✅ Core Functionality
- **File Upload**: Drag-and-drop Excel file uploads
- **Data Preview**: Real-time preview of uploaded data
- **Auto Column Detection**: Automatically detects required columns
- **Report Generation**: Full CostHead report generation
- **Download Functionality**: ZIP file download with all reports
- **Progress Indicators**: Loading spinners and progress bars
- **Error Handling**: Comprehensive error messages

### 🎨 User Interface
- **Modern Design**: Clean, professional interface
- **Responsive Layout**: Works on desktop and mobile
- **Sidebar Navigation**: Easy file upload and settings
- **Tabbed Preview**: Organized data viewing
- **Status Indicators**: Clear success/error feedback
- **Download Buttons**: One-click report downloads

## 📁 File Structure

```
excel-project/
├── streamlit_app.py          # Main Streamlit application
├── run_streamlit.py          # Launch script
├── test_streamlit.py         # Test suite
├── requirements.txt          # Python dependencies
├── STREAMLIT_DEPLOYMENT.md   # Deployment guide
├── README_STREAMLIT.md       # This file
├── streamlit_env/            # Virtual environment (created)
└── app/
    └── services/             # Core business logic
        ├── excel_service.py
        ├── gin_service.py
        ├── costhead_service.py
        └── export_service.py
```

## 🛠️ How to Use

### Step 1: Upload Files
1. **Activities File**: Upload your master activities Excel file
2. **GIN File**: Upload your Good Issue Note Excel file
3. **CostHead Mapping**: Upload your CostHead mapping Excel file

### Step 2: Configure Settings
- Select match mode: "contains" or "exact"
- Review data preview in the tabs
- Check that all required columns are detected

### Step 3: Generate Report
- Click "🚀 Generate CostHead Report"
- Wait for processing to complete
- Download the ZIP file with all reports

### Step 4: Review Results
- Check the generated reports
- Verify data accuracy
- Download additional files if needed

## 🌐 Deployment Options

### Option 1: Streamlit Cloud (Free)
1. Push your code to GitHub
2. Go to [share.streamlit.io](https://share.streamlit.io)
3. Connect your repository
4. Deploy with one click

### Option 2: Heroku
1. Create a `Procfile` with: `web: streamlit run streamlit_app.py --server.port=$PORT --server.address=0.0.0.0`
2. Deploy using Heroku CLI

### Option 3: AWS EC2
1. Launch an EC2 instance
2. Install dependencies
3. Run with: `nohup streamlit run streamlit_app.py --server.port 80 --server.address 0.0.0.0 &`

### Option 4: Docker
1. Create a Dockerfile
2. Build and run the container

See `STREAMLIT_DEPLOYMENT.md` for detailed instructions.

## 🔧 Configuration

### Environment Variables
```bash
STREAMLIT_SERVER_PORT=8501
STREAMLIT_SERVER_ADDRESS=0.0.0.0
STREAMLIT_BROWSER_GATHER_USAGE_STATS=false
```

### Streamlit Config (`.streamlit/config.toml`)
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

## 🧪 Testing

Run the test suite to verify everything is working:

```bash
# Activate virtual environment
source streamlit_env/bin/activate

# Run tests
python test_streamlit.py
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

## 📊 What's Different from Desktop Version?

### ✅ Advantages of Streamlit Version:
- **Web-based**: Access from any device with a browser
- **No Installation**: Users don't need to install Python or dependencies
- **Easy Sharing**: Share a URL instead of distributing files
- **Automatic Updates**: Deploy updates without user intervention
- **Better UI**: Modern, responsive web interface
- **Cross-platform**: Works on Windows, Mac, Linux, mobile

### 🔄 Same Core Functionality:
- All Excel processing logic is identical
- Same report generation algorithms
- Same file format support
- Same data validation and error handling

## 🆘 Support

If you encounter issues:

1. **Check the browser console** for JavaScript errors
2. **Review the Streamlit logs** in the terminal
3. **Ensure all dependencies** are installed correctly
4. **Verify file formats** are correct Excel files
5. **Check available disk space** for file processing

## 🔄 Updates and Maintenance

To update the application:

1. Make changes to the code
2. Test locally first: `python test_streamlit.py`
3. Commit changes to Git
4. Redeploy to your chosen platform
5. Verify the deployment works

## 📈 Performance Considerations

- **File Size Limits**: Large Excel files may take longer to process
- **Memory Usage**: Complex reports may require more RAM
- **Concurrent Users**: Consider server resources for multiple users
- **Caching**: Streamlit automatically caches data for better performance

---

## 🎉 Success!

Your CostHead Report Generator is now a modern web application!

**Next Steps:**
1. Test the application locally
2. Deploy to your preferred platform
3. Share the URL with your users
4. Enjoy the modern web interface!

**Need Help?** Check the `STREAMLIT_DEPLOYMENT.md` file for detailed deployment instructions.
