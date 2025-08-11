# Excel File Splitter Service

A production-ready Flask web service that splits multi-sheet Excel files into individual Excel files, one for each sheet. Perfect for processing large Excel workbooks with multiple tabs.

## 🚀 Features

- **Sheet Separation**: Automatically splits Excel files with multiple sheets into individual `.xlsx` files
- **Smart Response**: Returns a single `.xlsx` file for single-sheet workbooks, or a `.zip` file for multiple sheets
- **Format Preservation**: Maintains cell values, formulas, merged cells, column widths, row heights, and basic formatting
- **Memory Efficient**: Processes files entirely in memory without temporary file storage
- **Production Ready**: Includes health checks, comprehensive error handling, and detailed logging
- **Docker Support**: Fully containerized with Gunicorn for production deployment

## 📋 Requirements

- Python 3.11+
- Flask 3.0.0
- openpyxl 3.1.2
- Werkzeug 3.0.1
- gunicorn 21.2.0 (for production)

## 🛠️ Installation

### Local Development

1. Clone the repository:
```bash
git clone https://github.com/IshanB-25/excel-splitter-service.git
cd excel-splitter-service
```

2. Create a virtual environment:
```bash
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
```

3. Install dependencies:
```bash
pip install -r requirements.txt
```

4. Run the service:
```bash
python app.py
```

The service will start on `http://localhost:3070`

### Docker Deployment

1. Build the Docker image:
```bash
docker build -t excel-splitter .
```

2. Run the container:
```bash
docker run -p 3070:3070 -e MAX_FILE_SIZE_MB=100 excel-splitter
```

## 📡 API Endpoints

### `POST /split-excel`
Upload an Excel file to split into individual sheets.

**Request:**
- Method: `POST`
- Content-Type: `multipart/form-data`
- Field name: `file`

**Response:**
- Single sheet: Returns `.xlsx` file directly
- Multiple sheets: Returns `.zip` file containing all split files

**Example using curl:**
```bash
curl -X POST -F "file=@your-excel-file.xlsx" http://localhost:3070/split-excel -o output.zip
```

**Example using Python:**
```python
import requests

with open('your-excel-file.xlsx', 'rb') as f:
    files = {'file': f}
    response = requests.post('http://localhost:3070/split-excel', files=files)
    
if response.status_code == 200:
    with open('output.zip', 'wb') as output:
        output.write(response.content)
```

### `GET /health`
Health check endpoint for monitoring.

**Response:**
```json
{
  "status": "healthy",
  "timestamp": "2025-08-11T18:30:00.000000",
  "service": "excel-splitter"
}
```

### `GET /`
Service information and configuration details.

**Response:**
```json
{
  "service": "Excel File Splitter",
  "version": "2.1.0",
  "description": "Split Excel files by sheets while preserving data and structure",
  "features": [
    "Preserves cell values and formulas",
    "Maintains merged cells",
    "Keeps column widths and row heights",
    "Preserves basic cell formatting",
    "Maintains number formats"
  ],
  "configuration": {
    "max_file_size": "50.0 MB",
    "max_sheets": 100,
    "allowed_extensions": ["xlsx", "xls", "xlsm", "xlsb"]
  }
}
```

## ⚙️ Configuration

Configure the service using environment variables:

| Variable | Description | Default |
|----------|-------------|---------|
| `PORT` | Port to run the service on | `3070` |
| `MAX_FILE_SIZE_MB` | Maximum file size in megabytes | `50` |
| `MAX_SHEETS` | Maximum number of sheets to process | `100` |

### Example Configuration:
```bash
export PORT=8080
export MAX_FILE_SIZE_MB=200
export MAX_SHEETS=50
```

## 🐳 Docker Deployment

### Dockerfile
```dockerfile
FROM python:3.11-slim

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY app.py .

EXPOSE 3070

CMD ["gunicorn", "--bind", "0.0.0.0:3070", "--workers", "4", "--timeout", "120", "app:app"]
```

### Docker Compose
```yaml
version: '3.8'

services:
  excel-splitter:
    build: .
    ports:
      - "3070:3070"
    environment:
      - MAX_FILE_SIZE_MB=100
      - MAX_SHEETS=100
      - PORT=3070
    restart: unless-stopped
```

## 🚢 Deployment on Coolify

1. **Prepare your repository** with:
   - `app.py` - Main application file
   - `requirements.txt` - Python dependencies
   - `Dockerfile` - Container configuration

2. **In Coolify:**
   - Create a new service
   - Select "Docker" as the build type
   - Connect your Git repository
   - Set environment variables:
     ```
     MAX_FILE_SIZE_MB=100
     MAX_SHEETS=100
     PORT=3070
     ```
   - Deploy the service

3. **Configure health checks** in Coolify:
   - Health check path: `/health`
   - Expected status: `200`

## 📊 What Gets Preserved

When splitting Excel files, the service maintains:

✅ **Data & Content**
- Cell values
- Formulas
- Number formats (dates, currency, percentages)

✅ **Structure**
- Merged cells
- Column widths
- Row heights
- Sheet names

✅ **Basic Formatting**
- Font styles and sizes
- Cell colors and fills
- Borders
- Text alignment

## 🔍 Error Handling

The service includes comprehensive error handling for:

- Invalid file types (non-Excel files)
- File size limit exceeded
- Corrupted Excel files
- Processing errors for individual sheets
- Memory constraints

Error responses include clear messages:
```json
{
  "error": "File too large. Maximum size: 100.0 MB"
}
```

## 📝 File Naming Convention

Split files are named using the pattern:
```
{original_filename}_{sheet_name}.xlsx
```

For example:
- Input: `report.xlsx` with sheets "Sales" and "Inventory"
- Output: `report_Sales.xlsx`, `report_Inventory.xlsx`

## 🧪 Testing

### Test with curl:
```bash
# Test health endpoint
curl http://localhost:3070/health

# Test splitting an Excel file
curl -X POST -F "file=@test.xlsx" http://localhost:3070/split-excel -o result.zip

# Extract the ZIP file
unzip result.zip
```

### Test with Python:
```python
import requests
import zipfile
import io

# Upload and split file
with open('test.xlsx', 'rb') as f:
    response = requests.post(
        'http://localhost:3070/split-excel',
        files={'file': f}
    )

# Handle response based on content type
if response.headers.get('Content-Type') == 'application/zip':
    # Multiple sheets - extract ZIP
    with zipfile.ZipFile(io.BytesIO(response.content)) as z:
        z.extractall('output/')
else:
    # Single sheet - save directly
    with open('output.xlsx', 'wb') as f:
        f.write(response.content)
```

## 🐛 Troubleshooting

### Common Issues:

1. **"File too large" error**
   - Increase `MAX_FILE_SIZE_MB` environment variable
   - Default is 50MB

2. **"No sheets found" error**
   - Ensure the Excel file is not corrupted
   - Check if the file has at least one visible sheet

3. **Timeout errors with large files**
   - Increase Gunicorn timeout in Dockerfile: `--timeout 300`
   - Consider increasing worker memory

4. **Memory issues**
   - Large files are processed in memory
   - Ensure container has sufficient RAM
   - Consider splitting very large files in batches

## 📄 License

[Your License Here]

## 🤝 Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## 📧 Support

For issues and questions, please create an issue in the repository or contact [your-email@example.com]

---

**Version:** 2.1.0  
**Last Updated:** August 2025