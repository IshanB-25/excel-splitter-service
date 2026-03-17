# Excel File Splitter Service

A production-ready Flask web service that splits multi-sheet Excel files into individual Excel files, one for each sheet. Perfect for processing large Excel workbooks with multiple tabs.

## 🚀 Features

- **Sheet Separation**: Automatically splits Excel files with multiple sheets into individual files per sheet
- **Smart Response**: Returns a single file for one visible sheet, or a `.zip` for multiple sheets
- **Structure Preservation (XLSX mode)**: Keeps merged ranges, row heights, column widths, sheet names, and formulas (without style formatting)
- **Optional TXT Output**: Export each sheet as tab-separated `.txt` while retaining grid layout and metadata
- **Memory Conscious**: Uses temporary files during processing to reduce peak memory pressure
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
docker run -p 3070:3070 \
  -e MAX_FILE_SIZE_MB=100 \
  -e GUNICORN_WORKERS=1 \
  -e GUNICORN_TIMEOUT=1800 \
  -e GUNICORN_MAX_REQUESTS=25 \
  -e GUNICORN_MAX_REQUESTS_JITTER=10 \
  excel-splitter
```

## 📡 API Endpoints

### `POST /split-excel`
Upload an Excel file to split into individual sheets.

**Request:**
- Method: `POST`
- Content-Type: `multipart/form-data`
- Field name: `file`
- Optional query/form parameter: `format` (`xlsx` default, `txt` optional)

**Response:**
- Single sheet: Returns `.xlsx` or `.txt` directly (based on `format`)
- Multiple sheets: Returns `.zip` containing one file per sheet

**Example using curl:**
```bash
curl -X POST -F "file=@your-excel-file.xlsx" http://localhost:3070/split-excel -o output.zip

# Optional TXT mode
curl -X POST -F "file=@your-excel-file.xlsx" "http://localhost:3070/split-excel?format=txt" -o output.zip
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

### `POST /split-excel-ndjson`
Upload an Excel file and stream extracted sheet rows as NDJSON.

**Request:**
- Method: `POST`
- Content-Type: `multipart/form-data`
- Field name: `file`

**Response:**
- Content-Type: `application/x-ndjson`
- One JSON object per line:
```json
{"name":"Sheet1","row":1,"content":"A\tB\tC"}
{"name":"Sheet1","row":2,"content":"1\t2\t3"}
```

**Example using curl:**
```bash
curl -X POST -F "file=@your-excel-file.xlsx" \
  http://localhost:3070/split-excel-ndjson \
  -o output.ndjson
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
  "version": "2.2.0",
  "description": "Split Excel files by visible sheets while preserving structure",
  "features": [
    "Sheet structure preservation (no cell styling)",
    "Skips hidden sheets",
    "Preserves cell values and formulas",
    "Single-sheet direct response or multi-sheet ZIP",
    "Streaming NDJSON endpoint for low-memory extraction"
  ],
  "configuration": {
    "max_file_size": "50.0 MB",
    "max_sheets": 100,
    "allowed_extensions": ["xlsx", "xls", "xlsm", "xlsb"],
    "output_formats": ["xlsx", "txt", "ndjson"]
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
| `GUNICORN_WORKERS` | Number of Gunicorn worker processes | `1` |
| `GUNICORN_TIMEOUT` | Worker timeout in seconds for long sheet processing | `1800` |
| `GUNICORN_GRACEFUL_TIMEOUT` | Graceful shutdown timeout in seconds | `1800` |
| `GUNICORN_MAX_REQUESTS` | Restart worker after N requests (helps cap memory growth) | `25` |
| `GUNICORN_MAX_REQUESTS_JITTER` | Randomized spread for worker restarts | `10` |
| `ZIP_COMPRESSION_LEVEL` | ZIP compression level (0-9, lower is faster, applies to XLSX bundles) | `1` |

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

ENV GUNICORN_WORKERS=1
ENV GUNICORN_TIMEOUT=1800
ENV GUNICORN_GRACEFUL_TIMEOUT=1800
ENV GUNICORN_MAX_REQUESTS=25
ENV GUNICORN_MAX_REQUESTS_JITTER=10
CMD ["sh", "-c", "gunicorn --bind 0.0.0.0:3070 --workers ${GUNICORN_WORKERS} --timeout ${GUNICORN_TIMEOUT} --graceful-timeout ${GUNICORN_GRACEFUL_TIMEOUT} --max-requests ${GUNICORN_MAX_REQUESTS} --max-requests-jitter ${GUNICORN_MAX_REQUESTS_JITTER} app:app"]
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
      - GUNICORN_WORKERS=1
      - GUNICORN_TIMEOUT=1800
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

### In `xlsx` output mode (default)

✅ **Preserved structure/content**
- Cell values
- Formulas
- Sheet names
- Merged-cell ranges
- Column widths
- Row heights
- Freeze panes and auto-filter refs

⚠️ **Not preserved**
- Cell style formatting (font/fill/border/alignment)

### In `txt` output mode (`?format=txt`)

✅ **Preserved**
- Row/column grid layout as tab-separated rows
- Sheet values/formulas as text
- Structural metadata headers (`sheet_name`, `max_row`, `max_col`, `merged_ranges`)

⚠️ **Not preserved**
- Native Excel visual/layout features (styles, freeze panes, filters, widths/heights)

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
{original_filename}_{sheet_name}.{xlsx|txt}
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