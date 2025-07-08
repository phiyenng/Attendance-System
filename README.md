# Attendance Report System

A modern Flask web application that replicates the functionality of your Excel attendance report tool with VBA code. This system allows you to import different types of attendance data, calculate abnormal attendance, and export results - all through a beautiful web interface.

## Features

- 📁 **Import 3 Data Types**: Upload Sign In/Out Data, Apply Data, and OT Lieu Data
- 🔄 **Refresh**: Recalculate all attendance logic (equivalent to your VBA RecalculateWorkbook macro)
- ⚠️ **Calculate Abnormal**: Find and display abnormal attendance records
- 🗑️ **Clear Data**: Clear individual data types (Sign In/Out, Apply, OT Lieu)
- 📥 **Export**: Download processed data as Excel file with multiple sheets
- 📊 **Real-time Status**: View data counts and processing status
- 🎨 **Modern UI**: Beautiful, responsive design with gradient backgrounds
- 📱 **Mobile Friendly**: Works on desktop and mobile devices

## Installation

1. **Clone or download this repository**
   ```bash
   git clone <repository-url>
   cd Attendance-Record
   ```

2. **Install Python dependencies**
   ```bash
   pip install -r requirements.txt
   ```

3. **Run the application**
   ```bash
   python app.py
   ```

4. **Open your browser**
   Navigate to `http://localhost:5000`

## Usage

### 1. Import Data
Upload your three types of data files:

- **Sign In/Out Data**: Employee attendance records
- **Apply Data**: Leave, trip, and supplement applications
- **OT Lieu Data**: Overtime and lieu time records

### 2. Process Data
- Click **"Refresh"** to recalculate all attendance logic
- Click **"Calculate Abnormal"** to find attendance issues
- View abnormal records in the table below

### 3. Manage Data
- Use **"Clear"** buttons to remove specific data types
- View data status in the status card

### 4. Export Results
- Click **"Export"** to download all processed data as an Excel file
- The exported file contains separate sheets for each data type

## File Structure

```
Attendance-Record/
├── app.py                 # Main Flask application
├── requirements.txt       # Python dependencies
├── templates/
│   └── index.html        # Main web interface
├── uploads/              # Temporary file storage
├── start.bat            # Windows startup script
├── start.sh             # Unix/Linux startup script
├── run.py               # Smart startup script
└── README.md            # This file
```

## Technical Details

### Backend (Flask)
- **Data Processing**: Uses pandas for data manipulation
- **VBA Logic Translation**: Implements your VBA functions in Python
- **File Handling**: Uses openpyxl for Excel file operations
- **API Endpoints**: RESTful API for all operations

### Key Functions Implemented
- `check_apply()`: Python version of your CheckApply VBA function
- `check_lieu()`: Python version of your checkLieu VBA function
- `calculate_attendance()`: Main calculation logic (equivalent to RecalculateWorkbook)
- `is_holiday()`: Holiday checking logic
- `get_day_type()`: Weekday/Weekend determination

### Frontend (HTML/CSS/JavaScript)
- **Responsive Design**: Bootstrap 5 for mobile-friendly layout
- **Real-time Updates**: AJAX communication with backend
- **Modern UI**: Gradient backgrounds and smooth animations
- **Interactive Tables**: Display abnormal attendance data

### Supported File Types
- **Input**: .xlsx, .xls, .csv files
- **Output**: .xlsx files with multiple sheets

## API Endpoints

- `GET /` - Main page
- `POST /import/signinout` - Import Sign In/Out data
- `POST /import/apply` - Import Apply data
- `POST /import/otlieu` - Import OT Lieu data
- `POST /refresh` - Refresh/recalculate all data
- `POST /calculate_abnormal` - Calculate abnormal attendance
- `GET /get_abnormal_data` - Get abnormal data for display
- `POST /clear/signinout` - Clear Sign In/Out data
- `POST /clear/apply` - Clear Apply data
- `POST /clear/otlieu` - Clear OT Lieu data
- `GET /export` - Export data as Excel file
- `GET /get_data_status` - Get data status

## VBA to Python Translation

The application translates your VBA logic to Python:

### Core Functions
- **CheckApply**: Checks if employee has approved leave/trip/supplement
- **checkLieu**: Validates lieu time against attendance
- **RecalculateWorkbook**: Main calculation engine
- **CalcAbnormal**: Identifies attendance issues

### Data Processing
- **Sign In/Out Analysis**: Late/early detection
- **Apply Data Integration**: Leave and trip validation
- **OT Lieu Calculation**: Overtime and lieu time processing
- **Abnormal Detection**: Missing attendance, late arrival, early departure

## Error Handling

The application includes comprehensive error handling for:
- Invalid file types
- Corrupted Excel files
- Missing data columns
- Calculation errors
- Network issues

## Browser Compatibility

- Chrome (recommended)
- Firefox
- Safari
- Edge

## Troubleshooting

### Common Issues

1. **File won't upload**
   - Ensure file is .xlsx, .xls, or .csv format
   - Check file size (max 16MB)
   - Verify file is not corrupted

2. **Calculations not working**
   - Ensure all required data types are uploaded
   - Check data format matches expected structure
   - Verify date/time formats are correct

3. **Application won't start**
   - Verify all dependencies are installed
   - Check Python version (3.7+ required)
   - Ensure port 5000 is available

### Getting Help

If you encounter issues:
1. Check the browser console for JavaScript errors
2. Review the Flask application logs
3. Verify your data file structure

## Development

To modify or extend the application:

1. **Add new features**: Modify `app.py` for backend logic
2. **Update UI**: Edit `templates/index.html` for frontend changes
3. **Add dependencies**: Update `requirements.txt`
4. **Enhance VBA logic**: Extend the Python functions in `app.py`

## License

This project is open source and available under the MIT License.

## Contributing

Contributions are welcome! Please feel free to submit pull requests or open issues for bugs and feature requests. 