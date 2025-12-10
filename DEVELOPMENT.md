# Development Configuration

## Environment Setup

### Python Version
- Minimum: Python 3.9
- Recommended: Python 3.10+

### Virtual Environment
```powershell
python -m venv venv
.\venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

## Streamlit Configuration

### Development Server
```powershell
streamlit run Home.py
```

### Custom Port
```powershell
streamlit run Home.py --server.port=8502
```

### Debug Mode
```powershell
streamlit run Home.py --logger.level=debug
```

## Code Style Guidelines

### Imports
```python
# Order: Standard library → Third-party → Local
import os
import sys
from datetime import datetime

import pandas as pd
import streamlit as st

from utils import set_page, header, footer
from utils import render_card, render_alert
```

### Component Usage
```python
from utils import render_card

render_card(
    title="Example",
    content="Content here",
    footer="Optional footer",
    icon="🎯"
)
```

### Comments
- Use meaningful variable names
- Comment complex logic
- Use docstrings untuk functions

### File Structure
```python
# ====================================================
# Section Title
# ====================================================
# Code dalam section

# ====================================================
# Another Section
# ====================================================
# Code dalam section
```

## Git Workflow

### Before Commit
1. Test locally dengan `streamlit run Home.py`
2. Check untuk syntax errors: `python -m py_compile filename.py`
3. Verify imports bekerja

### Commit Messages
- Use present tense: "Add feature" not "Added feature"
- Be descriptive: "Add render_card component" good, "fix" bad
- Reference issues if applicable: "Fix #123"

## Testing

### Manual Testing
1. Test di local environment
2. Test di different screen sizes
3. Test file uploads
4. Test Excel exports
5. Check console untuk errors

### File Validation
```python
# Check syntax
python -m py_compile utils/components.py

# Check imports
python -c "from utils import render_card; print('OK')"
```

## Common Tasks

### Add New Component

1. Create function di `utils/components.py`
2. Add docstring dengan contoh penggunaan
3. Export di `utils/__init__.py`
4. Document di `UI_UX_GUIDE.md`
5. Add example di `EXAMPLE_COMPONENTS.py`

### Add New Tool Page

1. Create file di `pages/` folder: `N_Tool_Name.py`
2. Follow naming convention: `1_Tool_Name.py`
3. Use `set_page()` dan `header()` dari utils
4. Use reusable components untuk konsistensi
5. End dengan `footer()`

### Update Styling

1. Edit CSS di `utils/ui.py` function `_apply_custom_styles()`
2. Or update theme di `utils/theme.py`
3. Test rendering di `EXAMPLE_COMPONENTS.py`
4. Document changes di `CHANGELOG.md`

## Troubleshooting

### Import Error
```python
# Check sys.path
import sys
print(sys.path)

# Verify file exists
import os
print(os.path.exists('utils/components.py'))
```

### Port Already in Use
```powershell
# Kill process on port 8501
netstat -ano | findstr :8501
taskkill /PID <PID> /F
```

### Cache Issues
```powershell
# Clear Streamlit cache
Remove-Item -Recurse -Force .streamlit/cache/
```

## Resources

- Streamlit Docs: https://docs.streamlit.io
- Python Docs: https://docs.python.org/3
- Pandas Docs: https://pandas.pydata.org/docs

## Support

📧 nazarudin@gsid.co.id
