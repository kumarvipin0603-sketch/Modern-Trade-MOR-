@echo off
echo Installing required Python packages for PO Control Tower...
py -m pip install -r requirements.txt
echo.
echo Done. Now run:
echo py -m streamlit run app.py
pause
