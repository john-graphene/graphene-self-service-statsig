set "current_dir=%cd%"
if not exist "%current_dir%/statsig_streamlit" (
echo "venv doesnt exist, installing venv and requirements..."
py -3.10 -m venv statsig_streamlit
cd statsig_streamlit\Scripts
call activate.bat
pip install -r ../../requirements.txt
cd ../..
) else (
cd statsig_streamlit\Scripts
call activate.bat
cd ../..
)
streamlit run --server.port 8080 --server.enableCORS false statsig_streamlit.py