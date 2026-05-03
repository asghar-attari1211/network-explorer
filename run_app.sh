#!/bin/bash

# Port to run the Streamlit app on (default is 8501, you can change this)
PORT=8080

echo "Starting Network Explorer Dashboard on port $PORT..."
streamlit run NetworkExplorer.py --server.port $PORT --server.address 0.0.0.0
