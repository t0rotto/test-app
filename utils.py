import logging
import streamlit as st

# === CONSTANTS ===
METRICS = ['Pallets', 'Cubes', 'Cases', 'Pounds', 'Routes', 'Stops', 'Distance']
ANALYSIS_VALUES = [
    'No. of Pallets', 'Total Cube (cu ft)', 'Total Cases', 'Total Weight (lbs)', 
    'Routes', 'Stops', 'Distance (miles)', 'CPT', 'LOH'
]

def setup_logging():
    """Configure logging for the application"""
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        force=True  # Force reconfiguration
    )
    
    # Suppress warnings for cleaner output
    import warnings
    warnings.filterwarnings("ignore")
    
    return logging.getLogger(__name__)
