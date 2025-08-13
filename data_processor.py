import os
import pandas as pd
from datetime import datetime
import pytz
import zipfile
import logging
from utils import METRICS, ANALYSIS_VALUES

class DataProcessor:
    """Handles data extraction and processing from ZIP files"""
    
    def __init__(self, exclude_backhauls=True):
        self.exclude_backhauls = exclude_backhauls
        self.logger = logging.getLogger(__name__)
    
    def extract_files(self, input_path, output_path):
        """Extract Excel/CSV files from ZIP archives"""
        self.logger.info("Extracting ZIP contents...")
        os.makedirs(output_path, exist_ok=True)
        
        for root, _, files in os.walk(input_path):
            for file in files:
                if file.endswith('.zip'):
                    try:
                        with zipfile.ZipFile(os.path.join(root, file), 'r') as zip_ref:
                            for member in zip_ref.namelist():
                                should_extract = (
                                    member.endswith(('.xlsx', '.xls', '.csv')) and
                                    (not self.exclude_backhauls or "Unplanned_Backhauls_Reason" not in member)
                                )
                                if should_extract:
                                    zip_ref.extract(member, output_path)
                                    self.logger.info(f"Extracted: {member}")
                    except zipfile.BadZipFile:
                        self.logger.error(f"Bad ZIP file: {file}")
                    except Exception as e:
                        self.logger.error(f"Error extracting {file}: {e}")
    
    def process_extracted_files(self, input_folder):
        """Process extracted Excel files and return DataFrames"""
        self.logger.info("Processing extracted Excel files...")
        sheet1_data, sheet2_data, sheet3_data = [], [], []
        
        for file in os.listdir(input_folder):
            if file.endswith(('.xlsx', '.xls')):
                path = os.path.join(input_folder, file)
                try:
                    # Process each file
                    totals_data, mdt_data, dispatch_data = self._process_single_file(path, file)
                    
                    sheet1_data.extend(totals_data)
                    sheet2_data.extend(mdt_data)
                    sheet3_data.extend(dispatch_data)
                    
                    self.logger.info(f"Processed: {file}")
                    
                except Exception as e:
                    self.logger.error(f"Error processing {file}: {e}")
        
        # Create DataFrames
        df1 = self._create_totals_dataframe(sheet1_data)
        df2 = self._create_mdt_dataframe(sheet2_data)
        df3 = self._create_dispatch_dataframe(sheet3_data)
        
        # Apply enhancements
        df2 = self._map_time_ranges(df2)
        df3 = self._enhance_dispatch_data(df3)
        
        # Clean up extracted files
        self._cleanup_extracted_files(input_folder)
        
        return df1, df2, df3
    
    def _process_single_file(self, path, filename):
        """Process a single Excel file and extract data"""
        # Read full file for totals and MDT data
        df_full = pd.read_excel(path)
        df_full['Source'] = filename
        
        # Read dispatch data (skip header rows)
        start_row = self._detect_data_start_row(path)
        df_dispatch = pd.read_excel(path, skiprows=start_row)
        df_dispatch['Source'] = filename
        
        # Extract totals data
        df_totals = df_full[df_full.iloc[:, 0] == "Total"].copy()
        df_totals = pd.concat([df_totals, self._split_filename_metadata(df_totals)], axis=1)
        df_totals = self._clean_baseline_column(df_totals)
        
        # Extract MDT data
        df_mdt = df_full[
            (df_full.iloc[:, 4] == "DEPOT") &
            (df_full.iloc[:, 5] == "DC") &
            (df_full.iloc[:, 17] == 0)
        ].copy()
        df_mdt['Source'] = filename
        df_mdt.insert(1, 'Hours', df_mdt.iloc[:, 0].apply(self._convert_to_eastern_hour))
        df_mdt = pd.concat([df_mdt, self._split_filename_metadata(df_mdt)], axis=1)
        df_mdt = self._clean_baseline_column(df_mdt)
        
        # Process dispatch data
        df_dispatch = pd.concat([df_dispatch, self._split_filename_metadata(df_dispatch)], axis=1)
        df_dispatch = self._clean_baseline_column(df_dispatch)
        
        return df_totals.values.tolist(), df_mdt.values.tolist(), df_dispatch.values.tolist()
    
    def _detect_data_start_row(self, path):
        """Detect the row where actual data starts"""
        try:
            preview = pd.read_excel(path, nrows=20, header=None)
            for i, value in enumerate(preview.iloc[:, 0]):
                if str(value).strip().lower() == "activity start time":
                    return i
        except Exception as e:
            self.logger.warning(f"Could not detect header row in {path}: {e}")
        return 11
    
    def _convert_to_eastern_hour(self, timestamp):
        """Convert timestamp to Eastern timezone hour"""
        if pd.notnull(timestamp):
            try:
                timestamp = timestamp[:-4]  # Remove timezone suffix
                naive_dt = datetime.strptime(timestamp, "%Y-%m-%d %H:%M")
                return pytz.timezone('US/Eastern').localize(naive_dt).hour
            except Exception as e:
                self.logger.warning(f"Failed to convert timestamp '{timestamp}': {e}")
        return None
    
    def _split_filename_metadata(self, df, source_col='Source'):
        """Extract metadata from filename"""
        try:
            parts = df[source_col].str.split('_', expand=True).reindex(columns=range(7), fill_value=None)
            return parts.rename(columns={
                0: 'Date', 1: 'Report', 2: 'Store ID', 3: 'Country', 
                4: 'Product', 5: 'ID', 6: 'Baseline'
            })
        except Exception as e:
            self.logger.error(f"Error splitting filename metadata: {e}")
            return pd.DataFrame()
    
    def _clean_baseline_column(self, df):
        """Clean baseline column by removing file extensions"""
        if 'Baseline' in df.columns:
            df['Baseline'] = df['Baseline'].astype(str).apply(
                lambda x: x.split('.')[0] if pd.notnull(x) else x
            )
        return df
    
    def _map_time_ranges(self, df):
        """Map hours to time ranges"""
        if not df.empty and 'Hours' in df.columns:
            time_ranges = {i: f"{i:02d}:00-{(i + 1) % 24:02d}:00" for i in range(24)}
            df['Time Range'] = df['Hours'].map(time_ranges)
        else:
            df['Time Range'] = None
        return df
    
    def _enhance_dispatch_data(self, df):
        """Add stop/position counts and trailer utilization formulas"""
        # Add stop positions for STORE unit types
        if all(col in df.columns for col in ['Route Number', 'Simulation', 'Unit Type', 'Comment']):
            df_store = df[
                (df['Unit Type'] == 'STORE') &
                (df['Route Number'].notna()) &
                (df['Simulation'].notna())
            ].copy()
            
            def assign_stop_position(group):
                group = group.copy()
                total_stops = len(group)
                group['Comment'] = [f"Stop {i+1} of {total_stops}" for i in range(total_stops)]
                return group
            
            df_store_updated = df_store.groupby(
                ['Route Number', 'Simulation'], group_keys=False
            ).apply(assign_stop_position)
            
            df['Comment'] = df['Comment'].astype(str)
            df.loc[df_store_updated.index, 'Comment'] = df_store_updated['Comment'].values
        
        # Add trailer utilization formula to Trip Id column
        df = df.copy()
        for idx, row in df.iterrows():
            if pd.isna(row['Activity start time']):
                excel_row = idx + 2  # Excel rows start at 1, header is row 1
                df.at[idx, 'Trip Id'] = f'=TEXT(N{excel_row}/41500,"0.00%")'
        
        return df
    
    def _create_totals_dataframe(self, data):
        """Create totals DataFrame with proper column names"""
        columns = [
            'Total', 'Duration', 'Pallets', 'Cubes', 'Cases', 'Pounds', 'Routes', 
            'Stops', 'Distance', 'Cube Per Mile', '11', '12', '13', '14', '15', 
            '16', '17', '18', '19', '20', 'Source', 'Date', 'Report', 'DC', 
            'Country', 'Commodity', 'Release ID', 'Simulation'
        ]
        return pd.DataFrame(data, columns=columns)
    
    def _create_mdt_dataframe(self, data):
        """Create MDT DataFrame with proper column names"""
        columns = [
            'Activity start time', 'Hours', 'Trailer', 'Backhaul Info', 'Route Number',
            'Activity Type', 'Unit Type', 'Unit Profile', 'Drop Profile', 'Start Window',
            'End Window', 'DC', 'City', 'State', 'LBS', 'Pallets', 'Cubes', 'Cases',
            'Distance', 'Comment', 'Trip Id', 'Source', 'Date', 'Report', 'DC ID',
            'Country', 'Commodity', 'Release ID', 'Simulation'
        ]
        return pd.DataFrame(data, columns=columns)
    
    def _create_dispatch_dataframe(self, data):
        """Create dispatch DataFrame with proper column names"""
        columns = [
            'Activity start time', 'Trailer', 'Backhaul Info', 'Route Number',
            'Activity Type', 'Unit Type', 'Unit Profile', 'Drop Profile', 'Start Window',
            'End Window', 'Store Club Id', 'Store Club City', 'Store Club State', 'LBS',
            'Pallets', 'Cubes', 'Cases', 'Distance', 'Comment', 'Trip Id', 'Source',
            'Date', 'Report', 'DC', 'Country', 'Commodity', 'Release ID', 'Simulation'
        ]
        return pd.DataFrame(data, columns=columns)
    
    def _cleanup_extracted_files(self, folder):
        """Clean up extracted files after processing"""
        self.logger.info("Cleaning up extracted files...")
        for file in os.listdir(folder):
            try:
                os.remove(os.path.join(folder, file))
            except Exception as e:
                self.logger.warning(f"Could not delete {file}: {e}")
