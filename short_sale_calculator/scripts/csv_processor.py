import pandas as pd
import os
from typing import List

class CSVProcessor:
    def __init__(self):
        pass
    
    def combine_csvs(self, file_paths: List[str], output_path: str) -> bool:
        """
        Combines multiple CSV files into a single CSV file.
        
        Args:
            file_paths: List of paths to CSV files to combine
            output_path: Path where the combined CSV should be saved
            
        Returns:
            bool: True if successful, False otherwise
        """
        try:
            combined_data = []
            
            for file_path in file_paths:
                if not os.path.exists(file_path):
                    print(f"Warning: File {file_path} does not exist, skipping...")
                    continue
                
                # Read CSV file
                df = pd.read_csv(file_path)
                
                # Add source file column to track origin
                df['source_file'] = os.path.basename(file_path)
                
                combined_data.append(df)
            
            if not combined_data:
                print("No valid CSV files found to combine")
                return False
            
            # Combine all dataframes
            combined_df = pd.concat(combined_data, ignore_index=True)
            
            # Save to output path
            combined_df.to_csv(output_path, index=False)
            
            print(f"Successfully combined {len(file_paths)} files into {output_path}")
            print(f"Total rows: {len(combined_df)}")
            
            return True
            
        except Exception as e:
            print(f"Error combining CSV files: {str(e)}")
            return False
    
    def get_csv_info(self, file_path: str) -> dict:
        """
        Get basic information about a CSV file.
        
        Args:
            file_path: Path to the CSV file
            
        Returns:
            dict: Dictionary containing file information
        """
        try:
            df = pd.read_csv(file_path)
            return {
                'rows': len(df),
                'columns': len(df.columns),
                'column_names': list(df.columns),
                'file_size': os.path.getsize(file_path)
            }
        except Exception as e:
            return {'error': str(e)}
