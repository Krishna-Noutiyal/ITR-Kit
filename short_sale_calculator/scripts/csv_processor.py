import pandas as pd
import os
from typing import List
import glob

class CSVProcessor:
    def __init__(self):
        self.headers =[
                "Security Name (Security Code)",
                "Date of Sale/Transfer",
                "Asset Type",
                "Quantity",
                "Sales Consideration - Reported by Source",
                "Cost of Acquisition",
                "Status"
                
            ]

    
    def combine_csvs(self, file_paths: List[str], output_path: str) -> bool:
        """
        Combines multiple CSV files into a single CSV file.
        Skips the first row and uses the second row as headers.
        
        Args:
            file_paths: List of paths to CSV files to combine
            output_path: Path where the combined CSV should be saved
            
        Returns:
            bool: True if successful, False otherwise
        """
        try:
            combined_data = []
            
            for i, file_path in enumerate(file_paths):
                if not os.path.exists(file_path):
                    print(f"Warning: File {file_path} does not exist, skipping...")
                    continue
                
                # Read the CSV file, skip the first row and use the second row as header
                df = pd.read_csv(file_path,skiprows=1)
                
                # Set the column names to match the headers from the first file
                df = df[self.headers]
                
                
                # Convert columns to numeric, removing commas and handling errors
                df['Cost of Acquisition'] = pd.to_numeric(df['Cost of Acquisition'].replace({r',': ''}, regex=True))
                df['Sales Consideration - Reported by Source'] = pd.to_numeric(df['Sales Consideration - Reported by Source'].replace({r',': ''}, regex=True), errors='coerce')

                # Add source file column to track origin
                df['source_file'] = os.path.basename(file_path)
                df['Cost - Sell'] = df['Sales Consideration - Reported by Source'] - df['Cost of Acquisition']
                
                # Clean up any empty rows
                df = df.dropna(how='all')
                
                combined_data.append(df)
                print(f"Processed {os.path.basename(file_path)}: {len(df)} data rows")
            
            if not combined_data:
                print("No valid CSV files found to combine")
                return False
            
            # Combine all dataframes
            combined_df = pd.concat(combined_data, ignore_index=True)
            print(combined_df)
            
            # Clean up the combined data
            combined_df = combined_df.dropna(how='all')  # Remove completely empty rows
            
            # Save to output path
            combined_df.to_csv(output_path, index=False)
            
            print(f"Successfully combined {len(file_paths)} files into {output_path}")
            print(f"Total data rows: {len(combined_df)}")
            print(f"Total columns: {len(combined_df.columns)}")
            
            return True
            
        except Exception as e:
            print(f"Error combining CSV files: {str(e)}")
            return False
    
    def get_csv_info(self, file_path: str) -> dict:
        """
        Get basic information about a CSV file, skipping the first row and using the second row as headers.
        
        Args:
            file_path: Path to the CSV file
            
        Returns:
            dict: Dictionary containing file information
        """
        try:
            df = pd.read_csv(file_path, header=None, skiprows=[0], names=self.headers)
            df = df.dropna(how='all')
            return {
                'rows': len(df),
                'columns': len(df.columns),
                'column_names': list(df.columns),
                'file_size': os.path.getsize(file_path)
            }
        except Exception as e:
            return {'error': str(e)}



if __name__ == "__main__":
    test = CSVProcessor()

    # List all CSV files in the /test folder
    test_folder = "./short_sale_calculator/test"
    file_list = glob.glob(os.path.join(test_folder, "*.csv"))
    print("Files found:", file_list)

    # Example usage of combine_csvs with the found files
    test.combine_csvs(file_list, "./short_sale_calculator/test/combined.csv")