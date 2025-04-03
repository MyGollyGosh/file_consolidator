import os
import pandas as pd

class Consolidator:
    def __init__(self, input_directory, output_file='directory/for/file/to/be/saved/to'):
        self.input_directory = input_directory
        self.output_file = output_file
        self.all_data = []
        self.files_processed = 0

    def check_files_and_append(self):
        required_columns = ['EAN', 'QTY', 'EUR', 'USD', 'GBP', 'DESCRIPTION', 'FORMAT', 'source_file']

        for filename in os.listdir(self.input_directory):
            if filename.endswith(('.xlsx', '.xls', '.xlsm', '.xlsb', '.odf', '.ods', '.odt')):
                try:
                    file_path = os.path.join(self.input_directory, filename)

                    df = pd.read_excel(file_path)

                    df['source_file'] = filename

                    df = df.reindex(columns=required_columns, fill_value="N/A")

                    self.all_data.append(df)
                    self.files_processed += 1

                except Exception as e:
                    print(f"Error processing {filename}: {str(e)}")

    def consolidate(self):
        if self.all_data:
            master_df = pd.concat(self.all_data, ignore_index=True)
            master_df.to_excel(self.output_file, index=False)

            print(f"Processed {self.files_processed} files.")
            print(f"Master spreadsheet saved as {self.output_file}")

            return master_df
        else:
            print("No valid Excel files found.")
            return None

if __name__ == "__main__":
    consolidator_obj = Consolidator(r'your/file/path')
    consolidator_obj.check_files_and_append()
    consolidator_obj.consolidate()
