import os
import pandas as pd

def get_file_info(path):
    try:
        if path.endswith('.csv'):
            # Strict standard separation parsing
            df = pd.read_csv(path, nrows=2, sep=',', engine='python', encoding='utf-8-sig')
            
            # Fallback check: Agar abhi bhi 1 column dikha raha hai toh manual line split karenge
            if df.shape[1] == 1:
                with open(path, 'r', encoding='utf-8', errors='ignore') as f:
                    first_line = f.readline().strip()
                    cols_count = len(first_line.split(','))
                    if cols_count > 1:
                        return {
                            "rows": sum(1 for _ in open(path, encoding='utf-8', errors='ignore')) - 1,
                            "cols": cols_count,
                            "size": f"{(os.path.getsize(path)/1024):.2f} KB"
                        }
        else:
            df = pd.read_excel(path, nrows=2, engine='openpyxl')

        total_rows = 0
        if path.endswith('.csv'):
            total_rows = sum(1 for _ in open(path, encoding='utf-8', errors='ignore')) - 1
        else:
            total_rows = pd.read_excel(path, engine='openpyxl').shape[0]

        return {
            "rows": total_rows if total_rows >= 0 else 0,
            "cols": df.shape[1],
            "size": f"{(os.path.getsize(path)/1024):.2f} KB"
        }
    except Exception as e:
        print(f"[DEBUG ERROR] Info Fetch Failure: {str(e)}")
        return {"rows": 0, "cols": 0, "size": "0.00 KB"}

def process_and_convert(input_path, output_dir, mode, config, progress_callback):
    progress_callback(20)
    
    if mode == "CSV to Excel":
        df = pd.read_csv(input_path, sep=',', engine='python', encoding='utf-8-sig')
        if config["remove_duplicates"]:
            df.drop_duplicates(inplace=True)
        if config["handle_missing"] == "Drop Dropouts":
            df.dropna(inplace=True)
        elif config["handle_missing"] == "Fill with N/A":
            df.fillna("N/A", inplace=True)
            
        progress_callback(60)
        filename = os.path.basename(input_path).replace('.csv', '.xlsx')
        out_file = os.path.join(output_dir, filename)
        
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
        df.to_excel(out_file, index=False, engine='openpyxl')
    else:
        df = pd.read_excel(input_path, engine='openpyxl')
        if config["remove_duplicates"]:
            df.drop_duplicates(inplace=True)
        if config["handle_missing"] == "Drop Dropouts":
            df.dropna(inplace=True)
        elif config["handle_missing"] == "Fill with N/A":
            df.fillna("N/A", inplace=True)
            
        progress_callback(60)
        filename = os.path.basename(input_path).replace('.xlsx', '.csv')
        out_file = os.path.join(output_dir, filename)
        
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
        df.to_csv(out_file, index=False, encoding='utf-8-sig')

    progress_callback(100)
    return out_file