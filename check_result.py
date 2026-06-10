import openpyxl

def check_results():
    try:
        wb = openpyxl.load_workbook("Климатрейд_result.xlsx")
        sheet = wb.active
        
        headers = [cell.value for cell in sheet[1]]
        model_col_idx = headers.index('Модель')
        article_col_idx = len(headers) - 1
        if headers[-1] != 'Артикул':
            print("Warning: Last column not named Артикул")
            
        count = 0
        for row in sheet.iter_rows(min_row=2, max_row=10, values_only=True):
            model = row[model_col_idx]
            article = row[article_col_idx]
            if article:
                print(f"Model: {model} -> Article: {article}")
                count += 1
                
        print(f"Total entries shown: {count}")
    except Exception as e:
        print(f"Error checking results: {e}")

if __name__ == "__main__":
    check_results()
