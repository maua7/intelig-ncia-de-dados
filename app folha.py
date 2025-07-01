import pandas as pd
from datetime import timedelta

def clean_value(value):
    """Clean and convert values, handling various time and monetary formats."""
    if pd.isna(value) or value == '':
        return 0.0
    
    # Handle timedelta (worked hours)
    if isinstance(value, timedelta):
        hours = int(value.total_seconds() // 3600)
        minutes = int((value.total_seconds() % 3600) // 60)
        return f"{hours}.{minutes:02d}"
    
    # Handle string values
    if isinstance(value, str):
        # Clean and strip the value
        value = value.strip()
        
        # Time format (HH:MM or HH:MM:SS)
        if ":" in value:
            try:
                parts = value.split(":")
                hours = int(parts[0])
                minutes = int(parts[1])
                # Format as H.MM (e.g., 2.09 instead of 2.15)
                return f"{hours}.{minutes:02d}"
            except:
                return 0.0
                
        # Monetary values
        if value.startswith('R$'):
            try:
                # Remove R$, replace comma with dot, remove thousands separator
                value = value.replace('R$', '').replace('.', '').replace(',', '.').strip()
                return round(float(value), 2)
            except:
                return 0.0
                
        # Percentage format
        if value.endswith('%'):
            # Remove % symbol
            return float(value.replace('%', ''))
            
        try:
            # Round to 2 decimal places
            return round(float(value), 2)
        except:
            return 0.0
            
    # Handle numeric values
    try:
        # Round to 2 decimal places
        return round(float(value), 2)
    except:
        return 0.0

def process_payroll_excel(file_path):
    # Read Excel file
    df = pd.read_excel(file_path, header=None)
    insert_statements = []
    
    # não esqueça burro 
    column_to_event = {
        2: 2,
        3: 957,
        4: 64 ,
        5: 3,
        6: 225  }
    
    # Process each row
    for idx in range(1, len(df)):  # Start from row 1 (skip header)
        row = df.iloc[idx]
        
        # Check if it's a valid row (with registration number)
        if pd.notna(row[1]) and str(row[1]).strip().isdigit():
            matricula = int(row[1])
            
            # Process each mapped column
            for col, event_code in column_to_event.items():
                value = row[col]
                
                # Skip if value is NaN or empty
                if pd.isna(value) or value == '' or value == '-':
                    continue
                    
                cleaned_value = clean_value(value)
                
                # Only insert if value is not zero
                if cleaned_value != 0.0:
                    insert_statements.append(
                    f"""INSERT INTO movevento 
                    (cd_empresa, mes, ano, cd_funcionario, cd_evento, referencia, transferido, tipo_processamento, origem_digitacao)
                    VALUES (279, 6, 2025, {matricula}, {event_code}, {cleaned_value}, '', 2, 'M')
                    ON DUPLICATE KEY UPDATE 
                        referencia = {cleaned_value},
                        transferido = '',
                        tipo_processamento = 2,
                        origem_digitacao = 'M';"""
                )
    
    return insert_statements

# Usage
file_path = "C:\\Users\\Micro\\Documents\\GitHub\\intelig-ncia-de-dados\\Nova Planilha DP.xlsx"
insert_statements = process_payroll_excel(file_path)

# Print the INSERTs
for statement in insert_statements:
    print(statement)