from fastapi import FastAPI, HTTPException, Query
from utils import load_excel_sheets, get_row_names, calculate_row_sum

app = FastAPI()

# Load structured data at startup
EXCEL_PATH = "Data/capbudg.xlsx"
excel_data = load_excel_sheets(EXCEL_PATH)


@app.get("/list_tables")
def list_tables():
    """Lists all logical tables parsed from CapBudgWS"""
    return {"tables": list(excel_data.keys())}


@app.get("/get_table_details")
def get_table_details(table_name: str = Query(...)):
    """Lists all row names (1st column values) for the specified logical table"""
    if table_name not in excel_data:
        raise HTTPException(status_code=404, detail="Table not found")

    row_names = get_row_names(excel_data[table_name])
    return {
        "table_name": table_name,
        "row_names": row_names
    }


@app.get("/row_sum")
def row_sum(table_name: str = Query(...), row_name: str = Query(...)):
    """Returns sum of numeric values in a specified row of a logical table"""
    if table_name not in excel_data:
        raise HTTPException(status_code=404, detail="Table not found")

    total = calculate_row_sum(excel_data[table_name], row_name)
    if total is None:
        raise HTTPException(status_code=404, detail="Row not found or contains no numeric data")

    return {
        "table_name": table_name,
        "row_name": row_name,
        "sum": total
    }


# deploy the app using  - "uvicorn main:app --reload --port 9090"

""" 
https://localhost:9090/{ API Endpoint }                                 | Method | Description
'/list_tables'                                                          | GET    | Lists table/sheet names from Excel
'/get_table_details?table_name=Initial Investment'                      | GET    | Gets key details for given table name
'/row_sum?table_name=Initial Investment&row_name=Tax Credit (if any )=' | GET    | Gets sum of a table's row 
"""