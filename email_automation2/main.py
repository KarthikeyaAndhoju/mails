from fastapi import FastAPI, HTTPException
from pydantic import BaseModel
import pandas as pd

app = FastAPI()

# Load data from Excel file
excel_file_path = "C:/Users/karthikeya.andhoju/Desktop/email_automation/emails.xlsx"
df = pd.read_excel(excel_file_path)

# Convert dataframe to dictionary for easier access
database = df.set_index('case_id').to_dict()['email']


class Case(BaseModel):
    email: str


@app.get('/case/{case_id}')
def get_case(case_id: int):
    if case_id in database:
        return {'case_id': case_id, 'email': database[case_id]}
    else:
        raise HTTPException(status_code=404, detail="Case ID not found")


@app.post('/case/{case_id}')
def create_case(case_id: int, case: Case):
    if case_id in database:
        raise HTTPException(status_code=400, detail="Case ID already exists")
    database[case_id] = case.email
    # Append data to the Excel file
    df.loc[len(df)] = [case_id, case.email]
    df.to_excel(excel_file_path, index=False)
    return {'case_id': case_id, 'email': case.email}


@app.put('/case/{case_id}')
def update_case(case_id: int, case: Case):
    if case_id not in database:
        raise HTTPException(status_code=404, detail="Case ID not found")
    database[case_id] = case.email
    # Update data in the Excel file
    df.loc[df['case_id'] == case_id, 'email'] = case.email
    df.to_excel(excel_file_path, index=False)
    return {'case_id': case_id, 'email': case.email}


@app.delete('/case/{case_id}')
def delete_case(case_id: int):
    if case_id not in database:
        raise HTTPException(status_code=404, detail="Case ID not found")
    del database[case_id]
    # Update data in the Excel file
    df.drop(df[df['case_id'] == case_id].index, inplace=True)
    df.to_excel(excel_file_path, index=False)
    return {'message': 'Case deleted'}
