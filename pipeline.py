import os
import sys
import requests
import psycopg2
from psycopg2.extras import execute_values
from dotenv import load_dotenv
from datetime import datetime
 
load_dotenv()
 
# Microsoft Keys
TENANT_ID = os.getenv("TENANT_ID")
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
 
# Postgres Keys
DB_HOST = os.getenv("DB_HOST")
DB_NAME = os.getenv("DB_NAME")
DB_USER = os.getenv("DB_USER")
DB_PASSWORD = os.getenv("DB_PASSWORD")
DB_PORT = os.getenv("DB_PORT")
 
def get_access_token():
    url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
    data = {
        "grant_type": "client_credentials",
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "scope": "https://graph.microsoft.com/.default"
    }
    response = requests.post(url, data=data)
    response.raise_for_status() 
    return response.json()["access_token"]
 
def fetch_graph_data(endpoint, token):
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    url = f"https://graph.microsoft.com/v1.0/{endpoint}"
    results = []
    while url:
        response = requests.get(url, headers=headers)
        response.raise_for_status()
        data = response.json()
        results.extend(data.get("value", []))
        url = data.get("@odata.nextLink") 
    return results
 
def save_to_postgres(users, simulations):
    """Connects to Postgres and inserts/updates extracted data"""
    conn = psycopg2.connect(
        host=DB_HOST, database=DB_NAME, user=DB_USER, password=DB_PASSWORD, port=DB_PORT
    )
    cur = conn.cursor()
    now = datetime.utcnow()
 
    # 1. Map Users Data
    user_records = [
        (u.get('id'), u.get('displayName'), u.get('userPrincipalName'))
        for u in users if u.get('id')
    ]
    user_query = """
        INSERT INTO microsoft_users (id, display_name, user_principal_name)
        VALUES %s
        ON CONFLICT (id) DO UPDATE SET 
            display_name = EXCLUDED.display_name,
            user_principal_name = EXCLUDED.user_principal_name,
            updated_at = NOW();
    """
    execute_values(cur, user_query, user_records)
 
    # 2. Map Simulations Data
    sim_records = [
        (s.get('id'), s.get('displayName'), s.get('status'), s.get('attackType'), s.get('createdDateTime'))
        for s in simulations if s.get('id')
    ]
    sim_query = """
        INSERT INTO attack_simulations (id, display_name, status, attack_type, created_date_time)
        VALUES %s
        ON CONFLICT (id) DO UPDATE SET 
            display_name = EXCLUDED.display_name,
            status = EXCLUDED.status,
            attack_type = EXCLUDED.attack_type,
            created_date_time = EXCLUDED.created_date_time,
            updated_at = NOW();
    """
    execute_values(cur, sim_query, sim_records)
 
    conn.commit()
    cur.close()
    conn.close()
    print("🐘 Successfully pushed all records cleanly into PostgreSQL database!")
 
if __name__ == "__main__":
    print("🚀 Starting Microsoft Graph API Pipeline via uv environment...")
    try:
        print("Connecting to Microsoft...")
        token = get_access_token()
        print("✅ Connection Successful!")
        users = fetch_graph_data("users", token)
        print(f"👥 Successfully read {len(users)} users.")
        simulations = fetch_graph_data("security/attackSimulation/simulations", token)
        print(f"🎯 Successfully read {len(simulations)} simulation records.")
        print("Pushing data into database...")
        save_to_postgres(users, simulations)
    except Exception as e:
        print(f"❌ Pipeline failed: {e}")