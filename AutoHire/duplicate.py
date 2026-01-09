from pymongo import MongoClient
from datetime import datetime, timedelta

client = MongoClient("mongodb://localhost:27017/")
db = client["resume_db"]
collection = db["jd_evaluation"]
duplicate_collection = db["duplicate_candidates"]

def check_and_insert(data):
    email = data.get("email", "").lower()
    phone = data.get("phone", "")
    timestamp = datetime.utcnow()

    # Check for recent email duplicates (within 6 months)
    existing_email = collection.find_one({
        "email": email,
        "timestamp": {"$gte": timestamp - timedelta(days=180)}
    })
    if existing_email:
        print("⛔ Duplicate (Email within 6 months)")
        data["duplicate_reason"] = "Email Duplicate"
        data["timestamp"] = timestamp
        data["is_duplicate"] = True
        duplicate_collection.insert_one(data)
        return "Duplicate"

    # Check phone only if valid
    if phone and phone.strip().lower() != "not found":
        existing_phone = collection.find_one({"phone": phone})
        if existing_phone:
            print("⛔ Duplicate (Phone matched)")
            data["duplicate_reason"] = "Phone Duplicate"
            data["timestamp"] = timestamp
            data["is_duplicate"] = True
            duplicate_collection.insert_one(data)
            return "Duplicate"

    # No duplicates found
    data["timestamp"] = timestamp
    data["is_duplicate"] = False
    return "Insert"
