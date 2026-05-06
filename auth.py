from supabase_client import supabase


def sign_up(email, password, full_name, role):
    try:
        res = supabase.auth.sign_up({
            "email": email,
            "password": password,
            "options": {
                "data": {
                    "full_name": full_name,
                    "role": role
                }
            }
        })
        return True, res.user
    except Exception as e:
        return False, str(e)


def sign_in(email, password):
    try:
        res = supabase.auth.sign_in_with_password({
            "email": email,
            "password": password
        })
        return True, res.user
    except Exception as e:
        return False, str(e)


def sign_out():
    try:
        supabase.auth.sign_out()
    except Exception:
        pass


# ── Lab Settings ─────────────────────────────────────

def save_lab_settings(user_id, lab):
    try:
        supabase.table("lab_settings").upsert({
            "user_id": user_id,
            "name":          lab.get("name", ""),
            "address":       lab.get("address", ""),
            "phone":         lab.get("phone", ""),
            "accreditation": lab.get("accreditation", ""),
            "technician":    lab.get("technician", ""),
        }).execute()
    except Exception as e:
        print("save_lab_settings error:", e)


def load_lab_settings(user_id):
    try:
        res = supabase.table("lab_settings")\
            .select("*")\
            .eq("user_id", user_id)\
            .single()\
            .execute()
        return res.data
    except Exception:
        return None


# ── Patients ─────────────────────────────────────────

def save_patient(user_id, patient):
    try:
        supabase.table("patients").insert({
            "user_id":          user_id,
            "full_name":        patient.get("name", ""),
            "patient_id":       patient.get("pid", ""),
            "date_of_birth":    patient.get("dob", ""),
            "age":              patient.get("age", ""),
            "gender":           patient.get("gender", "Male"),
            "mobile":           patient.get("mobile", ""),
            "email":            patient.get("email", ""),
            "referring_doctor": patient.get("doctor", ""),
            "sample_id":        patient.get("sample_id", ""),
            "collection_date":  patient.get("collection_date", ""),
            "clinical_notes":   patient.get("notes", ""),
        }).execute()
    except Exception as e:
        print("save_patient error:", e)


def load_patients(user_id):
    try:
        res = supabase.table("patients")\
            .select("*")\
            .eq("user_id", user_id)\
            .order("created_at", desc=True)\
            .execute()
        return res.data
    except Exception:
        return []


# ── Analyses ─────────────────────────────────────────

def save_analysis(user_id, image_name, rbc, wbc,
                  rbc_flag, wbc_flag, severity,
                  confidence, health_score, ai_insight, notes):
    try:
        supabase.table("analyses").insert({
            "user_id":      user_id,
            "image_name":   image_name,
            "rbc":          rbc,
            "wbc":          wbc,
            "rbc_flag":     rbc_flag,
            "wbc_flag":     wbc_flag,
            "severity":     severity,
            "confidence":   confidence,
            "health_score": health_score,
            "ai_insight":   ai_insight,
            "notes":        notes,
        }).execute()
    except Exception as e:
        print("save_analysis error:", e)


# ── Reports ──────────────────────────────────────────

def save_report(user_id, report_id, patient_name,
                total_rbc, total_wbc, severity,
                health_score, images_count):
    try:
        supabase.table("reports").insert({
            "user_id":      user_id,
            "report_id":    report_id,
            "patient_name": patient_name,
            "total_rbc":    total_rbc,
            "total_wbc":    total_wbc,
            "severity":     severity,
            "health_score": health_score,
            "images_count": images_count,
        }).execute()
    except Exception as e:
        print("save_report error:", e)


def load_reports(user_id):
    try:
        res = supabase.table("reports")\
            .select("*")\
            .eq("user_id", user_id)\
            .order("created_at", desc=True)\
            .execute()
        return res.data
    except Exception:
        return []