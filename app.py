import os
import re
import json
import hashlib
from flask import Flask, render_template, request, jsonify
from docx import Document
from google import genai

app = Flask(__name__)

DB_FILE = "questions.json"
RAW_NOTES_DIR = "raw_notes"
CONFIG_FILE = "config.json"

os.makedirs(RAW_NOTES_DIR, exist_ok=True)
os.makedirs("templates", exist_ok=True)

def load_dotenv():
    """Loads key=value pairs from a local .env file into os.environ if present."""
    env_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env")
    if os.path.exists(env_path):
        try:
            with open(env_path, "r", encoding="utf-8") as f:
                for line in f:
                    line = line.strip()
                    if line and not line.startswith("#") and "=" in line:
                        k, v = line.split("=", 1)
                        os.environ.setdefault(k.strip(), v.strip().strip('"').strip("'"))
        except Exception:
            pass

load_dotenv()


def normalize_text(text):
    """Normalizes question text to create a unique fingerprint for deduplication."""
    return re.sub(r'[^a-z0-9]', '', text.lower())


def parse_raw_questions(raw_text, source_name="Manual Import"):
    """
    Parses raw text in the format of CEH V13 NOTE1 / NOTE2:
    1. Question text...
    A. Option A
    B. Option B
    C. Option C
    D. Option D
    Correct Answer: X Explanation: ...
    """
    # Clean escaped characters if pasted from markdown
    text = raw_text.replace("\\.", ".").replace("\\<", "<").replace("\\>", ">").replace("\\_", "_").replace("\\*", "*")
    text = re.sub(r'You have not answered this question', '', text, flags=re.IGNORECASE)

    # Regex to match each question block
    pattern = re.compile(
        r'(?:^|\n)\s*\d+\.\s*(.*?)\s*'
        r'\n\s*A\.\s*(.*?)\s*'
        r'\n\s*B\.\s*(.*?)\s*'
        r'\n\s*C\.\s*(.*?)\s*'
        r'\n\s*D\.\s*(.*?)\s*'
        r'\n\s*Correct Answer:\s*([A-D])(?:\s*[—\-].*?)?\s*Explanation:\s*(.*?)(?=\n\s*\d+\.\s+|\Z)',
        re.DOTALL
    )

    parsed_list = []
    for match in pattern.finditer(text):
        q_text, opt_a, opt_b, opt_c, opt_d, ans, exp = match.groups()
        q_clean = re.sub(r'\s+', ' ', q_text).strip()
        q_hash = hashlib.md5(normalize_text(q_clean).encode('utf-8')).hexdigest()[:10]

        parsed_list.append({
            "id": f"q-{q_hash}",
            "question": q_clean,
            "options": {
                "A": re.sub(r'\s+', ' ', opt_a).strip(),
                "B": re.sub(r'\s+', ' ', opt_b).strip(),
                "C": re.sub(r'\s+', ' ', opt_c).strip(),
                "D": re.sub(r'\s+', ' ', opt_d).strip()
            },
            "correctAnswer": ans.strip().upper(),
            "explanation": re.sub(r'\s+', ' ', exp).strip(),
            "source": source_name,
            "sources": [source_name]
        })

    return parsed_list


def load_db():
    if not os.path.exists(DB_FILE):
        return []
    with open(DB_FILE, "r", encoding="utf-8") as f:
        try:
            return json.load(f)
        except json.JSONDecodeError:
            return []


def save_db(questions):
    with open(DB_FILE, "w", encoding="utf-8") as f:
        json.dump(questions, f, indent=2, ensure_ascii=False)


def sync_raw_notes_folder():
    """Reads all .docx and .txt files in raw_notes/ and merges unique questions into questions.json,
    tracking all sources (e.g., NOTE1, NOTE2, NOTE3) for each question."""
    existing_questions = load_db()
    q_map = {normalize_text(q["question"]): q for q in existing_questions}
    added_count = 0
    updated_sources = False

    # Ensure all existing questions have a valid sources list
    for q in existing_questions:
        if "sources" not in q or not isinstance(q["sources"], list):
            q["sources"] = [q["source"]] if q.get("source") else []

    if os.path.exists(RAW_NOTES_DIR):
        for filename in sorted(os.listdir(RAW_NOTES_DIR)):
            filepath = os.path.join(RAW_NOTES_DIR, filename)
            raw_text = ""

            if filename.lower().endswith(".docx"):
                doc = Document(filepath)
                raw_text = "\n".join([p.text for p in doc.paragraphs if p.text.strip()])
            elif filename.lower().endswith(".txt"):
                with open(filepath, "r", encoding="utf-8") as f:
                    raw_text = f.read()

            if raw_text:
                extracted = parse_raw_questions(raw_text, source_name=filename)
                for q in extracted:
                    norm_q = normalize_text(q["question"])
                    if norm_q in q_map:
                        target = q_map[norm_q]
                        if "sources" not in target or not isinstance(target["sources"], list):
                            target["sources"] = [target["source"]] if target.get("source") else []
                        if filename not in target["sources"]:
                            target["sources"].append(filename)
                            updated_sources = True
                    else:
                        q["sources"] = [filename]
                        existing_questions.append(q)
                        q_map[norm_q] = q
                        added_count += 1

    if added_count > 0 or updated_sources or not os.path.exists(DB_FILE):
        save_db(existing_questions)
    return len(existing_questions), added_count


def get_api_key():
    # Priority 1: Environment variable (standard for cloud hosting like Render)
    env_key = os.environ.get("GEMINI_API_KEY", "").strip()
    if env_key:
        return env_key

    # Priority 2: Local config.json file (for local development fallback)
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                return json.load(f).get("gemini_api_key", "").strip()
        except Exception:
            pass
    return ""


@app.route("/")
def index():
    sync_raw_notes_folder()
    return render_template("index.html")


@app.route("/api/questions", methods=["GET"])
def get_questions():
    source_filter = request.args.get("source", "").strip()
    questions = load_db()

    # Calculate note statistics across all sources
    notes_counter = {}
    for q in questions:
        q_sources = q.get("sources") or ([q["source"]] if q.get("source") else ["Manual Entry"])
        for s in q_sources:
            notes_counter[s] = notes_counter.get(s, 0) + 1

    notes_list = []
    # Natural sort order for notes (e.g. NOTE 1, NOTE 2, NOTE 3)
    def natural_sort_key(s):
        return [int(text) if text.isdigit() else text.lower() for text in re.split(r'(\d+)', s)]

    for s_name in sorted(notes_counter.keys(), key=natural_sort_key):
        notes_list.append({
            "id": s_name,
            "filename": s_name,
            "count": notes_counter[s_name]
        })

    if source_filter and source_filter.lower() != "all":
        filtered = []
        for q in questions:
            q_sources = q.get("sources") or ([q["source"]] if q.get("source") else [])
            if any(s.lower() == source_filter.lower() for s in q_sources):
                filtered.append(q)
        return jsonify({
            "total": len(filtered),
            "questions": filtered,
            "all_notes": notes_list,
            "all_total": len(questions)
        })

    return jsonify({
        "total": len(questions),
        "questions": questions,
        "all_notes": notes_list
    })


@app.route("/api/questions/add-single", methods=["POST"])
def add_single_question():
    data = request.json
    questions = load_db()
    seen_hashes = {normalize_text(q["question"]) for q in questions}

    q_clean = data.get("question", "").strip()
    if not q_clean:
        return jsonify({"error": "Question text is required"}), 400

    norm_q = normalize_text(q_clean)
    if norm_q in seen_hashes:
        return jsonify({"error": "Duplicate question already exists in database!"}), 409

    q_hash = hashlib.md5(norm_q.encode('utf-8')).hexdigest()[:10]
    src = data.get("source", "Manual Form Entry").strip() or "Manual Form Entry"
    new_q = {
        "id": f"q-{q_hash}",
        "question": q_clean,
        "options": {
            "A": data["options"]["A"].strip(),
            "B": data["options"]["B"].strip(),
            "C": data["options"]["C"].strip(),
            "D": data["options"]["D"].strip(),
        },
        "correctAnswer": data["correctAnswer"].upper(),
        "explanation": data.get("explanation", "").strip(),
        "source": src,
        "sources": [src]
    }
    questions.append(new_q)
    save_db(questions)
    return jsonify({"message": "Question saved to disk!", "total": len(questions), "added": 1})


@app.route("/api/questions/bulk-import", methods=["POST"])
def bulk_import_questions():
    data = request.json
    raw_text = data.get("rawText", "")
    source_name = data.get("source", "Bulk Paste").strip() or "Bulk Paste"

    parsed = parse_raw_questions(raw_text, source_name)
    if not parsed:
        return jsonify({"error": "Could not parse any questions. Check the formatting."}), 400

    questions = load_db()
    q_map = {normalize_text(q["question"]): q for q in questions}

    added = 0
    duplicates = 0
    for q in parsed:
        norm_q = normalize_text(q["question"])
        if norm_q not in q_map:
            q["sources"] = [source_name]
            questions.append(q)
            q_map[norm_q] = q
            added += 1
        else:
            target = q_map[norm_q]
            if "sources" not in target or not isinstance(target["sources"], list):
                target["sources"] = [target.get("source")] if target.get("source") else []
            if source_name not in target["sources"]:
                target["sources"].append(source_name)
            duplicates += 1

    save_db(questions)
    return jsonify({
        "message": f"Added {added} new questions ({duplicates} duplicates mapped to source).",
        "added": added,
        "duplicates": duplicates,
        "total": len(questions)
    })


@app.route("/api/config/apikey", methods=["POST"])
def save_api_key():
    api_key = request.json.get("apiKey", "").strip()
    os.environ["GEMINI_API_KEY"] = api_key
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump({"gemini_api_key": api_key}, f)
    except Exception:
        pass
    return jsonify({"message": "Gemini API Key updated successfully!"})


SUPPORTED_MODELS = [
    "gemini-3.8-flash",
    "gemini-3.7-flash",
    "gemini-3.5-flash"
]


def generate_with_fallback(client, contents, preferred_model="auto"):
    """
    Tries the preferred model first, and automatically falls back to other Flash models
    if quota (429), high demand (503), or availability issues arise.
    """
    if preferred_model and preferred_model != "auto" and preferred_model in SUPPORTED_MODELS:
        models_to_try = [preferred_model] + [m for m in SUPPORTED_MODELS if m != preferred_model]
    else:
        models_to_try = list(SUPPORTED_MODELS)

    errors = []
    for model_name in models_to_try:
        try:
            print(f"[*] Calling model: {model_name}...")
            response = client.models.generate_content(
                model=model_name,
                contents=contents
            )
            return response.text, model_name
        except Exception as e:
            err_msg = str(e)
            print(f"[!] Model {model_name} failed: {err_msg}. Trying next fallback...")
            errors.append(f"{model_name}: {err_msg}")
            continue

    raise RuntimeError("All available Gemini models reached their limit:\n" + "\n".join(errors))


@app.route("/api/ai-explain", methods=["POST"])
def ai_explain():
    data = request.json
    api_key = get_api_key()
    if not api_key:
        return jsonify({"error": "Gemini API key not configured. Click 'API Key Settings' in the top bar first."}), 400

    q = data.get("question")
    user_choice = data.get("userChoice", "None")
    custom_followup = data.get("followUp", "")
    preferred_model = data.get("model", "auto")

    prompt = f"""You are an expert cybersecurity instructor and CEH v13 mentor.
Break down this exam question clearly and concisely:

Question: {q['question']}
Options:
A. {q['options']['A']}
B. {q['options']['B']}
C. {q['options']['C']}
D. {q['options']['D']}

Student's Selected Answer: {user_choice}
Correct Answer: {q['correctAnswer']}
Official Note Explanation: {q['explanation']}
"""
    if custom_followup:
        prompt += f"\nStudent's Follow-up Question: {custom_followup}\nAnswer the student's follow-up directly and technically."
    else:
        prompt += """
Provide a structured breakdown with:
1. **Core Technical Concept**: Explain the tool, protocol, or attack vector in plain, practical cybersecurity terms.
2. **Why the Correct Answer Wins**: Why the correct option is the best technical fit.
3. **Why Your Choice Was Wrong** (if the student picked the wrong option): Specifically contrast their choice with the right one.
4. **Real-World / CLI Example**: Show a brief command or defense configuration example if applicable.
"""

    try:
        client = genai.Client(api_key=api_key)
        explanation, used_model = generate_with_fallback(client, prompt, preferred_model=preferred_model)
        return jsonify({
            "explanation": explanation,
            "modelUsed": used_model
        })
    except Exception as e:
        return jsonify({"error": f"Gemini API Error: {str(e)}"}), 500


@app.route("/api/transcribe", methods=["POST"])
def transcribe_audio():
    if "audio" not in request.files:
        return jsonify({"error": "No audio file provided in request."}), 400

    api_key = get_api_key()
    if not api_key:
        return jsonify({"error": "Gemini API key not configured."}), 400

    audio_file = request.files["audio"]
    audio_bytes = audio_file.read()
    if not audio_bytes:
        return jsonify({"error": "Empty audio data received."}), 400

    preferred_model = request.form.get("model", "auto")

    mime_type = audio_file.content_type or "audio/webm"
    if "webm" in mime_type:
        mime_type = "audio/webm"
    elif "wav" in mime_type:
        mime_type = "audio/wav"
    elif "ogg" in mime_type:
        mime_type = "audio/ogg"
    elif "mp4" in mime_type or "m4a" in mime_type:
        mime_type = "audio/mp4"

    try:
        from google.genai import types
        client = genai.Client(api_key=api_key)
        audio_part = types.Part.from_bytes(data=audio_bytes, mime_type=mime_type)
        transcript, used_model = generate_with_fallback(
            client,
            contents=[
                audio_part,
                "Transcribe this spoken cybersecurity / technical question clearly and accurately into plain text. Output ONLY the raw transcript text with no extra commentary, no quotation marks, and no introductory words."
            ],
            preferred_model=preferred_model
        )
        return jsonify({"transcript": transcript.strip(), "modelUsed": used_model})
    except Exception as e:
        return jsonify({"error": f"Gemini transcription error: {str(e)}"}), 500


if __name__ == "__main__":
    total, newly_added = sync_raw_notes_folder()
    port = int(os.environ.get("PORT", 5000))
    is_production = os.environ.get("RENDER") is not None
    print(f"[*] Database ready: {total} total unique questions ({newly_added} newly imported from raw_notes/)")
    print(f"[*] Server listening on port: {port}")
    app.run(host="0.0.0.0", port=port, debug=not is_production)