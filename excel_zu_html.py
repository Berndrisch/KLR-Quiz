import openpyxl
import html
import random
import re
import json

# ==========================================
# KONFIGURATION
# ==========================================
# Trage hier deine E-Mail-Adresse ein, an die das Feedback der User geschickt werden soll:
# (Beispiel: "max.mustermann@web.de")
ZIEL_EMAIL = "bernd.rischkau@t-online.de"
# ==========================================

def bereinige_latex(text):
    if not text:
        return ""
    text = str(text)
    
    # 1. Mathematische Zeichen und Prozent-Artefakte bereinigen
    text = text.replace(r'\times', '×')
    text = text.replace(r'\cdot', '·')
    text = text.replace(r'\%', '%')
    
    # 2. NEU: Die chirurgische Reparatur für den Autoren-Tippfehler
    # Sucht nach "\text{...)" und wandelt es in sauberen Text ohne die falsche Klammer um
    text = re.sub(r'\\text\{([^)]*)\)', r'\1', text)
    
    # 3. Der alte "Holzhammer" als Fallback für Reste
    text = text.replace(r'\text{', '')
    text = text.replace(r'\text', '')
    text = text.replace('{', '')
    text = text.replace('}', '')
    
    return text

def generiere_html_aus_excel(excel_datei, html_datei):
    print(f"Öffne Excel-Datei '{excel_datei}'...")
    
    try:
        workbook = openpyxl.load_workbook(excel_datei, data_only=True)
        sheet = workbook.active
    except Exception as e:
        print(f"Kritischer Fehler beim Öffnen: {e}")
        return

    # Teil 1: HTML und CSS
    html_content = """<!DOCTYPE html>
<html lang="de">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Interaktives Multiple-Choice Quiz</title>
    <style>
        :root { --primary: #0056b3; --bg: #f4f7f6; --card-bg: #ffffff; --text: #333; --success: #28a745; --danger: #dc3545; --warning: #ffc107; --secondary: #6c757d; }
        body { font-family: -apple-system, BlinkMacSystemFont, sans-serif; background-color: var(--bg); color: var(--text); margin: 0; line-height: 1.6; }
        .header { background: var(--primary); color: white; padding: 15px 20px; position: sticky; top: 0; z-index: 100; display: flex; justify-content: space-between; align-items: center; box-shadow: 0 2px 5px rgba(0,0,0,0.2); }
        .header-left { display: flex; align-items: center; gap: 20px; }
        .header h1 { margin: 0; font-size: 1.2em; }
        .score-board { font-size: 1.2em; font-weight: bold; background: rgba(255,255,255,0.2); padding: 5px 15px; border-radius: 20px; }
        .btn-header { background: rgba(255,255,255,0.2); color: white; border: 1px solid white; padding: 5px 15px; border-radius: 5px; cursor: pointer; font-weight: bold; transition: background 0.2s; }
        .btn-header:hover { background: rgba(255,255,255,0.4); }
        .container { max-width: 800px; margin: 40px auto; padding: 0 20px; }
        .question-card { display: none; background: var(--card-bg); border-radius: 8px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); padding: 30px; margin-bottom: 30px; position: relative; }
        .question-card.active { display: block; animation: fadeIn 0.3s; }
        .category-badge { position: absolute; top: -15px; left: 30px; color: white; padding: 5px 15px; border-radius: 20px; font-size: 0.9em; font-weight: bold; box-shadow: 0 2px 4px rgba(0,0,0,0.2); cursor: pointer; transition: transform 0.2s, opacity 0.2s; }
        .category-badge:hover { transform: scale(1.05); opacity: 0.9; }
        @keyframes fadeIn { from { opacity: 0; transform: translateY(10px); } to { opacity: 1; transform: translateY(0); } }
        .question-title { font-size: 1.1em; color: var(--secondary); margin-top: 10px; margin-bottom: 15px; border-bottom: 2px solid #eee; padding-bottom: 10px; }
        .question-text { font-size: 1.2em; font-weight: bold; margin-bottom: 25px; }
        .options { margin-bottom: 25px; }
        .option-label { display: block; background: #e9ecef; padding: 15px; margin-bottom: 10px; border-radius: 5px; cursor: pointer; transition: all 0.2s; border: 2px solid transparent; }
        .option-label:hover { background: #dee2e6; }
        .option-label.correct-highlight { background: #d4edda; border: 2px solid var(--success); }
        .option-label.wrong-highlight { background: #f8d7da; border: 2px solid var(--danger); }
        input[type="radio"] { margin-right: 15px; transform: scale(1.3); }
        .controls { display: flex; gap: 10px; flex-wrap: wrap; margin-bottom: 20px; align-items: center; }
        .btn { background: var(--primary); color: white; border: none; padding: 12px 20px; border-radius: 5px; cursor: pointer; font-size: 1em; font-weight: bold; width: auto; text-align: center; }
        .btn-block { width: 100%; margin-bottom: 10px; } /* Neu: Für die Startansicht } */
        .btn:hover { opacity: 0.8; }
        .btn-prev { background: var(--secondary); }
        .btn-hint { background: #17a2b8; }
        .btn-skip { background: var(--warning); color: #333; margin-left: auto; }
        .btn-next { background: var(--primary); display: none; margin-left: auto; }
        @media (max-width: 600px) {
            .controls { gap: 8px; }
            .btn { flex: 1 1 auto; padding: 10px 5px; font-size: 0.85em; }
            .btn-prev { order: 1; min-width: 30%; }
            button[onclick^="openFeedbackModal"] { order: 2; min-width: 30%; }
            button[onclick^="toggleHint"] { order: 3; min-width: 30%; }
            [id^="check_btn_"], .btn-next { order: 4; width: 100%; font-size: 1.1em; padding: 12px; margin-top: 5px; }
            .btn-skip { order: 5; width: 100%; margin-left: 0; background: transparent; border: 2px solid var(--warning); color: #666; margin-top: 0px; }
        }
        .hint-box, .feedback-box { display: none; margin-top: 20px; padding: 15px; border-radius: 5px; border-left: 5px solid; }
        .hint-box { background: #e0f7fa; border-color: #17a2b8; }
        .feedback-box.correct { background: #d4edda; border-color: var(--success); }
        .feedback-box.incorrect { background: #f8d7da; border-color: var(--danger); }
        .modal-overlay { display: none; position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.6); z-index: 999; justify-content: center; align-items: center; }
        .modal-content { background: white; width: 90%; max-width: 800px; max-height: 80vh; overflow-y: auto; border-radius: 8px; padding: 30px; position: relative; box-shadow: 0 10px 25px rgba(0,0,0,0.2); }
        .modal-close { position: absolute; top: 20px; right: 20px; background: var(--danger); color: white; border: none; padding: 8px 15px; border-radius: 5px; cursor: pointer; font-weight: bold; }
        .schema-box { background: #f8f9fa; border: 1px solid #dee2e6; padding: 15px; border-radius: 5px; margin-bottom: 20px; font-family: monospace; font-size: 1.1em; line-height: 1.8; }
        .modal-content h2 { color: var(--primary); margin-top: 0; }
        .modal-content h3 { border-bottom: 2px solid var(--primary); padding-bottom: 5px; margin-top: 25px; }
        .highlight { color: var(--primary); font-weight: bold; }
        .warning { color: var(--danger); font-size: 0.9em; font-weight: bold; }
        
        /* Video Container CSS */
        .video-container { margin-bottom: 20px; overflow: hidden; border-radius: 8px; box-shadow: 0 4px 10px rgba(0,0,0,0.1); width: 100%; display: flex; justify-content: center; background: #000; }
        .video-container video, .video-container iframe { width: 100%; aspect-ratio: 16 / 9; border: none; }
    </style>
</head>
<body>
    <div class="header">
        <div class="header-left">
            <h1 id="progress-text">Lade Quiz...</h1>
            <button class="btn-header" onclick="toggleModal()">📖 Formelsammlung</button>
        </div>
        <div class="score-board">Richtig: <span id="score-right">0</span> | Falsch: <span id="score-wrong">0</span></div>
    </div>
    
    <div class="modal-overlay" id="knowledge-modal" onclick="closeModalOutside(event)">
        <div class="modal-content">
            <button class="modal-close" onclick="toggleModal()">Schließen ✖</button>
            <h2>KLR Formelsammlung & Definitionen</h2>
            
            <h3>Video-Tutorials & Erklärungen</h3>
            <div style="display: flex; flex-direction: column; gap: 40px; margin-bottom: 40px; background: #eee; padding: 20px; border-radius: 8px;">
                <div style="width: 100%;">
                    <p style="font-weight: bold; margin-bottom: 10px;">➔ Zusatzaufträge</p>
                    <div class="video-container">
                        <video controls><source src="zusatzauftraege.mp4" type="video/mp4"></video>
                    </div>
                </div>
                
                <div style="width: 100%;">
                    <p style="font-weight: bold; margin-bottom: 10px;">➔ Abgrenzungsrechnung</p>
                    <div class="video-container">
                        <video controls><source src="abgrenzungsrechnung.mp4" type="video/mp4"></video>
                    </div>
                </div>
            </div>
            
            <h3>1. Komplettes Schema der Zuschlagskalkulation (Vorwärtskalkulation)</h3>
            <div class="schema-box">
                Materialeinzelkosten (MEK)<br>
                + Materialgemeinkosten (MGK)<br>
                <span class="highlight">= Materialkosten (MK)</span><br><br>
                
                Fertigungseinzelkosten (FEK / Fertigungslöhne)<br>
                + Fertigungsgemeinkosten (FGK)<br>
                + Sondereinzelkosten der Fertigung (SEKF)<br>
                <span class="highlight">= Fertigungskosten (FK)</span><br><br>
                
                Materialkosten (MK) + Fertigungskosten (FK)<br>
                <span class="highlight">= Herstellkosten (HK)</span><br><br>
                
                + Verwaltungsgemeinkosten (VwGK)<br>
                + Vertriebsgemeinkosten (VtGK)<br>
                + Sondereinzelkosten des Vertriebs (SEKV)<br>
                <span class="highlight">= Selbstkosten (SK)</span><br><br>
                
                + Gewinnzuschlag <span class="warning">(vom Hundert)</span><br>
                <span class="highlight">= Barverkaufspreis (BVP)</span><br><br>
                
                + Kundenskonto <span class="warning">(Achtung: "im Hundert" rechnen!)</span><br>
                + Vertreterprovision <span class="warning">(Achtung: "im Hundert" rechnen!)</span><br>
                <span class="highlight">= Zielverkaufspreis (ZVP)</span><br><br>
                
                + Kundenrabatt <span class="warning">(Achtung: "im Hundert" rechnen!)</span><br>
                <span class="highlight">= Netto-Listenverkaufspreis (LVP)</span><br><br>
                
                + Umsatzsteuer <span class="warning">(vom Hundert)</span><br>
                <span class="highlight">= Brutto-Verkaufspreis</span>
            </div>

            <h3>2. Deckungsbeitragsrechnung</h3>
            <div class="schema-box">
                Stückdeckungsbeitrag (db) = Nettoverkaufserlös (p) - variable Stückkosten (kv)<br>
                Gesamtdeckungsbeitrag (DB) = db * Menge (x)<br>
                Break-Even-Point (Menge) = Fixkosten (Kf) / Stückdeckungsbeitrag (db)
            </div>
            
            <h3>3. Wichtige Definitionen</h3>
            <ul>
                <li><strong>Sondereinzelkosten:</strong> Auftragsbezogene Spezialkosten (z.B. Spezialwerkzeug).</li>
                <li><strong>Grundkosten:</strong> Aufwandsgleiche Kosten.</li>
                <li><strong>Anderskosten:</strong> Aufwandsungleiche Kosten (z.B. kalk. Abschreibungen).</li>
                <li><strong>Zusatzkosten:</strong> Aufwandslose Kosten (z.B. kalk. Unternehmerlohn).</li>
            </ul>
        </div>
    </div>

    <!-- Feedback Modal -->
    <div class="modal-overlay" id="feedback-modal" onclick="closeFeedbackOutside(event)">
        <div class="modal-content" style="max-width: 500px;">
            <button class="modal-close" onclick="closeFeedbackModal()">✖</button>
            <h2 id="feedback-title" style="margin-bottom: 5px;">Feedback zu Frage</h2>
            <form id="feedback-form" onsubmit="submitFeedback(event)">
                <input type="hidden" id="feedback-question-id" name="Frage">
                
                <label style="display: block; margin-bottom: 20px; font-weight: bold;">
                    Deine Frage oder Anmerkung:
                    <textarea name="message" required style="width: 100%; padding: 8px; margin-top: 5px; border: 1px solid #ccc; border-radius: 4px; min-height: 100px;" placeholder="Was ist dir aufgefallen? Gibts ein Problem mit der Lösung?"></textarea>
                </label>
                
                <button type="submit" class="btn btn-block" id="feedback-submit-btn">Absenden 🚀</button>
                <div id="feedback-status" style="margin-top: 15px; font-weight: bold; text-align: center;"></div>
            </form>
        </div>
    </div>

    <div class="container" id="quiz-container">
        <!-- Start Screen -->
        <div class="question-card" id="card_start">
            <h2 style="text-align: center;">KLR Quiz Starten</h2>
            <div style="text-align: center; margin-bottom: 20px;">
                <p id="stats-overview" style="font-size: 1.1em;"></p>
            </div>
            <div style="max-width: 400px; margin: 0 auto; background: #f8f9fa; padding: 20px; border-radius: 8px;">
                <label style="display: block; margin-bottom: 10px; font-weight: bold;">
                    Hauptgebiet wählen:
                    <select id="gebiet-select" style="width: 100%; padding: 8px; margin-top: 5px; border-radius: 4px; border: 1px solid #ccc;">
                        <option value="Alle">Alle Gebiete</option>
                    </select>
                </label>
                <label style="display: block; margin-bottom: 20px; cursor: pointer;">
                    <input type="checkbox" id="randomize-checkbox" checked>
                    Fragen in zufälliger Reihenfolge
                </label>
                
                <button class="btn btn-block" onclick="startNewQuiz()">Neues Quiz starten</button>
                <button class="btn btn-hint btn-block" id="btn-continue" style="display: none;" onclick="continueQuiz()">Letztes Quiz fortsetzen</button>
                <button class="btn btn-block" id="btn-retry-global" style="background: var(--warning); color: #333; display: none;" onclick="retryWrongGlobal()">Falsch beantwortete wiederholen</button>
                <button class="btn btn-danger btn-block" id="btn-reset" style="background: var(--danger); display: none; margin-top: 20px;" onclick="resetQuizData()">Fortschritt & Daten löschen</button>
            </div>
        </div>
"""

    fragen_zaehler = 0
    farbe_mapping = {
        "Teilkostenrechnung": "#6f42c1", "Kostenstellenrechnung": "#fd7e14",
        "Kostenträgerrechnung": "#20c997", "Kostenartenrechnung": "#e83e8c"
    }

    questions_data = []
    unique_gebiete = set()

    # Teil 2: Datenverarbeitung
    for row in sheet.iter_rows(min_row=2, values_only=True):
        if not row[0] or not row[1]:
            continue
            
        fragen_zaehler += 1
        titel = f"Frage {row[0]}"
        text = html.escape(bereinige_latex(row[1]))
        gebiet = str(row[6]).strip() if row[6] else 'Allgemein'
        gebiet_farbe = farbe_mapping.get(gebiet, "#6c757d")
        
        questions_data.append({"id": fragen_zaehler, "gebiet": gebiet})
        unique_gebiete.add(gebiet)
        
        hint_text = html.escape(bereinige_latex(row[7])) if row[7] else "Kein Tipp verfügbar."
        feedback_text = html.escape(bereinige_latex(row[8])) if row[8] else ""
        
        # NEU: Video-Spalte (Spalte J / Index 9)
        video_src = str(row[9]).strip() if len(row) > 9 and row[9] else ""
        video_html = ""
        if video_src:
            # Überprüfen ob es ein YouTube oder lokaler Film ist
            if "youtube.com" in video_src or "youtu.be" in video_src:
                # Einfache YouTube Embed Logik
                yt_id = video_src.split("v=")[1].split("&")[0] if "v=" in video_src else video_src.split("/")[-1]
                video_html = f'<div class="video-container"><iframe src="https://www.youtube.com/embed/{yt_id}" allowfullscreen></iframe></div>'
            else:
                video_html = f'<div class="video-container"><video controls><source src="{video_src}" type="video/mp4">Dein Browser unterstützt das Video-Format nicht.</video></div>'

        answers_raw = [
            (html.escape(bereinige_latex(row[2])) if row[2] else "", "true"),
            (html.escape(bereinige_latex(row[3])) if row[3] else "", "false"),
            (html.escape(bereinige_latex(row[4])) if row[4] else "", "false"),
            (html.escape(bereinige_latex(row[5])) if row[5] else "", "false")
        ]
        answers_valid = [ans for ans in answers_raw if ans[0].strip() != ""]
        random.shuffle(answers_valid)

        html_content += f'''
        <div class="question-card" id="card_{fragen_zaehler}">
            <div class="category-badge" style="background-color: {gebiet_farbe};" onclick="showView('start')" title="Zurück zum Startmenü / Gebiet wechseln">🏠 {gebiet}</div>
            <div class="question-title">{titel}</div>
            {video_html}
            <div class="question-text">{text}</div>
            <div class="options">'''
            
        for i, (ans_text, is_correct) in enumerate(answers_valid):
            ans_id = f"ans_{fragen_zaehler}_{i}"
            html_content += f'''
                <label class="option-label" for="{ans_id}">
                    <input type="radio" name="q_{fragen_zaehler}" id="{ans_id}" value="{is_correct}">
                    {ans_text}
                </label>'''
                
        html_content += f'''
            </div>
            <div class="controls">
                <button class="btn btn-prev" id="prev_btn_{fragen_zaehler}" onclick="prevQuestion()">⬅️ Zurück</button>
                <button class="btn btn-hint" style="background: var(--secondary); font-size: 0.85em; padding: 6px 12px; opacity: 0.8;" onclick="openFeedbackModal({fragen_zaehler})">Feedback</button>
                <button class="btn btn-hint" onclick="toggleHint({fragen_zaehler})">Tipp</button>
                <button class="btn" id="check_btn_{fragen_zaehler}" onclick="checkAnswer({fragen_zaehler})">Antwort prüfen</button>
                <button class="btn btn-skip" id="skip_btn_{fragen_zaehler}" onclick="skipQuestion()">Überspringen ⏭️</button>
                <button class="btn btn-next" id="next_btn_{fragen_zaehler}" onclick="nextQuestion()">Weiter ➔</button>
            </div>
            <div class="hint-box" id="hint_{fragen_zaehler}"><strong>Tipp:</strong> {hint_text}</div>
            <div class="feedback-box" id="feedback_{fragen_zaehler}" data-feedback="{feedback_text}"></div>
        </div>'''

    questions_json = json.dumps(questions_data, ensure_ascii=False)
    gebiete_json = json.dumps(sorted(list(unique_gebiete)), ensure_ascii=False)

    # Teil 3: Endbildschirm und JavaScript
    html_content += f'''
        <div class="question-card" id="card_end">
            <h2 style="text-align: center;">Quiz-Auswertung</h2>
            <div id="end-stats" style="text-align: center; margin-bottom: 30px;"></div>
            <div style="text-align: center; display: flex; flex-direction: column; align-items: center; justify-content: center; gap: 15px; max-width: 400px; margin: 0 auto;">
                <button class="btn btn-block" id="btn-retry-end" style="background: var(--warning); color: #333;" onclick="retryWrongSession()">Falsch beantwortete dieser Sitzung wiederholen</button>
                <button class="btn btn-prev btn-block" onclick="showView('start')">Zum Startmenü 🏠</button>
            </div>
        </div>
    </div>
    <script>
        const allQuestions = {questions_json};
        const allGebiete = {gebiete_json};
        const TARGET_EMAIL = "{ZIEL_EMAIL}";
        
        function openFeedbackModal(qId) {{
            document.getElementById('feedback-title').innerText = "Feedback zu Frage " + qId;
            document.getElementById('feedback-question-id').value = "Feedback/Fehler zu Frage " + qId;
            document.getElementById('feedback-modal').style.display = 'flex';
            document.getElementById('feedback-status').innerText = "";
            document.getElementById('feedback-submit-btn').disabled = false;
        }}
        
        function closeFeedbackModal() {{
            document.getElementById('feedback-modal').style.display = 'none';
        }}
        
        function closeFeedbackOutside(event) {{
            if (event.target.id === 'feedback-modal') {{ closeFeedbackModal(); }}
        }}
        
        function submitFeedback(event) {{
            event.preventDefault();
            const form = event.target;
            const status = document.getElementById('feedback-status');
            const btn = document.getElementById('feedback-submit-btn');
            
            if (TARGET_EMAIL === "deine@email.de" || !TARGET_EMAIL) {{
                status.style.color = "var(--danger)";
                status.innerText = "❌ Fehler: Der Ersteller hat (noch) keine E-Mail-Adresse hinterlegt!";
                return;
            }}
            
            btn.disabled = true;
            status.style.color = "var(--text)";
            status.innerText = "Öffne E-Mail-Programm... 🚀";
            
            const qId = document.getElementById('feedback-question-id').value;
            const msg = form.elements['message'].value;
            
            const subject = encodeURIComponent(qId);
            const body = encodeURIComponent(msg);
            
            window.location.href = `mailto:${{TARGET_EMAIL}}?subject=${{subject}}&body=${{body}}`;
            
            setTimeout(() => {{
                closeFeedbackModal();
                status.innerText = "";
                btn.disabled = false;
                form.reset();
            }}, 2000);
        }}

        let state = {{
            queue: [],
            currentIndex: 0,
            results: {{}},
            scoreRight: 0,
            scoreWrong: 0,
            activeView: 'start'
        }};

        function saveState() {{
            localStorage.setItem('klr_quiz_state', JSON.stringify(state));
        }}

        function loadState() {{
            const s = localStorage.getItem('klr_quiz_state');
            if (s) {{
                try {{
                    state = JSON.parse(s);
                }} catch (e) {{
                    console.error("Fehler beim Laden des State");
                }}
            }}
        }}

        function init() {{
            loadState();
            
            // Populate dropdown
            const select = document.getElementById('gebiet-select');
            allGebiete.forEach(g => {{
                const opt = document.createElement('option');
                opt.value = g;
                opt.innerText = g;
                select.appendChild(opt);
            }});
            
            updateHeaderStats();
            
            const answeredCount = Object.keys(state.results).length;
            if (state.queue && state.queue.length > 0 && state.currentIndex < state.queue.length) {{
                document.getElementById('btn-continue').style.display = 'inline-block';
                document.getElementById('btn-reset').style.display = 'inline-block';
                const total = state.queue.length;
                
                // Zähle korrekte in DIESER Session Queue
                document.getElementById('stats-overview').innerText = `Gestartetes Quiz: ${{answeredCount}} beantwortet. Richtig: ${{state.scoreRight}}, Falsch: ${{state.scoreWrong}}. Noch offen: ${{total - state.currentIndex}}`;
            }} else if (answeredCount > 0) {{
                document.getElementById('btn-reset').style.display = 'inline-block';
                document.getElementById('stats-overview').innerText = `Willkommen zurück! Letzter Stand: ${{state.scoreRight}} Richtig, ${{state.scoreWrong}} Falsch.`;
            }} else {{
                document.getElementById('stats-overview').innerText = `Willkommen zum KLR Quiz! (${{allQuestions.length}} Fragen verfügbar)`;
            }}
            
            // Falsch beantwortete Button anzeigen, wenn es welche gibt
            const totalErrors = Object.values(state.results).filter(x => x === 'incorrect').length;
            if (totalErrors > 0) {{
                document.getElementById('btn-retry-global').style.display = 'inline-block';
            }}
            
            showView('start');
        }}

        function updateHeaderStats() {{
            document.getElementById('score-right').innerText = state.scoreRight || 0;
            document.getElementById('score-wrong').innerText = state.scoreWrong || 0;
        }}

        function showView(viewName) {{
            document.querySelectorAll('.question-card').forEach(card => card.classList.remove('active'));
            state.activeView = viewName;
            
            if (viewName === 'start') {{
                document.getElementById('card_start').classList.add('active');
                document.getElementById('progress-text').innerText = "Startmenü";
                
                // Aktualisiere Startscreen Texte
                const answeredCount = Object.keys(state.results).length;
                const totalErrors = Object.values(state.results).filter(x => x === 'incorrect').length;
                document.getElementById('btn-retry-global').style.display = totalErrors > 0 ? 'inline-block' : 'none';
                
                if (state.queue && state.queue.length > 0 && state.currentIndex < state.queue.length) {{
                    document.getElementById('btn-continue').style.display = 'inline-block';
                    document.getElementById('stats-overview').innerText = `Gestartetes Quiz: ${{answeredCount}} beantwortet. Richtig: ${{state.scoreRight}}, Falsch: ${{state.scoreWrong}}. Noch offen: ${{state.queue.length - state.currentIndex}}`;
                }} else if (answeredCount > 0) {{
                    document.getElementById('btn-continue').style.display = 'none';
                    document.getElementById('stats-overview').innerText = `Willkommen zurück! Letztes Quiz beendet.`;
                }}
                
            }} else if (viewName === 'end') {{
                document.getElementById('card_end').classList.add('active');
                document.getElementById('progress-text').innerText = "Auswertung";
                
                const errors = state.scoreWrong;
                const correct = state.scoreRight;
                const skipped = state.queue.length - (errors + correct);
                
                document.getElementById('end-stats').innerHTML = `
                    <p style="font-size: 1.5em; color: var(--success);">Richtig beantwortet: ${{correct}}</p>
                    <p style="font-size: 1.2em; color: var(--danger);">Falsch beantwortet: ${{errors}}</p>
                    <p style="font-size: 1.2em; color: var(--warning);">Immer noch Übersprungen: ${{skipped}}</p>
                `;
                
                document.getElementById('btn-retry-end').style.display = errors > 0 ? 'inline-block' : 'none';
                
            }} else if (viewName === 'quiz') {{
                if (state.currentIndex >= state.queue.length) {{
                    showView('end');
                    return;
                }}
                
                const qId = state.queue[state.currentIndex];
                document.getElementById('card_' + qId).classList.add('active');
                document.getElementById('progress-text').innerText = `Frage ${{state.currentIndex + 1}} von ${{state.queue.length}}`;
                
                const prevBtn = document.getElementById('prev_btn_' + qId);
                if(prevBtn) prevBtn.style.display = (state.currentIndex === 0) ? 'none' : 'inline-block';
                
                // Wenn die Frage schon beantwortet wurde, Status wiederherstellen (damit man zurückgehen kann und sehen kann)
                if (state.results[qId]) {{
                    const fbBox = document.getElementById('feedback_' + qId);
                    fbBox.style.display = 'block';
                    document.getElementById('check_btn_' + qId).style.display = 'none';
                    document.getElementById('skip_btn_' + qId).style.display = 'none';
                    document.getElementById('next_btn_' + qId).style.display = 'inline-block';
                    document.querySelectorAll(`input[name="q_${{qId}}"]`).forEach(r => r.disabled = true);
                    
                    const correctOption = document.querySelector(`input[name="q_${{qId}}"][value="true"]`);
                    if (correctOption) {{
                        correctOption.parentElement.classList.add('correct-highlight');
                    }}
                }} else {{
                     const fbBox = document.getElementById('feedback_' + qId);
                     fbBox.style.display = 'none';
                     document.getElementById('check_btn_' + qId).style.display = 'inline-block';
                     document.getElementById('skip_btn_' + qId).style.display = 'inline-block';
                     document.getElementById('next_btn_' + qId).style.display = 'none';
                     document.querySelectorAll(`input[name="q_${{qId}}"]`).forEach(r => {{ r.disabled = false; r.checked = false; }});
                     document.getElementById('hint_' + qId).style.display = 'none';
                }}
            }}
        }}

        function resetQuizData() {{
            if(confirm("Möchtest du wirklich deinen gesamten Fortschritt löschen? Dies kann nicht rückgängig gemacht werden.")) {{
                localStorage.removeItem('klr_quiz_state');
                location.reload();
            }}
        }}

        function startNewQuiz() {{
            const gebiet = document.getElementById('gebiet-select').value;
            const random = document.getElementById('randomize-checkbox').checked;
            
            let pool = allQuestions;
            if (gebiet !== 'Alle') {{
                pool = allQuestions.filter(q => q.gebiet === gebiet);
            }}
            
            let queueIds = pool.map(q => q.id);
            if (random) {{
                queueIds.sort(() => Math.random() - 0.5);
            }}
            
            state = {{
                queue: queueIds,
                currentIndex: 0,
                results: {{}},
                scoreRight: 0,
                scoreWrong: 0,
                activeView: 'quiz'
            }};
            
            // UI für alle Fragen zurücksetzen
            document.querySelectorAll('input[type="radio"]').forEach(r => {{ r.checked = false; r.disabled = false; }});
            document.querySelectorAll('.feedback-box').forEach(b => b.style.display = 'none');
            document.querySelectorAll('.hint-box').forEach(b => b.style.display = 'none');
            document.querySelectorAll('.option-label').forEach(l => l.classList.remove('correct-highlight', 'wrong-highlight'));
            
            saveState();
            updateHeaderStats();
            showView('quiz');
        }}

        function continueQuiz() {{
            showView('quiz');
        }}

        function retryWrongGlobal() {{
            // Ermittelt alle bisherigen falschen Antworten über alle Sessions hinweg
            const wrongIds = Object.keys(state.results).filter(k => state.results[k] === 'incorrect').map(k => parseInt(k));
            startRetrySession(wrongIds);
        }}

        function retryWrongSession() {{
            // Sammle falsche Antworten aus der aktuellen Queue
            const wrongIdsInQueue = state.queue.filter(qId => state.results[qId] === 'incorrect');
            startRetrySession(wrongIdsInQueue);
        }}

        function startRetrySession(wrongIds) {{
            if (wrongIds.length === 0) {{
                alert("Keine falsch beantworteten Fragen gefunden!");
                return;
            }}
            
            const random = document.getElementById('randomize-checkbox').checked;
            if (random) wrongIds.sort(() => Math.random() - 0.5);
            
            // Ergebnisse für die zu wiederholenden Fragen löschen
            wrongIds.forEach(id => delete state.results[id]);
            
            state.queue = wrongIds;
            state.currentIndex = 0;
            state.scoreRight = 0;
            state.scoreWrong = 0;
            
            // UI nur für diese Fragen zurücksetzen
            wrongIds.forEach(id => {{
                document.querySelectorAll(`input[name="q_${{id}}"]`).forEach(r => {{ 
                    r.checked = false; 
                    r.disabled = false; 
                    r.parentElement.classList.remove('correct-highlight', 'wrong-highlight');
                }});
                document.getElementById(`feedback_${{id}}`).style.display = 'none';
                document.getElementById(`hint_${{id}}`).style.display = 'none';
            }});
            
            saveState();
            updateHeaderStats();
            showView('quiz');
        }}

        function toggleHint(id) {{
            const box = document.getElementById('hint_' + id);
            box.style.display = (box.style.display === 'block') ? 'none' : 'block';
        }}

        function checkAnswer(id) {{
            if (state.results[id]) return; 
            const selected = document.querySelector(`input[name="q_${{id}}"]:checked`);
            if (!selected) {{ alert('Bitte wählen Sie eine Antwort aus.'); return; }}
            
            const feedbackBox = document.getElementById('feedback_' + id);
            
            document.getElementById('check_btn_' + id).style.display = 'none'; 
            document.getElementById('skip_btn_' + id).style.display = 'none'; 
            document.getElementById('next_btn_' + id).style.display = 'inline-block'; 
            document.querySelectorAll(`input[name="q_${{id}}"]`).forEach(radio => radio.disabled = true);
            
            feedbackBox.style.display = 'block';
            
            if (selected.value === "true") {{
                feedbackBox.className = "feedback-box correct";
                feedbackBox.innerHTML = "<strong>✅ Richtig!</strong><br><br>" + feedbackBox.getAttribute('data-feedback');
                state.results[id] = 'correct';
                state.scoreRight++;
                selected.parentElement.classList.add('correct-highlight');
            }} else {{
                feedbackBox.className = "feedback-box incorrect";
                feedbackBox.innerHTML = "<strong>❌ Leider falsch.</strong><br><br>" + feedbackBox.getAttribute('data-feedback');
                state.results[id] = 'incorrect';
                state.scoreWrong++;
                selected.parentElement.classList.add('wrong-highlight');
                
                const correctOption = document.querySelector(`input[name="q_${{id}}"][value="true"]`);
                if (correctOption) {{
                    correctOption.parentElement.classList.add('correct-highlight');
                }}
            }}
            
            updateHeaderStats();
            
            // Wir scrollen zur Feedback-Box
            // feedbackBox.scrollIntoView({{ behavior: 'smooth', block: 'nearest' }});
            
            saveState();
        }}

        function skipQuestion() {{ nextQuestion(); }}

        function prevQuestion() {{ 
            if (state.currentIndex > 0) {{ 
                state.currentIndex--; 
                saveState();
                showView('quiz'); 
                window.scrollTo(0,0); 
            }} 
        }}

        function nextQuestion() {{ 
            state.currentIndex++; 
            saveState();
            showView('quiz'); 
            window.scrollTo(0,0); 
        }}

        function toggleModal() {{
            const modal = document.getElementById('knowledge-modal');
            modal.style.display = (modal.style.display === 'flex') ? 'none' : 'flex';
        }}

        function closeModalOutside(event) {{
            if (event.target.id === 'knowledge-modal') {{ toggleModal(); }}
        }}

        window.onload = init;
    </script>
</body>
</html>
'''

    try:
        with open(html_datei, 'w', encoding='utf-8') as f:
            f.write(html_content)
        print(f"Erfolg! Quiz mit neuen interaktiven Funktionen generiert. Datei '{html_datei}' ist fertig.")
    except Exception as e:
        print(f"Fehler beim Speichern: {e}")

# Anwendung starten
excel_input = 'Frageliste 2372-clean.xlsx'
html_output = 'index.html'
generiere_html_aus_excel(excel_input, html_output)