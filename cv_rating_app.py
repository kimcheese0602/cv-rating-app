import streamlit as st
import re
from collections import Counter
from io import BytesIO

try:
    import fitz
    HAS_PYMUPDF = True
except ImportError:
    HAS_PYMUPDF = False

try:
    from docx import Document
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False

ROLE_KEYWORDS = {
    "Procurement (Junior/Associate)": {
        "critical": ["procurement","sourcing","supplier","vendor","contract","negotiation","tender","rfp","rfq","purchase order","category management"],
        "important": ["sap ariba","sap","erp","cost optimization","spend analysis","supplier evaluation","kpi","stakeholder","compliance","supply chain","strategic sourcing","vendor management","purchase","buying","operational procurement","indirect procurement","direct procurement"],
        "nice_to_have": ["power bi","excel","data analysis","agile","project management","process optimization","dora","tco","benchmarking","market intelligence","framework agreement","msa","sla","capex","opex"]
    },
    "Procurement (Mid-Senior)": {
        "critical": ["procurement","strategic sourcing","category management","supplier management","contract negotiation","stakeholder management","cost savings","spend"],
        "important": ["sap ariba","tender management","supplier development","risk management","compliance","dora","governance","budget","team lead","cross-functional","global sourcing","it procurement","capex"],
        "nice_to_have": ["digital transformation","ai","automation","sustainability","esg","change management","mentoring","kpi dashboard","power bi"]
    },
    "Graduate Program": {
        "critical": ["international","business","analytical","teamwork","communication","problem-solving","adaptability","motivation"],
        "important": ["project management","data analysis","cross-functional","stakeholder","process improvement","erp","sap","agile","leadership"],
        "nice_to_have": ["procurement","supply chain","consulting","strategy","digital","innovation","sustainability","excel","power bi","sql"]
    },
    "Business Analyst (Junior)": {
        "critical": ["business analysis","requirements","stakeholder","documentation","process","workflow","uat","testing"],
        "important": ["erp","agile","scrum","sprint","jira","confluence","data analysis","reporting","kpi","dashboard","sql","process optimization"],
        "nice_to_have": ["odoo","sap","power bi","project management","cross-functional","change management","user stories","bpmn","lean"]
    },
    "ERP Consultant/Specialist (Junior)": {
        "critical": ["erp","implementation","configuration","module","go-live","requirements","system","integration"],
        "important": ["sap","odoo","agile","testing","uat","data migration","training","documentation","process mapping","workflow","customization"],
        "nice_to_have": ["procurement module","manufacturing","accounting","hr","sales","project management","sql","api","reporting","power bi"]
    },
    "Project Manager (Junior)": {
        "critical": ["project management","planning","timeline","scope","deliverables","stakeholder","team","coordination"],
        "important": ["agile","scrum","sprint","risk management","budget","milestone","cross-functional","reporting","communication","jira"],
        "nice_to_have": ["erp","procurement","change management","pmp","prince2","lean","six sigma","gantt","resource planning","kpi"]
    }
}
def extract_text(uploaded_file):
    file_bytes = uploaded_file.read()
    name = uploaded_file.name.lower()
    if name.endswith(".pdf"):
        if not HAS_PYMUPDF: return "[ERROR] PyMuPDF not installed. Run: pip install PyMuPDF"
        doc = fitz.open(stream=file_bytes, filetype="pdf")
        text = "\n".join([page.get_text() for page in doc])
        doc.close()
        return text
    elif name.endswith(".docx"):
        if not HAS_DOCX: return "[ERROR] python-docx not installed. Run: pip install python-docx"
        doc = Document(BytesIO(file_bytes))
        return "\n".join([p.text for p in doc.paragraphs])
    elif name.endswith(".txt"):
        return file_bytes.decode("utf-8")
    return "[ERROR] Unsupported format. Upload PDF, DOCX, or TXT."


def score_ats(cv_text):
    score = 100
    findings, tips = [], []
    has_email = bool(re.search(r"[\w.+-]+@[\w-]+\.[\w.-]+", cv_text))
    has_phone = bool(re.search(r"[+]?[\d\s\-()]{8,}", cv_text))
    has_linkedin = bool(re.search(r"linkedin", cv_text, re.I))
    if not has_email:
        score -= 15; findings.append("❌ No email detected"); tips.append("Add email in contact section")
    else:
        findings.append("✅ Email detected")
    if not has_phone:
        score -= 10; findings.append("❌ No phone detected"); tips.append("Add phone with +49 country code")
    else:
        findings.append("✅ Phone number detected")
    if not has_linkedin:
        score -= 5; findings.append("⚠️ No LinkedIn detected"); tips.append("Add LinkedIn URL")
    else:
        findings.append("✅ LinkedIn detected")
    pronouns = len(re.findall(r"\b(I |my |me |we |our )\b", cv_text, re.I))
    if pronouns > 0:
        score -= min(pronouns * 3, 15)
        findings.append(f"⚠️ {pronouns} personal pronouns found")
        tips.append("Remove pronouns (I, my, we) - use action verbs")
    bullets = re.findall(r"[-•*]\s*([A-Z][a-z]+)", cv_text)
    verbs = {"Led","Managed","Developed","Executed","Delivered","Supported","Conducted","Coordinated","Implemented","Created","Analyzed","Designed","Built","Established","Drove","Achieved","Improved","Reduced","Increased","Negotiated","Collaborated","Defined","Maintained","Tracked","Monitored","Prepared","Contributed","Bridged","Spearheaded","Secured","Streamlined"}
    if bullets:
        ratio = sum(1 for b in bullets if b in verbs) / len(bullets)
        if ratio < 0.5:
            score -= 10
            findings.append(f"⚠️ Only {ratio*100:.0f}% bullets start with action verbs")
            tips.append("Start bullets with: Led, Managed, Executed, Delivered, Supported...")
        else:
            findings.append(f"✅ {ratio*100:.0f}% bullets use strong action verbs")
    dates = re.findall(r"(\d{4}|(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\w* \d{4})", cv_text)
    if len(dates) < 3:
        score -= 10; findings.append("⚠️ Few dates found"); tips.append("Add clear dates (MM/YYYY) for all roles")
    else:
        findings.append(f"✅ {len(dates)} dates detected")
    headers = ["experience","education","skills","profile","summary","languages"]
    h_count = sum(1 for h in headers if re.search(r"\b" + h + r"\b", cv_text, re.I))
    if h_count < 3:
        score -= 10; findings.append(f"⚠️ Only {h_count} standard headers")
        tips.append("Use: PROFILE, PROFESSIONAL EXPERIENCE, EDUCATION, SKILLS & LANGUAGES")
    else:
        findings.append(f"✅ {h_count} standard section headers")
    nums = re.findall(r"\d+[%KkMm]|\d+[+]|€\d+|\d+€", cv_text)
    if len(nums) < 2:
        score -= 5; findings.append("⚠️ Limited quantification"); tips.append("Add metrics: €250K+, 10+ projects, 3+ years")
    else:
        findings.append(f"✅ {len(nums)} quantified metrics found")
    wc = len(cv_text.split())
    if wc > 800:
        score -= 5; findings.append(f"⚠️ Long ({wc} words)"); tips.append("Target 400-700 words for 1-page CV")
    elif wc < 250:
        score -= 10; findings.append(f"⚠️ Short ({wc} words)"); tips.append("Add more detail (400-700 words)")
    else:
        findings.append(f"✅ Good length ({wc} words)")
    return max(score, 0), findings, tips
def score_content(cv_text, role):
    score = 0
    findings, tips = [], []
    cv_lower = cv_text.lower()
    if re.search(r"(profile|summary)", cv_text, re.I):
        score += 5; findings.append("✅ Profile/Summary section present")
        if re.search(r"\d+[+]?\s*year", cv_text, re.I):
            score += 10; findings.append("✅ Years of experience in profile")
        else:
            score += 3; tips.append("Add years: '4+ years of experience in...'")
    else:
        findings.append("❌ No Profile section"); tips.append("Add 2-3 line professional summary")
    bullets = re.findall(r"[-•*]\s*.+", cv_text)
    if len(bullets) >= 12:
        score += 15; findings.append(f"✅ Strong detail ({len(bullets)} bullets)")
    elif len(bullets) >= 6:
        score += 10; findings.append(f"✅ Good detail ({len(bullets)} bullets)")
    else:
        score += 5; findings.append(f"⚠️ Few bullets ({len(bullets)})"); tips.append("Add 4-6 bullets per recent role")
    achiev = ["achieved","delivered","improved","reduced","increased","saved","secured","optimized","streamlined","transformed","established","launched","drove","exceeded"]
    a_count = sum(1 for w in achiev if w in cv_lower)
    if a_count >= 3:
        score += 15; findings.append(f"✅ Achievement language strong ({a_count} instances)")
    elif a_count >= 1:
        score += 8; findings.append(f"⚠️ Some achievement language ({a_count} instances)")
        tips.append("Use more: delivered, secured, optimized, drove, achieved")
    else:
        findings.append("❌ No achievement language"); tips.append("Reframe bullets: action verb + result")
    quants = re.findall(r"\d+[%KkMm]|\d+[+]|€\d+|\d+\s*(?:project|client|team|member|supplier|contract|countr|module)", cv_text, re.I)
    if len(quants) >= 4:
        score += 10; findings.append(f"✅ Well quantified ({len(quants)} metrics)")
    elif len(quants) >= 2:
        score += 5; findings.append(f"⚠️ Some quantification ({len(quants)})"); tips.append("Add: €250K+, 10+ projects, teams of 15, 100+ users")
    else:
        findings.append("❌ Limited quantification"); tips.append("Add numbers to 50%+ of bullets")
    if re.search(r"education", cv_text, re.I) and re.search(r"(M\.?Sc|B\.?Sc|Master|Bachelor)", cv_text, re.I):
        score += 10; findings.append("✅ Education with degree detected")
    else:
        score += 3; tips.append("Add Education with degree, institution, graduation date")
    if re.search(r"skill", cv_text, re.I):
        cats = re.findall(r"(ERP|Tool|Data|Analy|Language|Competen|Business)", cv_text, re.I)
        if len(set(cats)) >= 3:
            score += 10; findings.append("✅ Skills well-categorized")
        else:
            score += 5; tips.append("Categorize: ERP & Tools, Data & Analysis, Core Competencies, Languages")
    else:
        tips.append("Add a Skills section")
    kws = ROLE_KEYWORDS.get(role, ROLE_KEYWORDS["Procurement (Junior/Associate)"])
    found = sum(1 for k in kws["critical"] if k in cv_lower)
    total = len(kws["critical"])
    ratio = found / total if total else 0
    if ratio >= 0.6:
        score += 15; findings.append(f"✅ Critical keywords: {found}/{total}")
    elif ratio >= 0.3:
        score += 8
        missing = [k for k in kws["critical"] if k not in cv_lower]
        findings.append(f"⚠️ Partial keywords: {found}/{total}"); tips.append(f"Add: {', '.join(missing[:5])}")
    else:
        missing = [k for k in kws["critical"] if k not in cv_lower]
        findings.append(f"❌ Low keywords: {found}/{total}"); tips.append(f"Must add: {', '.join(missing[:6])}")
    return min(score, 100), findings, tips


def score_keywords(cv_text, jd_text, role):
    if not jd_text or not jd_text.strip():
        return None, ["ℹ️ No job description provided"], ["Upload a JD for keyword matching"]
    cv_lower, jd_lower = cv_text.lower(), jd_text.lower()
    findings, tips = [], []
    kws = ROLE_KEYWORDS.get(role, ROLE_KEYWORDS["Procurement (Junior/Associate)"])
    jd_crit = [k for k in kws["critical"] if k in jd_lower]
    jd_imp = [k for k in kws["important"] if k in jd_lower]
    jd_nice = [k for k in kws["nice_to_have"] if k in jd_lower]
    cv_crit = [k for k in jd_crit if k in cv_lower]
    cv_imp = [k for k in jd_imp if k in cv_lower]
    cv_nice = [k for k in jd_nice if k in cv_lower]
    miss_crit = [k for k in jd_crit if k not in cv_lower]
    miss_imp = [k for k in jd_imp if k not in cv_lower]
    miss_nice = [k for k in jd_nice if k not in cv_lower]
    s1 = (len(cv_crit)/len(jd_crit))*40 if jd_crit else 20
    s2 = (len(cv_imp)/len(jd_imp))*35 if jd_imp else 17
    s3 = (len(cv_nice)/len(jd_nice))*25 if jd_nice else 12
    score = s1 + s2 + s3
    total = len(jd_crit)+len(jd_imp)+len(jd_nice)
    matched = len(cv_crit)+len(cv_imp)+len(cv_nice)
    if total:
        findings.append(f"📊 Match: {matched}/{total} ({matched/total*100:.0f}%)")
    if cv_crit: findings.append(f"✅ Critical matched: {', '.join(cv_crit[:5])}")
    if miss_crit:
        findings.append(f"🔴 Missing critical: {', '.join(miss_crit[:5])}")
        tips.append(f"ADD critical keywords: {', '.join(miss_crit[:5])}")
    if cv_imp: findings.append(f"✅ Important matched: {', '.join(cv_imp[:5])}")
    if miss_imp:
        findings.append(f"🟡 Missing important: {', '.join(miss_imp[:5])}")
        tips.append(f"Consider adding: {', '.join(miss_imp[:5])}")
    if miss_nice and len(miss_nice) > 2:
        tips.append(f"Nice-to-have to consider: {', '.join(miss_nice[:4])}")
    stop = {"the","and","for","with","that","this","from","are","was","will","have","has","been","their","they","you","your","our","can","all","would","could","should","not","but","also","other","which","what","when","where","who","how","than","into","over","such","only","very","well","about","after","before","between","through","during","within","including","across","able","ensure","work","working","role","position","company","team","part","join","offer","based","may","must"}
    jd_words = re.findall(r"\b[a-z]{3,}\b", jd_lower)
    freq = Counter(jd_words)
    top_jd = [w for w,c in freq.most_common(25) if c>=2 and w not in stop][:15]
    in_cv = [w for w in top_jd if w in cv_lower]
    if top_jd:
        findings.append(f"📋 Top JD terms in CV: {len(in_cv)}/{len(top_jd)}")
        missing_top = [w for w in top_jd if w not in cv_lower][:5]
        if missing_top: tips.append(f"Frequently used in JD but missing: {', '.join(missing_top)}")
    return min(round(score), 100), findings, tips
def score_market(cv_text):
    score = 0
    findings, tips = [], []
    cv_lower = cv_text.lower()
    checks = [
        (r"(munich|berlin|frankfurt|hamburg|germany|deutschland)", 12, "German location", "Add city in header (e.g. Munich, Germany)"),
        (r"(english|german).{0,20}(fluent|native|a[12]|b[12]|c[12]|beginner|advanced)", 13, "Language levels (CEFR)", "Add CEFR levels: English - Fluent, German - A2"),
        (r"(universit|hochschule|htw|tu |fu |lmu)", 13, "Education institution", "Add full university name"),
        (r"(20[12]\d|present|current)", 12, "Employment dates", "Add clear dates (MM/YYYY)"),
        (r"linkedin", 7, "LinkedIn profile", "Add LinkedIn URL"),
        (r"english.{0,15}(fluent|proficient|native|c[12])", 7, "English proficiency stated", "State English level clearly (Fluent/C1)"),
        (r"(international|global|multinational|eu|us)", 7, "International experience", "Highlight international experience"),
    ]
    for pattern, points, label, tip in checks:
        if re.search(pattern, cv_lower):
            score += points; findings.append(f"✅ {label} detected")
        else:
            findings.append(f"⚠️ {label} not detected"); tips.append(tip)
    if re.search(r"german.{0,15}(a[12]|beginner|basic)", cv_lower):
        score += 7; findings.append("✅ German level stated honestly (A2/Beginner)")
        if not re.search(r"(improv|learn|target|currently)", cv_lower):
            tips.append("Add 'actively improving' next to German level")
    elif re.search(r"german.{0,15}(b[12]|c[12]|fluent)", cv_lower):
        score += 7; findings.append("✅ German proficiency stated")
    else:
        findings.append("⚠️ German level not clearly stated")
        tips.append("State German level: 'German - A2 (actively improving)'")
    if re.search(r"(berlin|munich).{0,40}(universit|htw|master)", cv_lower):
        score += 5; findings.append("✅ German/EU education (strong market signal)")
    if re.search(r"nationalit", cv_lower):
        score += 3; findings.append("✅ Nationality mentioned")
    else:
        tips.append("Consider adding nationality (common in German applications)")
    if re.search(r"(available|verfügbar|start date)", cv_lower):
        score += 4; findings.append("✅ Availability mentioned")
    else:
        tips.append("Add availability/start date (expected in Germany)")
    wc = len(cv_text.split())
    if wc <= 700:
        score += 5; findings.append(f"✅ Concise length ({wc} words)")
    elif wc <= 900:
        score += 3; findings.append(f"⚠️ Slightly long ({wc} words)")
    else:
        findings.append(f"⚠️ Long ({wc} words)"); tips.append("Consider condensing to 1 page")
    return min(score, 100), findings, tips


def section_tips(cv_text, jd_text, role):
    sections = {}
    cv_lower = cv_text.lower()
    jd_lower = jd_text.lower() if jd_text else ""
    p_tips = []
    if "strategic global" in cv_lower and "Junior" in role:
        p_tips.append("🔄 Remove 'Strategic Global' - too senior for junior/associate roles")
    if not re.search(r"\d+[+]?\s*year", cv_lower):
        p_tips.append("➕ Add years of experience: '4+ years of experience in...'")
    if jd_lower:
        kws = ROLE_KEYWORDS.get(role, ROLE_KEYWORDS["Procurement (Junior/Associate)"])
        prof = re.search(r"(?:profile|summary)(.{50,400}?)(?=\n\n|experience|professional)", cv_text, re.I|re.DOTALL)
        if prof:
            pt = prof.group(1).lower()
            miss = [k for k in kws["critical"] if k in jd_lower and k not in pt]
            if miss: p_tips.append(f"➕ Add to profile summary: {', '.join(miss[:4])}")
    sections["📝 Profile/Summary"] = p_tips if p_tips else ["✅ Profile looks good"]
    e_tips = []
    bullets = re.findall(r"[-•*]\s*(.+)", cv_text)
    if bullets:
        weak = [b for b in bullets if b.strip()[:20].lower().startswith(("responsible","in charge","helped","worked on","involved"))]
        if weak: e_tips.append(f"🔄 {len(weak)} weak openers - replace with action verbs")
        q_bullets = [b for b in bullets if re.search(r"\d+|€|%", b)]
        if len(q_bullets) < len(bullets) * 0.3:
            e_tips.append(f"📊 Only {len(q_bullets)}/{len(bullets)} bullets quantified - aim for 50%+")
    if jd_lower:
        if "dora" in jd_lower and "dora" not in cv_lower: e_tips.append("➕ Add DORA compliance experience")
        if "sap ariba" in jd_lower and "ariba" not in cv_lower: e_tips.append("➕ Mention SAP Ariba prominently")
        if "agile" in jd_lower and "agile" not in cv_lower: e_tips.append("➕ Mention Agile methodology")
    sections["💼 Professional Experience"] = e_tips if e_tips else ["✅ Strong experience section"]
    ed_tips = []
    if not re.search(r"(coursework|thesis|specializ)", cv_lower):
        ed_tips.append("💡 For junior roles: add relevant coursework or thesis topic")
    if not re.search(r"(gpa|grade|scholarship)", cv_lower):
        ed_tips.append("💡 Add GPA or scholarship if strong")
    sections["🎓 Education"] = ed_tips if ed_tips else ["✅ Education section looks good"]
    s_tips = []
    if jd_lower:
        tools = ["sap","ariba","odoo","power bi","excel","sql","jira","confluence","tableau","python"]
        missing_tools = [t for t in tools if t in jd_lower and t not in cv_lower]
        if missing_tools: s_tips.append(f"➕ JD mentions tools not in Skills: {', '.join(missing_tools)}")
    sections["🛠️ Skills & Languages"] = s_tips if s_tips else ["✅ Skills section looks good"]
    return sections
def main():
    st.set_page_config(page_title="CV Rating App", page_icon="📊", layout="wide")
    st.markdown("""
    <style>
    .score-box {padding:20px;border-radius:12px;text-align:center;margin:8px 0}
    .high {background:#d4edda;border:2px solid #28a745}
    .med {background:#fff3cd;border:2px solid #ffc107}
    .low {background:#f8d7da;border:2px solid #dc3545}
    </style>
    """, unsafe_allow_html=True)
    st.title("📊 CV Rating App")
    st.caption("🎯 Procurement | Graduate Program | Business Analyst | ERP | Project Manager")
    st.caption("🇩🇪 German Market | International Companies | English-speaking roles")
    st.divider()
    with st.sidebar:
        st.header("⚙️ Settings")
        role = st.selectbox("🎯 Target Role", list(ROLE_KEYWORDS.keys()))
        st.divider()
        st.markdown("**Scoring Weights:**\n- ATS Friendly: 25%\n- Content: 35%\n- Keywords: 30%\n- Market Fit: 10%")
        st.divider()
        st.info("🔒 100% local. No data sent externally.")
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("📄 Upload Your CV")
        cv_file = st.file_uploader("CV (PDF, DOCX, or TXT)", type=["pdf","docx","txt"])
        if cv_file: st.success(f"✅ {cv_file.name}")
    with col2:
        st.subheader("📋 Job Description (optional)")
        jd_method = st.radio("Input:", ["Paste text","Upload file"], horizontal=True)
        jd_text = ""
        if jd_method == "Paste text":
            jd_text = st.text_area("Paste JD here", height=150, placeholder="Paste job description...")
        else:
            jd_file = st.file_uploader("Upload JD", type=["pdf","docx","txt"], key="jd")
            if jd_file:
                jd_text = extract_text(jd_file)
                st.success(f"✅ {jd_file.name}")
    st.divider()
    if cv_file and st.button("🔍 Analyze CV", type="primary", use_container_width=True):
        cv_text = extract_text(cv_file)
        if cv_text.startswith("[ERROR]"):
            st.error(cv_text); return
        with st.expander("📝 Extracted Text (verify parsing)"):
            st.text(cv_text[:2000])
        st.divider()
        s_ats, f_ats, t_ats = score_ats(cv_text)
        s_con, f_con, t_con = score_content(cv_text, role)
        s_kw, f_kw, t_kw = score_keywords(cv_text, jd_text, role)
        s_mkt, f_mkt, t_mkt = score_market(cv_text)
        if s_kw is not None:
            overall = round(s_ats*0.25 + s_con*0.35 + s_kw*0.30 + s_mkt*0.10)
        else:
            overall = round(s_ats*0.35 + s_con*0.50 + s_mkt*0.15)
        css = "high" if overall >= 75 else "med" if overall >= 50 else "low"
        emoji = "🟢" if overall >= 75 else "🟡" if overall >= 50 else "🔴"
        st.markdown(f'<div class="score-box {css}"><h1>{emoji} {overall}/100</h1><p>Role: <b>{role}</b></p></div>', unsafe_allow_html=True)
        st.divider()
        cols = st.columns(4)
        for col, lbl, sc in zip(cols, ["🤖 ATS","📝 Content","🔑 Keywords","🇩🇪 Market"], [s_ats,s_con,s_kw,s_mkt]):
            with col:
                if sc is None: st.metric(lbl, "N/A")
                else: st.metric(lbl, f"{sc}/100", delta="Strong" if sc>=75 else "Improve" if sc>=50 else "Critical", delta_color="normal" if sc>=75 else "off" if sc>=50 else "inverse")
        st.divider()
        st.header("🔍 Detailed Analysis")
        with st.expander(f"🤖 ATS Friendliness — {s_ats}/100", expanded=True):
            c1,c2 = st.columns(2)
            with c1:
                st.markdown("**Findings:**")
                for f in f_ats: st.markdown(f"- {f}")
            with c2:
                st.markdown("**Suggestions:**")
                for t in t_ats: st.markdown(f"- {t}")
        with st.expander(f"📝 Content Quality — {s_con}/100", expanded=True):
            c1,c2 = st.columns(2)
            with c1:
                st.markdown("**Findings:**")
                for f in f_con: st.markdown(f"- {f}")
            with c2:
                st.markdown("**Suggestions:**")
                for t in t_con: st.markdown(f"- {t}")
        with st.expander(f"🔑 Keyword Match — {s_kw if s_kw else 'N/A'}/100", expanded=True):
            if s_kw is not None:
                c1,c2 = st.columns(2)
                with c1:
                    st.markdown("**Findings:**")
                    for f in f_kw: st.markdown(f"- {f}")
                with c2:
                    st.markdown("**Suggestions:**")
                    for t in t_kw: st.markdown(f"- {t}")
            else:
                st.info("📋 Upload a job description to unlock keyword analysis")
        with st.expander(f"🇩🇪 German Market Fit — {s_mkt}/100", expanded=True):
            c1,c2 = st.columns(2)
            with c1:
                st.markdown("**Findings:**")
                for f in f_mkt: st.markdown(f"- {f}")
            with c2:
                st.markdown("**Suggestions:**")
                for t in t_mkt: st.markdown(f"- {t}")
        st.divider()
        st.header("📋 Section Improvements")
        for name, tips_list in section_tips(cv_text, jd_text, role).items():
            with st.expander(name):
                for t in tips_list: st.markdown(f"- {t}")
        st.divider()
        st.header("🎯 Priority Actions")
        all_tips = t_ats + t_con + t_kw + t_mkt
        critical_tips = [t for t in all_tips if any(w in t.lower() for w in ["add","must","critical","missing"])]
        other_tips = [t for t in all_tips if t not in critical_tips]
        if critical_tips:
            st.markdown("**🔴 High Priority:**")
            for i,t in enumerate(critical_tips[:5],1): st.markdown(f"{i}. {t}")
        if other_tips:
            st.markdown("**🟡 Medium Priority:**")
            for i,t in enumerate(other_tips[:5],1): st.markdown(f"{i}. {t}")
    elif not cv_file:
        st.info("👆 Upload your CV to get started!")


if __name__ == "__main__":
    main()
