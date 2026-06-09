from __future__ import annotations

from collections import Counter
from datetime import datetime
from io import BytesIO
import re
from typing import Iterable

from docx import Document
import pandas as pd
import streamlit as st


STANDARD_COLUMNS = [
    "naam",
    "huidig_bedrijf",
    "functietitel",
    "locatie",
    "sector",
    "tenure_huidige_rol_jaren",
    "info_eerdere_rollen",
    "hoogst_afgeronde_opleiding",
    "interesse_signalen",
    "activiteit",
]

DEFAULT_SECTOR_RULES = """GGZ: ggz, geestelijke gezondheidszorg, psychiatrie, psychiatrisch, verslavingszorg, autisme, mental care, parnassia, arkin, yulius, pro persona, altrecht, lentis, mondriaan, fivoor, dimence, ggze, ggz breburg, ggz centraal, tactus, ggz drenthe, ggz friesland, reinier van arkel, emergis, karakter, vnn, novadic, eleos, lievegoed, de viersprong, yes we can, hsk groep, mentaal beter
MSZ: msz, ziekenhuis, ziekenhuizen, medisch centrum, umc, universitair medisch, kliniek, klinieken, diagnostiek, laboratorium, zbc, erasmus mc, amsterdam umc, umcg, radboudumc, lumc, olvg, zuyderland, maastricht umc, canisius, st. antonius, elisabeth-tweesteden, etz, bravis, maxima mc, máxima mc, hagaziekenhuis, gelre, alrijne, adrz, tergooi, martini, bergman, diakonessenhuis, franciscus, antoni van leeuwenhoek, jeroen bosch, viecuri, catharina, rijnstate, dijklander, slingeland, zgt, flevoziekenhuis, maasstad, laurentius ziekenhuis, curaMare, curamare, dicoon, acibadem, leiden university medical, medisch spectrum twente, amphia, saxenburgh, sjg weert, caresq, rivas zorggroep, unilabs
VVT: vvt, ouderenzorg, thuiszorg, verpleeghuis, verpleeghuizen, woonzorg, wijkverpleging, zorgcentrum, zorgcentra, brabantzorg, careyn, laurens, cordaan, aafje, tantelouise, de zorgcirkel, beweging 3.0, zonnehuisgroep, zorgsaam, actief zorg, bloezem, lelie zorggroep, surplus, mijzo, zorggroep solis, carinova, zinn, magentazorg, florence, coloriet, sevagram, van neynsel, domus valuas, groenhuysen, pantein, pieter van foreest, welthuis, amstelring, dignis, axioncontinu, sensire, meandergroep, amarijn, saffier, kwadrantgroep, omring, cicero, humanitas, opella, viva! zorggroep, amsta, hilverzorg, avoord, topaz, zorggroep ter weel, korian, marente, activite, zuidoostzorg, zorg groep beek, miep huishoudelijke, zorgspectrum, vilente, viattence, bartholomeus gasthuis, zorgspectrum het zand, quarijn, stichting land van horne
WLZ: wlz, gehandicaptenzorg, beperking, beperkingen, lvb, langdurige zorg, forensische zorg, jeugdwet, jeugdzorg, pluryn, middin, cosis, 's heeren loo, cello, siza, amarant, gemiva, ipse de bruggen, koninklijke visio, ambiq, timon, ons tweede thuis, triade vitree, zuidwester, stichting radar, wender, sdw zorg, oro, lunet, zozijn, omega groep, tragel, prinsenstichting, sovak, sensa zorg, zideris, aveleijn, driestroom, baalderborg, abrona, philadelphia, ribw, raphaëlstichting, raphaelstichting
Pharma: pharma, pharmaceutical, geneesmiddel, geneesmiddelen, farma, fabrikant, manufacturing, batch release, gmp, qa, quality assurance, qualified person, qp
Biotech: biotech, biotechnology, biologics, biologisch, cell therapy, gene therapy
CDMO: cdmo, cmo, contract manufacturing, contract development, fill finish
Medical Devices: medical device, medical devices, medtech, hulpmiddel, hulpmiddelen, device
Healthcare: healthcare, health care, zorg, ziekenhuis, clinic, kliniek
Overig: overig"""

SECTOR_PRESETS = {
    "Automatisch": {
        "buckets": [],
        "include_rule_buckets": [],
    },
    "Zorg": {
        "buckets": ["GGZ", "MSZ", "VVT", "WLZ", "Overig"],
        "include_rule_buckets": ["GGZ", "MSZ", "VVT", "WLZ", "Overig"],
    },
    "Pharma / biotech": {
        "buckets": ["Pharma", "Biotech", "CDMO", "Medical Devices", "Healthcare", "Overig"],
        "include_rule_buckets": ["Pharma", "Biotech", "CDMO", "Medical Devices", "Healthcare", "Overig"],
    },
}

DEFAULT_INCLUDE_TERMS = {
    "Zorg": [
        "zorgverkoop",
        "zorgverkoper",
        "zorgcontractering",
        "zorgcontract",
        "contractering",
        "zorginkoop",
        "zorginkoper",
        "zorgbemiddeling",
        "zorgbemiddelaar",
        "accountmanager zorg",
        "adviseur zorgverkoop",
        "strategisch zorgverkoper",
        "relatiebeheer",
        "wlz",
        "zvw",
        "wmo",
        "jeugdwet",
    ],
    "Pharma / biotech": [
        "qualified person",
        "qp",
        "q.p.",
        "gmp",
        "batch release",
        "qa",
        "quality assurance",
        "quality",
        "regulatory affairs",
        "regulatory",
        "pharma",
        "pharmaceutical",
        "biotech",
        "cdmo",
        "medical device",
        "medtech",
    ],
}

DEFAULT_EXCLUDE_TERMS = [
    "recruiter",
    "student",
    "stagiair",
    "intern",
    "finance consultant",
    "treasury",
    "opticien",
    "ergotherapeut",
    "fysiotherapeut",
    "verpleegkundige",
    "docent",
    "hogeschooldocent",
    "raad van toezicht",
]


BADGES = (
    "Dit profiel heeft een LinkedIn Premium-abonnement.",
    "Derdegraads connectie",
    "Tweedegraads connectie",
    "Eerstegraads connectie",
    "Gesourced",
    "This profile has a LinkedIn Premium subscription",
    "Third degree connection",
    "Second degree connection",
    "First degree connection",
    "Sourced",
)

STOP_PREFIXES = (
    "Stadium wijzigen",
    "Bericht sturen naar ",
    "Meer acties voor ",
    "Inuncontacted",
    "Opgeslagen op ",
    "door ",
    "Change stage",
    "Message ",
    "More actions for ",
    "Saved on ",
    "by ",
)


def clean(text: object) -> str:
    text = "" if text is None else str(text)
    text = text.replace("\xa0", " ").replace("Â·", "·").replace("â€“", "–")
    return " ".join(text.split()).strip()


def tokens(text: str) -> list[str]:
    return [clean(x).lower() for x in re.split(r"[,;\n]+", text or "") if clean(x)]


def parse_sector_rules(text: str, allowed_buckets: list[str]) -> dict[str, list[str]]:
    allowed = {bucket.lower(): bucket for bucket in allowed_buckets}
    rules: dict[str, list[str]] = {bucket: [] for bucket in allowed_buckets}
    for raw_line in (text or "").splitlines():
        line = clean(raw_line)
        if not line or ":" not in line:
            continue
        bucket_raw, keywords_raw = line.split(":", 1)
        bucket = allowed.get(clean(bucket_raw).lower())
        if not bucket:
            continue
        rules[bucket].extend(tokens(keywords_raw))
    return rules


def sector_preset_for_df(df: pd.DataFrame) -> str:
    blob = " ".join(
        " ".join(clean(row.get(c)) for c in ["huidig_bedrijf", "functietitel", "sector", "info_eerdere_rollen"])
        for _, row in df.head(200).iterrows()
    ).lower()
    zorg_score = sum(
        blob.count(term)
        for term in ["zorgverkoop", "zorgcontract", "ggz", "ziekenhuis", "ouderenzorg", "wlz", "vvt", "parnassia", "pluryn"]
    )
    pharma_score = sum(
        blob.count(term)
        for term in ["qualified person", "batch release", "gmp", "pharma", "pharmaceutical", "biotech", "cdmo", "medical device"]
    )
    return "Pharma / biotech" if pharma_score > zorg_score else "Zorg"


def sector_config(preset_name: str, df: pd.DataFrame) -> tuple[list[str], dict[str, list[str]], str]:
    resolved = sector_preset_for_df(df) if preset_name == "Automatisch" else preset_name
    preset = SECTOR_PRESETS[resolved]
    buckets = preset["buckets"]
    rules = parse_sector_rules(DEFAULT_SECTOR_RULES, preset["include_rule_buckets"])
    return buckets, rules, resolved


def is_badge(text: str) -> bool:
    return any(text.startswith(prefix) for prefix in BADGES)


def is_sector_line(text: str) -> bool:
    return text.startswith(("Â· ", "· "))


def clean_sector_line(text: str) -> str:
    return clean(re.sub(r"^(?:Â·|·)\s*", "", text))


def is_experience_marker(text: str) -> bool:
    return text in {"ErvaringErvaring in profiel", "ExperienceProfile experience"}


def is_education_marker(text: str) -> bool:
    return text in {"OpleidingOpleidingen op het profiel", "EducationProfile education"}


def is_interest_marker(text: str) -> bool:
    return text.startswith(
        (
            "LabelsLabels",
            "Decoraties voor rij met profielinteresses",
            "TagsCandidate tags",
            "Profile interest row decorations",
        )
    )


def is_activity_marker(text: str) -> bool:
    return text.startswith(("Decoraties voor rij met profielactiviteit", "Profile activity row decorations"))



def strip_date_suffix(line: str) -> str:
    return clean(re.sub(r"\s*(?:Â·|·|愧)\s*.*$", "", line))


def split_title_company(headline: str) -> tuple[str, str]:
    h = clean(headline)
    for sep in (" bij ", " at ", " @ "):
        if sep in h:
            title, company = h.rsplit(sep, 1)
            return clean(title), clean(company)
    return h, ""


def years_from_line(line: str) -> int | None:
    if not line:
        return None
    years = [int(y) for y in re.findall(r"\b(19\d{2}|20\d{2})\b", line)]
    if not years:
        return None
    start = years[0]
    end = datetime.now().year if re.search(r"\bheden\b|\bpresent\b", line, re.I) else years[-1]
    return max(0, end - start)


def infer_education_level(line: str) -> str:
    t = line.lower()
    patterns = [
        ("PhD/Doctorate", r"\b(phd|doctor|doctorate|promotie)\b"),
        ("MBA", r"\bmba\b"),
        ("Master", r"\b(master|msc|m\.sc|ma|m\.a\.|m[a-z]*\.)\b"),
        ("Bachelor/HBO", r"\b(bachelor|bsc|b\.sc|hbo|hogeschool|bba|bachelorgraad|post bachelor|post-hbo)\b"),
        ("WO", r"\b(university|universiteit|wo)\b"),
        ("MBO", r"\bmbo\b"),
    ]
    for label, pattern in patterns:
        if re.search(pattern, t):
            return label
    return ""


def education_summary(lines: list[str]) -> str:
    if not lines:
        return ""
    rank = {"PhD/Doctorate": 6, "MBA": 5, "Master": 4, "WO": 3, "Bachelor/HBO": 2, "MBO": 1}
    best = max(lines, key=lambda x: rank.get(infer_education_level(x), 0))
    level = infer_education_level(best)
    detail = best.split(",", 1)[1] if "," in best else best
    detail = re.sub(r"\s*·\s*[\d–\- heden]+.*$", "", detail).strip(" -")
    return f"{level} - {clean(detail)}" if level and clean(detail) else level or clean(detail)


def compact_role(line: str) -> str:
    years = re.findall(r"\b(19\d{2}|20\d{2})\b", line or "")
    role = strip_date_suffix(line)
    if len(years) >= 2:
        return f"{role} ({years[0]}-{years[-1]})"
    if len(years) == 1:
        return f"{role} ({years[0]})"
    return role


def best_current_experience(experience: list[str], title: str, company: str) -> str:
    current = [x for x in experience if re.search(r"\bheden\b|\bpresent\b", x, re.I)]
    search = current or experience
    title_l, company_l = title.lower(), company.lower()
    for line in search:
        ll = line.lower()
        if title_l and title_l in ll:
            return line
    for line in search:
        ll = line.lower()
        if company_l and company_l in ll:
            return line
    return search[0] if search else ""


def parse_docx(uploaded_file) -> pd.DataFrame:
    doc = Document(uploaded_file)
    texts = [clean(p.text) for p in doc.paragraphs]
    texts = [t for t in texts if t]
    exp_idxs = [i for i, t in enumerate(texts) if t == "ErvaringErvaring in profiel"]

    records = []
    if not exp_idxs:
        return pd.DataFrame(columns=STANDARD_COLUMNS)

    headers = []
    for exp_idx in exp_idxs:
        if exp_idx < 4:
            continue
        sector = texts[exp_idx - 1][2:].strip() if texts[exp_idx - 1].startswith("· ") else ""
        location = texts[exp_idx - 2]
        headline = texts[exp_idx - 3]
        name_idx = exp_idx - 4
        while name_idx >= 0 and is_badge(texts[name_idx]):
            name_idx -= 1
        if texts[name_idx].endswith(" selecteren") and name_idx + 1 < len(texts):
            name_idx += 1
        headers.append((name_idx, texts[name_idx], headline, location, sector))

    for pos, (start, name, headline, location, sector) in enumerate(headers):
        end = headers[pos + 1][0] if pos + 1 < len(headers) else len(texts)
        block = texts[start:end]
        info = [x for x in block[1:] if not is_badge(x)]
        experience: list[str] = []
        education: list[str] = []
        interest: list[str] = []
        activity: list[str] = []
        section = None

        for line in info:
            if line.startswith("· "):
                continue
            if section is None and line in {headline, location}:
                continue
            if line == "ErvaringErvaring in profiel":
                section = "experience"
                continue
            if line == "OpleidingOpleidingen op het profiel":
                section = "education"
                continue
            if line.startswith("LabelsLabels") or line.startswith("Decoraties voor rij met profielinteresses"):
                section = "interest"
                continue
            if line.startswith("Decoraties voor rij met profielactiviteit"):
                section = "activity"
                continue
            if line.startswith("Alles weergeven"):
                continue
            if any(line.startswith(prefix) for prefix in STOP_PREFIXES):
                section = None
                continue
            if line.startswith("Kandidaat werd opgeslagen") or line.startswith("Opgeslagen door"):
                section = "activity"
                continue
            if section == "experience":
                experience.append(line)
            elif section == "education":
                education.append(line)
            elif section == "interest":
                interest.append(line)
            elif section == "activity":
                activity.append(line)

        title, company = split_title_company(headline)
        current = best_current_experience(experience, title, company)
        if current:
            current_title, current_company = split_title_company(strip_date_suffix(current))
            if current_company and not company:
                company = current_company
            if current_company and title.lower() == current_company.lower():
                title = current_title

        prior = [x for x in experience if x != current]
        if "Premium" in " ".join(block):
            interest.append("LinkedIn Premium")

        records.append(
            {
                "naam": name,
                "huidig_bedrijf": company,
                "functietitel": title,
                "locatie": location,
                "sector": sector,
                "tenure_huidige_rol_jaren": years_from_line(current),
                "info_eerdere_rollen": " | ".join(compact_role(x) for x in prior[:3]),
                "hoogst_afgeronde_opleiding": education_summary(education),
                "interesse_signalen": "; ".join(dict.fromkeys([x for x in interest if x])),
                "activiteit": " | ".join(activity[:4]),
            }
        )
    return pd.DataFrame(records, columns=STANDARD_COLUMNS)


def parse_docx(uploaded_file) -> pd.DataFrame:
    doc = Document(uploaded_file)
    texts = [clean(p.text) for p in doc.paragraphs]
    texts = [t for t in texts if t]
    exp_idxs = [i for i, t in enumerate(texts) if is_experience_marker(t)]

    records = []
    if not exp_idxs:
        return pd.DataFrame(columns=STANDARD_COLUMNS)

    headers = []
    for exp_idx in exp_idxs:
        if exp_idx < 3:
            continue

        scan_idx = exp_idx - 1
        sector = ""
        if scan_idx >= 0 and is_sector_line(texts[scan_idx]):
            sector = clean_sector_line(texts[scan_idx])
            scan_idx -= 1
        if scan_idx < 2:
            continue

        location = texts[scan_idx]
        headline = texts[scan_idx - 1]
        name_idx = scan_idx - 2
        while name_idx >= 0 and (is_badge(texts[name_idx]) or texts[name_idx].startswith("Select ")):
            name_idx -= 1
        if name_idx < 0:
            continue
        if texts[name_idx].endswith(" selecteren") and name_idx + 1 < len(texts):
            name_idx += 1
        headers.append((name_idx, texts[name_idx], headline, location, sector))

    for pos, (start, name, headline, location, sector) in enumerate(headers):
        end = headers[pos + 1][0] if pos + 1 < len(headers) else len(texts)
        block = texts[start:end]
        info = [x for x in block[1:] if not is_badge(x)]
        experience: list[str] = []
        education: list[str] = []
        interest: list[str] = []
        activity: list[str] = []
        section = None

        for line in info:
            if is_sector_line(line):
                continue
            if section is None and line in {headline, location}:
                continue
            if is_experience_marker(line):
                section = "experience"
                continue
            if is_education_marker(line):
                section = "education"
                continue
            if is_interest_marker(line):
                section = "interest"
                continue
            if is_activity_marker(line):
                section = "activity"
                continue
            if line.startswith(("Alles weergeven", "Show all", "Similar skills to saved candidates")):
                continue
            if any(line.startswith(prefix) for prefix in STOP_PREFIXES):
                section = None
                continue
            if line.startswith(("Kandidaat werd opgeslagen", "Opgeslagen door", "Candidate saved to project", "Saved by")):
                section = "activity"
                continue
            if section == "experience":
                experience.append(line)
            elif section == "education":
                education.append(line)
            elif section == "interest":
                interest.append(line)
            elif section == "activity":
                activity.append(line)

        title, company = split_title_company(headline)
        current = best_current_experience(experience, title, company)
        if current:
            current_title, current_company = split_title_company(strip_date_suffix(current))
            if current_company and not company:
                company = current_company
            if current_company and title.lower() == current_company.lower():
                title = current_title

        prior = [x for x in experience if x != current]
        if "Premium" in " ".join(block):
            interest.append("LinkedIn Premium")

        records.append(
            {
                "naam": name,
                "huidig_bedrijf": company,
                "functietitel": title,
                "locatie": location,
                "sector": sector,
                "tenure_huidige_rol_jaren": years_from_line(current),
                "info_eerdere_rollen": " | ".join(compact_role(x) for x in prior[:3]),
                "hoogst_afgeronde_opleiding": education_summary(education),
                "interesse_signalen": "; ".join(dict.fromkeys([x for x in interest if x])),
                "activiteit": " | ".join(activity[:4]),
            }
        )
    return pd.DataFrame(records, columns=STANDARD_COLUMNS)


def read_table(uploaded_file) -> pd.DataFrame:
    name = uploaded_file.name.lower()
    if name.endswith(".docx"):
        return parse_docx(uploaded_file)
    if name.endswith(".csv"):
        df = pd.read_csv(uploaded_file)
    else:
        df = pd.read_excel(uploaded_file)
    return normalize_columns(df)


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    aliases = {
        "naam": ["naam", "name", "candidate", "kandidaat"],
        "huidig_bedrijf": ["huidig bedrijf", "company", "bedrijf", "current company", "organisatie"],
        "functietitel": ["functietitel", "title", "functie", "job title", "headline"],
        "locatie": ["locatie", "location", "plaats"],
        "sector": ["sector", "industry", "branche"],
        "tenure_huidige_rol_jaren": ["tenure huidige rol (jaren)", "tenure", "tenure_huidige_rol_jaren"],
        "info_eerdere_rollen": ["info eerdere rollen", "previous roles", "ervaring"],
        "hoogst_afgeronde_opleiding": ["hoogst afgeronde opleiding", "education", "opleiding"],
        "interesse_signalen": ["interesse signalen", "interest", "signals"],
        "activiteit": ["activiteit", "activity"],
    }
    by_lower = {clean(c).lower(): c for c in df.columns}
    out = pd.DataFrame()
    for target, names in aliases.items():
        source = next((by_lower[n] for n in names if n in by_lower), None)
        out[target] = df[source] if source else ""
    return out[STANDARD_COLUMNS]


def classify_sector(row: pd.Series, sector_buckets: list[str], sector_rules: dict[str, list[str]] | None = None) -> str:
    if not sector_buckets:
        return clean(row.get("sector")) or "Overig"
    blob = " | ".join(clean(row.get(c)) for c in ["huidig_bedrijf", "functietitel", "info_eerdere_rollen"]).lower()
    original_sector = clean(row.get("sector")).lower()
    for bucket in sector_buckets:
        keywords = (sector_rules or {}).get(bucket, [])
        if any(keyword and keyword in blob for keyword in keywords):
            return bucket
    for bucket in sector_buckets:
        bucket_l = bucket.lower()
        bucket_words = [bucket_l, *bucket_l.replace("/", " ").split()]
        if any(word and word in blob for word in bucket_words):
            return bucket
        if original_sector == bucket_l:
            return bucket
    return "Overig" if "Overig" in sector_buckets else sector_buckets[-1]


def flag_outliers(df: pd.DataFrame, include_terms: list[str], exclude_terms: list[str]) -> pd.DataFrame:
    rows = []
    for idx, row in df.iterrows():
        blob = " | ".join(clean(row.get(c)) for c in STANDARD_COLUMNS).lower()
        include_hits = [t for t in include_terms if t in blob]
        exclude_hits = [t for t in exclude_terms if t in blob]
        reasons = []
        score = 0
        if include_terms and not include_hits:
            score += 2
            reasons.append("geen doelprofiel-keyword gevonden")
        if exclude_hits:
            score += min(3, len(exclude_hits))
            reasons.append("uitsluitingskeyword(s): " + ", ".join(exclude_hits[:5]))
        if not clean(row.get("huidig_bedrijf")) or not clean(row.get("functietitel")):
            score += 1
            reasons.append("bedrijf of functietitel ontbreekt")
        if score >= 2:
            rows.append(
                {
                    "verwijderen": False,
                    "index": idx,
                    "naam": row.get("naam", ""),
                    "huidig_bedrijf": row.get("huidig_bedrijf", ""),
                    "functietitel": row.get("functietitel", ""),
                    "sector": row.get("sector", ""),
                    "reden": "; ".join(reasons),
                }
            )
    return pd.DataFrame(rows)


def summary_frames(df: pd.DataFrame) -> dict[str, pd.DataFrame]:
    total = max(len(df), 1)
    frames = {
        "KPI": pd.DataFrame(
            [
                ["Profielen", len(df), 1],
                ["Met huidig bedrijf", df["huidig_bedrijf"].astype(str).str.len().gt(0).sum(), None],
                ["Met tenure", pd.to_numeric(df["tenure_huidige_rol_jaren"], errors="coerce").notna().sum(), None],
                ["Gemiddelde tenure", pd.to_numeric(df["tenure_huidige_rol_jaren"], errors="coerce").mean(), None],
                ["Met opleiding", df["hoogst_afgeronde_opleiding"].astype(str).str.len().gt(0).sum(), None],
            ],
            columns=["metriek", "waarde", "percentage"],
        )
    }
    for col, label in [("huidig_bedrijf", "Bedrijven"), ("sector", "Sectoren"), ("locatie", "Locaties")]:
        counts = df[col].replace("", "Onbekend").fillna("Onbekend").value_counts().head(20)
        frames[label] = pd.DataFrame({"categorie": counts.index, "aantal": counts.values, "percentage": counts.values / total})
    return frames


def make_excel(df: pd.DataFrame) -> bytes:
    output = BytesIO()
    excel_df = df[STANDARD_COLUMNS].copy().where(pd.notna(df[STANDARD_COLUMNS]), "")
    max_row = 2000
    total_formula = f'=COUNTA(Profielen!$A$2:$A${max_row})'
    top_companies = excel_df["huidig_bedrijf"].replace("", "Onbekend").fillna("Onbekend").value_counts().head(20).index.tolist()
    top_sectors = excel_df["sector"].replace("", "Onbekend").fillna("Onbekend").value_counts().head(20).index.tolist()
    top_locations = excel_df["locatie"].replace("", "Onbekend").fillna("Onbekend").value_counts().head(20).index.tolist()
    top_education = (
        excel_df["hoogst_afgeronde_opleiding"]
        .replace("", "Onbekend")
        .fillna("Onbekend")
        .value_counts()
        .head(20)
        .index.tolist()
    )
    tenure_buckets = ["<1 jaar", "1-2 jaar", "3-5 jaar", "6-10 jaar", "10+ jaar", "Onbekend"]

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        excel_df.to_excel(writer, sheet_name="Profielen", index=False)
        workbook = writer.book
        workbook.set_calc_mode("auto")
        ws = writer.sheets["Profielen"]
        header_fmt = workbook.add_format({"bold": True, "font_color": "white", "bg_color": "#155E75", "border": 1})
        wrap = workbook.add_format({"text_wrap": True, "valign": "top"})
        for col_num, col in enumerate(excel_df.columns):
            ws.write(0, col_num, col, header_fmt)
            width = min(max(14, int(excel_df[col].astype(str).str.len().quantile(0.9) if len(excel_df) else 14)), 42)
            ws.set_column(col_num, col_num, width, wrap)
        ws.set_column(5, 5, 16)
        ws.freeze_panes(1, 1)
        ws.autofilter(0, 0, len(excel_df), len(excel_df.columns) - 1)

        summary = workbook.add_worksheet("Samenvatting")
        title_fmt = workbook.add_format({"bold": True, "font_color": "white", "bg_color": "#155E75", "font_size": 14})
        sub_fmt = workbook.add_format({"bold": True, "bg_color": "#E0F2FE"})
        header_fmt_2 = workbook.add_format({"bold": True, "bg_color": "#E0F2FE", "border": 1})
        pct_fmt = workbook.add_format({"num_format": "0.0%"})
        num_fmt = workbook.add_format({"num_format": "0.0"})
        summary.write(0, 0, "Market mapping - samenvatting", title_fmt)
        summary.write(1, 0, "Deze tab rekent mee wanneer je waarden op tab Profielen aanpast.", None)

        summary.write_row(3, 0, ["Metriek", "Waarde", "Percentage", "Opmerking"], header_fmt_2)
        kpis = [
            ("Profielen", total_formula, '=IF($B$5=0,"",1)', "Telt namen op tab Profielen"),
            ("Met huidig bedrijf", f'=COUNTIF(Profielen!$B$2:$B${max_row},"<>")', '=IF($B$5=0,"",$B5/$B$5)', ""),
            ("Met tenure", f'=COUNT(Profielen!$F$2:$F${max_row})', '=IF($B$5=0,"",$B6/$B$5)', ""),
            ("Gemiddelde tenure", f'=IFERROR(AVERAGE(Profielen!$F$2:$F${max_row}),"")', "", "In jaren"),
            ("Met eerdere rollen", f'=COUNTIF(Profielen!$G$2:$G${max_row},"<>")', '=IF($B$5=0,"",$B8/$B$5)', ""),
            ("Met opleiding", f'=COUNTIF(Profielen!$H$2:$H${max_row},"<>")', '=IF($B$5=0,"",$B9/$B$5)', ""),
            ("Met activiteit", f'=COUNTIF(Profielen!$J$2:$J${max_row},"<>")', '=IF($B$5=0,"",$B10/$B$5)', ""),
        ]
        for offset, (label, value_formula, pct_formula, note) in enumerate(kpis, start=4):
            summary.write(offset, 0, label)
            summary.write_formula(offset, 1, value_formula, num_fmt if label == "Gemiddelde tenure" else None)
            if pct_formula:
                summary.write_formula(offset, 2, pct_formula, pct_fmt)
            summary.write(offset, 3, note)

        def write_count_block(start_row: int, title: str, labels: Iterable[str], source_col: str, chart: bool = True) -> int:
            labels = list(labels)
            summary.write(start_row, 0, title, sub_fmt)
            summary.write_row(start_row + 1, 0, ["categorie", "aantal", "percentage"], header_fmt_2)
            for i, label in enumerate(labels, start=start_row + 2):
                summary.write(i, 0, label)
                if label == "Onbekend":
                    formula = f'=COUNTIFS(Profielen!$A$2:$A${max_row},"<>",Profielen!${source_col}$2:${source_col}${max_row},"")'
                else:
                    safe_label = str(label).replace('"', '""')
                    formula = f'=COUNTIF(Profielen!${source_col}$2:${source_col}${max_row},"{safe_label}")'
                summary.write_formula(i, 1, formula)
                summary.write_formula(i, 2, f'=IF($B$5=0,"",$B{i + 1}/$B$5)', pct_fmt)
            if chart and labels:
                chart_obj = workbook.add_chart({"type": "column"})
                chart_end = start_row + 1 + min(len(labels), 10)
                chart_obj.add_series(
                    {
                        "name": title,
                        "categories": ["Samenvatting", start_row + 2, 0, chart_end, 0],
                        "values": ["Samenvatting", start_row + 2, 1, chart_end, 1],
                    }
                )
                chart_obj.set_title({"name": title})
                chart_obj.set_legend({"none": True})
                summary.insert_chart(start_row, 5, chart_obj, {"x_scale": 1.25, "y_scale": 1.1})
            return start_row + len(labels) + 4

        row = 13
        row = write_count_block(row, "Bedrijven", top_companies, "B", chart=False)
        row = write_count_block(row, "Sectoren", top_sectors, "E", chart=True)
        row = write_count_block(row, "Locaties", top_locations, "D", chart=True)
        row = write_count_block(row, "Opleiding", top_education, "H", chart=False)

        summary.write(row, 0, "Tenure-buckets", sub_fmt)
        summary.write_row(row + 1, 0, ["categorie", "aantal", "percentage"], header_fmt_2)
        tenure_formulas = [
            f'=COUNTIFS(Profielen!$F$2:$F${max_row},">=0",Profielen!$F$2:$F${max_row},"<1")',
            f'=COUNTIFS(Profielen!$F$2:$F${max_row},">=1",Profielen!$F$2:$F${max_row},"<=2")',
            f'=COUNTIFS(Profielen!$F$2:$F${max_row},">=3",Profielen!$F$2:$F${max_row},"<=5")',
            f'=COUNTIFS(Profielen!$F$2:$F${max_row},">=6",Profielen!$F$2:$F${max_row},"<=10")',
            f'=COUNTIF(Profielen!$F$2:$F${max_row},">10")',
            f'=$B$5-COUNT(Profielen!$F$2:$F${max_row})',
        ]
        for i, (label, formula) in enumerate(zip(tenure_buckets, tenure_formulas), start=row + 2):
            summary.write(i, 0, label)
            summary.write_formula(i, 1, formula)
            summary.write_formula(i, 2, f'=IF($B$5=0,"",$B{i + 1}/$B$5)', pct_fmt)

        summary.set_column(0, 0, 34)
        summary.set_column(1, 2, 14)
        summary.set_column(3, 3, 48)
    return output.getvalue()


def main() -> None:
    st.set_page_config(page_title="Market Mapping", layout="wide")
    st.title("Market Mapping")
    st.caption("Upload een LinkedIn Recruiter Word-export of een Excel/CSV en maak een nette profielmapping met twijfelcheck.")

    with st.sidebar:
        st.header("Instellingen")
        sector_preset = st.selectbox(
            "Sectorindeling",
            ["Automatisch", "Zorg", "Pharma / biotech"],
            index=0,
            help="Automatisch kiest zelf tussen zorg-buckets en pharma/biotech-buckets. De onderliggende keywordregels zitten verborgen in de app.",
        )
        extra_exclude_terms = tokens(
            st.text_area(
                "Extra twijfel-/uitsluitingswoorden",
                "",
                help="Optioneel. De app gebruikt al een standaardlijst. Vul hier alleen extra woorden in die voor jouw zoekopdracht verdacht zijn.",
            )
        )

    uploaded = st.file_uploader("Upload Word, Excel of CSV", type=["docx", "xlsx", "xls", "csv"])

    if not uploaded:
        st.info("Upload een bestand om te beginnen.")
        return

    df = read_table(uploaded)
    if df.empty:
        st.error("Ik kon geen profielblokken herkennen. Probeer een Excel/CSV of controleer of de Word-export LinkedIn-profielregels bevat.")
        st.stop()

    sector_buckets, sector_rules, resolved_preset = sector_config(sector_preset, df)
    if sector_buckets:
        df["sector"] = df.apply(lambda row: classify_sector(row, sector_buckets, sector_rules), axis=1)
        st.caption(f"Gebruikte sectorindeling: {resolved_preset}")
    include_terms = DEFAULT_INCLUDE_TERMS.get(resolved_preset, DEFAULT_INCLUDE_TERMS["Pharma / biotech"])
    exclude_terms = DEFAULT_EXCLUDE_TERMS + extra_exclude_terms

    st.subheader("1. Ingelezen profielen")
    st.write(f"{len(df)} profielen ingelezen.")
    edited_df = st.data_editor(df, use_container_width=True, num_rows="dynamic", height=420)

    st.subheader("2. Twijfelprofielen")
    suspects = flag_outliers(edited_df, include_terms, exclude_terms)
    if suspects.empty:
        st.success("Geen duidelijke twijfelprofielen gevonden met de huidige keywords.")
        final_df = edited_df.copy()
    else:
        st.write("Vink profielen aan die uit de export moeten.")
        edited_suspects = st.data_editor(suspects, use_container_width=True, hide_index=True, height=280)
        remove_idx = set(edited_suspects.loc[edited_suspects["verwijderen"], "index"].tolist())
        final_df = edited_df.drop(index=remove_idx).reset_index(drop=True)
        st.info(f"{len(remove_idx)} profiel(en) gemarkeerd voor verwijdering. Export bevat straks {len(final_df)} profielen.")

    st.subheader("3. Export")
    excel = make_excel(final_df)
    st.download_button(
        "Download market mapping Excel",
        data=excel,
        file_name="market_mapping_output.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    main()
