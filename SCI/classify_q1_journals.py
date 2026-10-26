"""从 Q1 期刊清单中筛选偏向 9 个领域的专业刊。

综合类、大类刊（Nature/Science/PNAS、综合化学/物理/材料等）不纳入。
一种期刊只归入最匹配的一类；多类同样特异则列入「交叉未分」。

调整方法：改 DOMAIN_KEYWORDS、FORCE_INCLUDE、FORCE_EXCLUDE、EXCLUDE_EXACT。

示例：
  python classify_q1_journals.py
"""

from __future__ import annotations

import argparse
import re
import sys
import unicodedata
from pathlib import Path

import pandas as pd

DEFAULT_Q1 = Path(r"D:\AI科技类内容图文匹配\题录\Q1.xlsx")
DEFAULT_OUT = Path(r"D:\AI科技类内容图文匹配\题录\Q1九类期刊.xlsx")

DOMAINS = [
    "集成电路",
    "高端科学仪器",
    "智能制造",
    "高端机器人",
    "人工智能",
    "智能网联新能源车",
    "商业航天",
    "生物医药",
    "生物育种",
]

# 综合 / 大类刊：整刊名匹配（规范化后）
EXCLUDE_EXACT = {
    "NATURE",
    "SCIENCE",
    "SCIENCE ADVANCES",
    "NATURE COMMUNICATIONS",
    "PROCEEDINGS OF THE NATIONAL ACADEMY OF SCIENCES OF THE UNITED STATES OF AMERICA",
    "NATIONAL SCIENCE REVIEW",
    "SCIENCE BULLETIN",
    "CHINESE SCIENCE BULLETIN CHINESE",
    "ADVANCED SCIENCE",
    "INNOVATION",
    "RESEARCH",
    "JOURNAL OF ADVANCED RESEARCH",
    "ENGINEERED SCIENCE",
    "ESCIENCE",
    "AGGREGATE",
    "MATTER",
    "MATERIALS TODAY",
    "ADVANCED MATERIALS",
    "ADVANCED FUNCTIONAL MATERIALS",
    "NATURE MATERIALS",
    "NATURE NANOTECHNOLOGY",
    "NATURE CHEMISTRY",
    "NATURE PHYSICS",
    "NATURE REVIEWS CHEMISTRY",
    "NATURE REVIEWS MATERIALS",
    "NATURE REVIEWS PHYSICS",
    "NATURE REVIEWS METHODS PRIMERS",
    "CHEMICAL REVIEWS",
    "CHEMICAL SOCIETY REVIEWS",
    "CHEMICAL SCIENCE",
    "CHEM",
    "ANGEWANDTE CHEMIE INTERNATIONAL EDITION",
    "JOURNAL OF THE AMERICAN CHEMICAL SOCIETY",
    "JACS AU",
    "ACS CENTRAL SCIENCE",
    "PHYSICAL REVIEW LETTERS",
    "PHYSICAL REVIEW X",
    "PHYSICS REPORTS REVIEW SECTION OF PHYSICS LETTERS",
    "REVIEWS OF MODERN PHYSICS",
    "ADVANCES IN PHYSICS",
    "PROCEEDINGS OF THE IEEE",
    "INTERDISCIPLINARY MATERIALS",
    "MATERIALS FUTURES",
    "PROGRESS IN MATERIALS SCIENCE",
    "INTERNATIONAL MATERIALS REVIEWS",
    "SCIENCE CHINA MATERIALS",
    "SCIENCE CHINA TECHNOLOGICAL SCIENCES",
    "SCIENCE CHINA PHYSICS MECHANICS ASTRONOMY",
    "SCIENCE CHINA CHEMISTRY",
    "CCS CHEMISTRY",
    "CHINESE CHEMICAL LETTERS",
    "CHINESE JOURNAL OF CHEMISTRY",
    "ACCOUNTS OF CHEMICAL RESEARCH",
    "ACCOUNTS OF MATERIALS RESEARCH",
    "COORDINATION CHEMISTRY REVIEWS",
    "TRENDS IN CHEMISTRY",
    "NATURE SYNTHESIS",
    "NATURE CATALYSIS",
    "CHEM CATALYSIS",
    "CHINESE JOURNAL OF CATALYSIS",
    "GREEN CHEMISTRY",
    "ACS NANO",
    "NANO RESEARCH",
    "NANO MICRO LETTERS",
    "NANO MATERIALS SCIENCE",
    "ACTA MATERIALIA",
    "ADVANCED COMPOSITES AND HYBRID MATERIALS",
    "INFOMAT",
    "SMARTMAT",
    "MATERIALS TODAY BIO",
    "COMMUNICATIONS CHEMISTRY",
    "COMMUNICATIONS PHYSICS",
    "COMMUNICATIONS BIOLOGY",
    "COMMUNICATIONS EARTH ENVIRONMENT",
    "ELIFE",
    "CURRENT BIOLOGY",
    "PLOS BIOLOGY",
    "CELL REPORTS",
    "SCIENCE CHINA LIFE SCIENCES",
    "BIOLOGICAL REVIEWS",
    "BIOSCIENCE",
    "PHILOSOPHICAL TRANSACTIONS OF THE ROYAL SOCIETY B BIOLOGICAL SCIENCES",
    "NATURE ECOLOGY EVOLUTION",
    "NATURE HUMAN BEHAVIOUR",
    "NATURE CLIMATE CHANGE",
    "NATURE SUSTAINABILITY",
    "NATURE GEOSCIENCE",
    "ACM COMPUTING SURVEYS",
    "COMPUTER SCIENCE REVIEW",
    "SCIENCE CHINA INFORMATION SCIENCES",
}

# 各领域刊名关键词（整词/词组）。一种刊按“最长命中词”归入一类。
DOMAIN_KEYWORDS: dict[str, list[str]] = {
    "集成电路": [
        "solid state circuit",
        "solid-state circuit",
        "integrated circuit",
        "electron device",
        "microelectron",
        "nanoelectron",
        "semiconductor",
        "flexible electronic",
        "nature electronics",
        "microsystems",
        "nanoengineering",
        "superconductor",
        "superconductivity",
        "microwave theory",
        "photonics",
        "opto electronic",
        "opto-electronic",
        "quantum information",
        "quantum materials",
        "acs photonics",
        "advanced photonics",
        "light science",
        "elight",
        "optica",
        "photonics research",
        "photonix",
        "laser photonics",
        "optics and photonics",
        "advances in optics",
        "circuits and systems",
        "npj flexible electronics",
        "prx quantum",
    ],
    "高端科学仪器": [
        "analytical chemistry",
        "nature methods",
        "sensors and actuators",
        "biosensor",
        "bioelectronics",
        "medical imaging",
        "medical image",
        "radiology",
        "nuclear medicine",
        "photoacoustic",
        "photoacoustics",
        "radiographics",
        "photonic sensor",
        "endoscopic ultrasound",
        "investigative radiology",
        "analysis and testing",
        "trac trends in analytical",
        "high power laser",
        "acs sensors",
        "measurement",
        "instrumentation",
        "microscopy",
        "spectroscopy",
        "mass spectrometry",
        "chromatography",
        "lab on a chip",
        "sensors",
    ],
    "智能制造": [
        "additive manufacturing",
        "extreme manufacturing",
        "manufacturing systems",
        "manufacturing processes",
        "manufacturing technology",
        "machine tools",
        "machine tool",
        "materials processing technology",
        "computer integrated manufacturing",
        "computers in industry",
        "industrial informatics",
        "industrial electronics",
        "industrial information",
        "automation in construction",
        "virtual and physical prototyping",
        "bio design and manufacturing",
        "engineering informatics",
        "automatica sinica",
        "mechatronics",
        "mechanical systems and signal",
        "mechanism and machine theory",
        "mechanical engineering",
        "mechanical sciences",
        "plasticity",
        "hydromechatronics",
        "production economics",
        "composites part b",
        "advanced industrial and engineering polymer",
        "systems and control",
    ],
    "高端机器人": [
        "robotics",
        "robot",
        "bionic",
        "cyborg",
        "autonomous systems",
        "neuroengineering",
        "human robot",
    ],
    "人工智能": [
        "artificial intelligence",
        "machine learning",
        "neural networks and learning",
        "neural network",
        "neural systems",
        "pattern analysis and machine intelligence",
        "pattern recognition",
        "knowledge based systems",
        "knowledge and data engineering",
        "evolutionary computation",
        "fuzzy systems",
        "fuzzy sets",
        "affective computing",
        "intelligence technology",
        "information fusion",
        "expert systems",
        "nature machine intelligence",
        "nature computational science",
        "big data mining",
        "data science and engineering",
        "image processing",
        "cybernetics",
        "foundations and trends in machine",
        "foundations and trends in information retrieval",
        "caai transactions",
        "computational science",
    ],
    "智能网联新能源车": [
        "intelligent vehicles",
        "transportation electrification",
        "etransportation",
        "e transportation",
        "electric vehicle",
        "energy storage",
        "power electronics",
        "acs energy letters",
        "advanced energy materials",
        "electrochemical energy",
        "energy material",
        "nano research energy",
        "carbon energy",
        "communications in transportation research",
        "transportation research part c",
        "vehicular",
        "automobile",
        "automotive",
        "battery",
        "lithium",
    ],
    "商业航天": [
        "aerospace",
        "aeronautics",
        "astronautics",
        "space science",
        "satellite navigation",
        "gps solutions",
        "progress in aerospace",
    ],
    "生物医药": [
        "lancet",
        "jama",
        "bmj",
        "new england journal of medicine",
        "nature medicine",
        "nature biotechnology",
        "nature biomedical",
        "nature reviews drug",
        "nature reviews cancer",
        "nature reviews immunology",
        "nature reviews molecular",
        "nature reviews microbiology",
        "nature reviews genetics",
        "nature reviews neuroscience",
        "nature reviews disease",
        "nature reviews clinical",
        "nature reviews cardiology",
        "nature reviews endocrinology",
        "nature reviews gastroenterology",
        "nature reviews nephrology",
        "nature reviews neurology",
        "nature reviews rheumatology",
        "nature reviews urology",
        "nature cancer",
        "nature immunology",
        "nature neuroscience",
        "nature microbiology",
        "nature metabolism",
        "nature aging",
        "nature cardiovascular",
        "molecular therapy",
        "molecular medicine",
        "molecular cancer",
        "clinical",
        "medicine",
        "medical",
        "pharma",
        "drug",
        "therapeut",
        "oncolog",
        "cancer",
        "tumor",
        "tumour",
        "immunolog",
        "vaccine",
        "patholog",
        "hematolog",
        "haematolog",
        "cardiol",
        "cardiac",
        "circulation",
        "heart",
        "neurolog",
        "neurosurg",
        "neuron",
        "brain",
        "psychiatr",
        "diabetes",
        "endocrin",
        "hepatol",
        "gastroenter",
        "nephrol",
        "kidney",
        "urolog",
        "ophthalm",
        "dermatol",
        "rheumat",
        "allerg",
        "respirator",
        "thorax",
        "chest",
        "anesthes",
        "anaesthes",
        "surgic",
        "surgery",
        "transplant",
        "biomaterial",
        "biomedical",
        "pharmaceut",
        "pharmacol",
        "toxicolog",
        "vaccine",
        "virolog",
        "microbiome",
        "stem cell",
        "regenerative",
        "gene therapy",
        "nucleic acid",
        "antibody",
        "peptide",
        "vaccine",
        "infectious",
        "pathogen",
        "blood",
        "leukemia",
        "leukaemia",
        "cell death",
        "autophagy",
        "apoptosis",
        "immunity",
        "inflammation",
        "alzheimer",
        "parkinson",
        "stroke",
        "pain",
        "obesity",
        "nutrition",
        "metaboli",
        "physiol",
        "anatomy",
        "radiat",
        "chemotherapy",
        "immunotherapy",
        "theranostic",
        "controlled release",
        "drug delivery",
        "acta pharmaceutica",
        "medicinal chemistry",
        "organic letters",
        "ebiomedicine",
        "eclinicalmedicine",
        "gut",
        "embo",
        "cell chemical biology",
        "cell host",
        "cell metabolism",
        "cell stem cell",
        "cell systems",
        "cell research",
        "cell discovery",
        "cell genomics",
        "cell proliferation",
        "cancer cell",
        "molecular cell",
        "developmental cell",
        "nature cell biology",
        "trends in molecular medicine",
        "trends in pharmacological",
        "trends in cancer",
        "trends in immunology",
        "trends in neurosciences",
        "trends in biotechnology",
        "trends in biochemical",
        "trends in cell biology",
        "trends in genetics",
        "trends in microbiology",
        "trends in endocrinology",
        "signal transduction",
        "experimental and molecular medicine",
        "journal of clinical investigation",
        "journal of experimental medicine",
        "genome medicine",
        "genome research",
        "genome biology",
        "nature genetics",
        "nature structural",
        "redox biology",
        "matrix biology",
        "aging cell",
        "ageing research",
        "npj digital medicine",
        "npj regenerative",
        "npj precision oncology",
        "npj parkinsons",
        "science translational medicine",
        "science immunology",
        "science signaling",
        "plos medicine",
        "plos pathogens",
        "bmc medicine",
        "military medical",
        "medcomm",
        "biomacromolecules",
        "acta biomaterialia",
        "bioactive materials",
        "journal of nanobiotechnology",
        "biosensors bioelectronics",
        "angiotensin",
        "angiogenesis",
        "thrombosis",
        "vascular",
        "arteriosclerosis",
        "hypertension",
        "fertility",
        "reproduction",
        "obstetrics",
        "gynecology",
        "pediatric",
        "neonat",
        "dental",
        "periodontol",
        "endodontic",
        "implant",
        "orthopaedic",
        "orthopedic",
        "bone and joint",
        "osteoporosis",
        "spine",
        "neuro oncology",
        "brain stimulation",
        "movement disorders",
        "epilepsia",
        "schizophrenia",
        "autism",
        "sleep medicine",
        "intensive care",
        "critical care",
        "emergency medicine",
        "resuscitation",
        "mmwr",
        "ca a cancer",
        "oncogene",
        "gut microbes",
        "cellular molecular immunology",
        "cellular molecular biology",
        "cellular and molecular gastroenterology",
        "cold spring harbor",
        "genes development",
        "is me journal",
        "isme journal",
        "fem s microbiology",
        "fems microbiology",
        "emerging microbes",
        "virulence",
        "protein cell",
        "molecular systems biology",
        "synthetic and systems biotechnology",
        "biotechnology advances",
        "annual review of medicine",
        "annual review of immunology",
        "annual review of biochemistry",
        "annual review of pathology",
        "annual review of pharmacology",
        "annual review of physiology",
        "annual review of neuroscience",
        "annual review of biomedical",
        "annual review of biophysics",
        "annual review of cell and developmental",
        "annual review of genetics",
        "annual review of genomics",
        "annual review of microbiology",
        "annual review of nutrition",
        "annual review of virology",
        "cell",
        "lancet",
    ],
    "生物育种": [
        "plant biotechnology",
        "theoretical and applied genetics",
        "genetics selection evolution",
        "nature plants",
        "molecular plant",
        "plant cell",
        "plant journal",
        "plant physiology",
        "plant communications",
        "plant diversity",
        "plant phenomics",
        "new phytologist",
        "trends in plant science",
        "annual review of plant",
        "annual review of phytopathology",
        "horticulture",
        "horticultural",
        "agronomy",
        "crop",
        "rice",
        "field crops",
        "integrative agriculture",
        "integrative plant",
        "industrial crops",
        "pest management",
        "pesticide",
        "phytopathology",
        "plant pathology",
        "entomologia",
        "insect science",
        "insect biochemistry",
        "animal science and biotechnology",
        "annual review of animal",
        "animal nutrition",
        "artificial intelligence in agriculture",
        "computers and electronics in agriculture",
        "information processing in agriculture",
        "agricultural systems",
        "european journal of agronomy",
        "journal of dairy science",
        "aquaculture",
        "fisheries",
        "fish and fisheries",
        "reviews in fish",
        "zoological",
        "breeding",
        "seed",
        "germplasm",
        "cultivar",
        "livestock",
        "poultry",
        "soybean",
        "wheat",
        "maize",
        "corn",
        "barley",
        "genome editing",
        "crispr",
        "transgenic",
        "plant",
        "agriculture",
        "agricultural",
        "agri",
        "crop",
        "animal",
    ],
}

DOMAIN_PREFIXES: dict[str, list[str]] = {
    "生物医药": [
        "pharma",
        "oncolog",
        "cardiol",
        "neurolog",
        "immunolog",
        "hematolog",
        "haematolog",
        "therapeut",
        "patholog",
        "psychiatr",
        "anaesth",
        "anesth",
        "gastroenter",
        "ophthalm",
        "dermatol",
        "rheumat",
        "nephrol",
        "transplant",
        "toxicol",
        "physiol",
        "pediatr",
        "gynecol",
        "obstetric",
        "urolog",
        "endocrin",
        "hepatol",
        "biomater",
        "biomedic",
        "pharmaceut",
        "medicin",
        "neuropath",
        "photoacoustic",
        "theranostic",
        "virolog",
    ],
}

FORCE_INCLUDE: dict[str, list[str]] = {
    "集成电路": [
        "IEEE JOURNAL OF SOLID-STATE CIRCUITS",
        "Nature Electronics",
        "npj Flexible Electronics",
        "Microsystems & Nanoengineering",
        "npj Quantum Information",
        "PRX Quantum",
        "ACS Photonics",
        "Nature Photonics",
    ],
    "高端科学仪器": [
        "ANALYTICAL CHEMISTRY",
        "NATURE METHODS",
        "ACS Sensors",
        "BIOSENSORS & BIOELECTRONICS",
        "SENSORS AND ACTUATORS B-CHEMICAL",
        "IEEE TRANSACTIONS ON MEDICAL IMAGING",
        "MEDICAL IMAGE ANALYSIS",
        "Radiology-Artificial Intelligence",
        "TRAC-TRENDS IN ANALYTICAL CHEMISTRY",
        "Photoacoustics",
        "RADIOGRAPHICS",
    ],
    "智能制造": [
        "Additive Manufacturing",
        "INTERNATIONAL JOURNAL OF MACHINE TOOLS & MANUFACTURE",
        "JOURNAL OF MANUFACTURING SYSTEMS",
        "ROBOTICS AND COMPUTER-INTEGRATED MANUFACTURING",
        "IEEE Transactions on Industrial Informatics",
        "IEEE TRANSACTIONS ON INDUSTRIAL ELECTRONICS",
        "IEEE-ASME TRANSACTIONS ON MECHATRONICS",
        "IEEE-CAA Journal of Automatica Sinica",
        "Virtual and Physical Prototyping",
        "International Journal of Extreme Manufacturing",
        "INTERNATIONAL JOURNAL OF PLASTICITY",
        "INTERNATIONAL JOURNAL OF MECHANICAL SCIENCES",
        "International Journal of Hydromechatronics",
    ],
    "高端机器人": [
        "IEEE Transactions on Robotics",
        "Science Robotics",
        "Annual Review of Control Robotics and Autonomous Systems",
        "Cyborg and Bionic Systems",
        "Journal of NeuroEngineering and Rehabilitation",
    ],
    "人工智能": [
        "Nature Machine Intelligence",
        "IEEE TRANSACTIONS ON PATTERN ANALYSIS AND MACHINE INTELLIGENCE",
        "IEEE Transactions on Neural Networks and Learning Systems",
        "Foundations and Trends in Machine Learning",
        "ARTIFICIAL INTELLIGENCE REVIEW",
        "ENGINEERING APPLICATIONS OF ARTIFICIAL INTELLIGENCE",
        "PATTERN RECOGNITION",
        "KNOWLEDGE-BASED SYSTEMS",
        "Information Fusion",
        "IEEE Transactions on Cybernetics",
        "IEEE TRANSACTIONS ON IMAGE PROCESSING",
        "IEEE TRANSACTIONS ON KNOWLEDGE AND DATA ENGINEERING",
        "CAAI Transactions on Intelligence Technology",
        "Nature Computational Science",
        "FUZZY SETS AND SYSTEMS",
        "INFORMATION PROCESSING & MANAGEMENT",
    ],
    "智能网联新能源车": [
        "IEEE Transactions on Intelligent Vehicles",
        "IEEE Transactions on Transportation Electrification",
        "eTransportation",
        "TRANSPORTATION RESEARCH PART C-EMERGING TECHNOLOGIES",
        "Energy Storage Materials",
        "ACS Energy Letters",
        "ADVANCED ENERGY MATERIALS",
        "IEEE TRANSACTIONS ON POWER ELECTRONICS",
        "Communications in Transportation Research",
    ],
    "商业航天": [
        "Aerospace Science and Technology",
        "PROGRESS IN AEROSPACE SCIENCES",
        "Chinese Journal of Aeronautics",
        "Space: Science & Technology",
        "Satellite Navigation",
        "GPS SOLUTIONS",
    ],
    "生物医药": [
        "CELL",
        "NATURE MEDICINE",
        "NATURE BIOTECHNOLOGY",
        "Nature Biomedical Engineering",
        "NEW ENGLAND JOURNAL OF MEDICINE",
        "LANCET",
        "JAMA-JOURNAL OF THE AMERICAN MEDICAL ASSOCIATION",
        "BMJ-British Medical Journal",
        "Science Translational Medicine",
        "ADVANCED DRUG DELIVERY REVIEWS",
        "BIOMATERIALS",
        "Acta Biomaterialia",
        "Med",
        "MEDICINAL RESEARCH REVIEWS",
        "Asian Journal of Pharmaceutical Sciences",
        "Journal of Pharmaceutical Analysis",
        "Theranostics",
        "Journal of Extracellular Vesicles",
        "ALLERGY",
        "ANAESTHESIA",
        "ANESTHESIOLOGY",
        "AMERICAN JOURNAL OF TRANSPLANTATION",
        "ALIMENTARY PHARMACOLOGY & THERAPEUTICS",
        "Biological Psychiatry",
        "INTERNATIONAL JOURNAL OF RADIATION ONCOLOGY BIOLOGY PHYSICS",
        "JOURNAL OF THE AMERICAN COLLEGE OF CARDIOLOGY",
        "International Journal of Oral Science",
        "Acta Neuropathologica",
        "ACTA NEUROPATHOLOGICA",
        "Acta Physiologica",
        "iMeta",
        "PHYTOMEDICINE",
        "NATURAL PRODUCT REPORTS",
        "HAEMATOLOGICA",
        "Gastroenterology",
        "HEPATOLOGY",
        "OPHTHALMOLOGY",
        "THYROID",
        "MODERN PATHOLOGY",
        "World Psychiatry",
        "Military Medical Research",
    ],
    "生物育种": [
        "THEORETICAL AND APPLIED GENETICS",
        "PLANT BIOTECHNOLOGY JOURNAL",
        "Nature Plants",
        "Molecular Plant",
        "PLANT CELL",
        "NEW PHYTOLOGIST",
        "Crop Journal",
        "GENETICS SELECTION EVOLUTION",
        "Artificial Intelligence in Agriculture",
        "COMPUTERS AND ELECTRONICS IN AGRICULTURE",
        "Journal of Animal Science and Biotechnology",
        "Horticulture Research",
        "Rice",
        "TRENDS IN PLANT SCIENCE",
        "Annual Review of Plant Biology",
    ],
}

FORCE_EXCLUDE: set[str] = {
    "ACADEMIC MEDICINE",
    "MEDICAL EDUCATION",
    "Perspectives on Medical Education",
    "BMC Medical Ethics",
    "NURSING ETHICS",
    "INTERNATIONAL JOURNAL OF NURSING STUDIES",
    "International Journal of Nursing Sciences",
    "FOOD POLICY",
    "Food Quality and Preference",
    "TOURISM MANAGEMENT",
    "ANNALS OF TOURISM RESEARCH",
    "ORGANIC LETTERS",
    "ACS Catalysis",
    "ACS Sustainable Chemistry & Engineering",
    "ACS Macro Letters",
    "JOURNAL OF CATALYSIS",
    "Applied Catalysis B-Environment and Energy",
    "Green Synthesis and Catalysis",
    "Nature Astronomy",
    "Annual Review of Astronomy and Astrophysics",
    "Astrophysical Journal Letters",
    "ASTRONOMY AND ASTROPHYSICS REVIEW",
    "Living Reviews in Solar Physics",
    "Living Reviews in Relativity",
    "NUCLEAR FUSION",
    "Annual Review of Earth and Planetary Sciences",
    "EARTH AND PLANETARY SCIENCE LETTERS",
    "Nature Food",
    "FOOD CHEMISTRY",
    "Food Chemistry",
    "Food Chemistry-X",
    "TRENDS IN FOOD SCIENCE & TECHNOLOGY",
    "COMPREHENSIVE REVIEWS IN FOOD SCIENCE AND FOOD SAFETY",
    "Annual Review of Food Science and Technology",
    "JOURNAL OF AGRICULTURAL AND FOOD CHEMISTRY",
    "AGRICULTURAL AND FOREST METEOROLOGY",
    "AGRICULTURE ECOSYSTEMS & ENVIRONMENT",
    "Agricultural Water Management",
    "Global Food Security-Agriculture Policy Economics and Environment",
    "ZOOLOGICAL RESEARCH",
    "INTEGRATIVE ZOOLOGY",
    "Avian Research",
    "Animal",
    "FISH AND FISHERIES",
    "Reviews in Aquaculture",
    "REVIEWS IN FISH BIOLOGY AND FISHERIES",
    "Reviews in Fisheries Science & Aquaculture",
    "Journal of Plant Ecology",
    "BRITISH JOURNAL OF EDUCATIONAL TECHNOLOGY",
    "COMPUTERS & EDUCATION",
    "COMPUTERS IN HUMAN BEHAVIOR",
    "IEEE TRANSACTIONS ON SOFTWARE ENGINEERING",
    "ACM TRANSACTIONS ON SOFTWARE ENGINEERING AND METHODOLOGY",
    "ACM Transactions on Computer-Human Interaction",
    "ACM TRANSACTIONS ON GRAPHICS",
    "IEEE TRANSACTIONS ON MULTIMEDIA",
    "IEEE TRANSACTIONS ON MOBILE COMPUTING",
    "IEEE JOURNAL ON SELECTED AREAS IN COMMUNICATIONS",
    "IEEE Communications Surveys and Tutorials",
    "IEEE TRANSACTIONS ON WIRELESS COMMUNICATIONS",
    "IEEE Transactions on Cognitive Communications and Networking",
    "IEEE TRANSACTIONS ON ANTENNAS AND PROPAGATION",
    "IEEE Transactions on Information Forensics and Security",
    "IEEE TRANSACTIONS ON POWER SYSTEMS",
    "IEEE Transactions on Smart Grid",
    "IEEE Transactions on Sustainable Energy",
    "IEEE TRANSACTIONS ON CIRCUITS AND SYSTEMS FOR VIDEO TECHNOLOGY",
    "TRANSPORTATION RESEARCH PART A-POLICY AND PRACTICE",
    "TRANSPORTATION RESEARCH PART B-METHODOLOGICAL",
    "TRANSPORTATION RESEARCH PART D-TRANSPORT AND ENVIRONMENT",
    "TRANSPORTATION RESEARCH PART E-LOGISTICS AND TRANSPORTATION REVIEW",
    "TRANSPORT REVIEWS",
    "Transport Policy",
    "ACCIDENT ANALYSIS AND PREVENTION",
    "INTERNATIONAL JOURNAL OF PRODUCTION ECONOMICS",
    "CLINICAL PSYCHOLOGY REVIEW",
    "JOURNAL OF CONSULTING AND CLINICAL PSYCHOLOGY",
    "JOURNAL OF CLINICAL CHILD AND ADOLESCENT PSYCHOLOGY",
    "CLINICAL CHILD AND FAMILY PSYCHOLOGY REVIEW",
    "Annual Review of Clinical Psychology",
    "Policy Insights from the Behavioral and Brain Sciences",
    "BEHAVIORAL AND BRAIN SCIENCES",
    "APPLIED ENERGY",
    "Energy",
    "Nature Energy",
    "Joule",
    "ENERGY CONVERSION AND MANAGEMENT",
    "Energy & Environmental Science",
    "Energy Economics",
    "Green Energy & Environment",
    "RENEWABLE & SUSTAINABLE ENERGY REVIEWS",
    "Renewable Energy",
    "Advances in Applied Energy",
    "Progress in Energy and Combustion Science",
}


def normalize(text: object) -> str:
    if text is None or (isinstance(text, float) and pd.isna(text)):
        return ""
    s = unicodedata.normalize("NFKC", str(text)).upper()
    s = s.replace("&", " ").replace("'", " ")
    s = re.sub(r"[^A-Z0-9]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def contains_keyword(text: str, keyword: str) -> bool:
    t = f" {normalize(text)} "
    k = f" {normalize(keyword)} "
    return bool(k.strip()) and k in t


def prefix_hits(text: str, prefixes: list[str]) -> list[str]:
    tokens = normalize(text).split()
    hits = []
    for prefix in prefixes:
        p = normalize(prefix)
        if len(p) < 5:
            continue
        if any(tok == p or tok.startswith(p) for tok in tokens):
            hits.append(prefix)
    return hits


def force_map() -> dict[str, str]:
    mapping: dict[str, str] = {}
    for domain, names in FORCE_INCLUDE.items():
        for name in names:
            mapping[normalize(name)] = domain
    return mapping


def exclude_set() -> set[str]:
    names = {normalize(x) for x in EXCLUDE_EXACT}
    names.update(normalize(x) for x in FORCE_EXCLUDE)
    names.discard("")
    return names


def classify_journal(name: str) -> tuple[str, str]:
    key = normalize(name)
    if not key:
        return "", ""
    if key in exclude_set():
        return "", "综合/排除"
    forced = force_map().get(key)
    if forced:
        return forced, "强制纳入"

    best_domain = ""
    best_kw = ""
    best_len = -1
    ties: list[str] = []
    for domain in DOMAINS:
        for kw in DOMAIN_KEYWORDS.get(domain, []):
            if not contains_keyword(name, kw):
                continue
            score = len(normalize(kw))
            if score > best_len:
                best_len = score
                best_domain = domain
                best_kw = kw
                ties = [domain]
            elif score == best_len and domain not in ties:
                ties.append(domain)
        for prefix in DOMAIN_PREFIXES.get(domain, []):
            if not prefix_hits(name, [prefix]):
                continue
            score = len(normalize(prefix))
            if score > best_len:
                best_len = score
                best_domain = domain
                best_kw = prefix + "*"
                ties = [domain]
            elif score == best_len and domain not in ties:
                ties.append(domain)
    if len(ties) > 1:
        return "", "交叉:" + "；".join(ties) + f"（{best_kw}）"
    if best_domain:
        return best_domain, best_kw
    return "", ""


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="按 9 类筛选 Q1 专业期刊（不含综合刊）")
    parser.add_argument("--q1", type=Path, default=DEFAULT_Q1)
    parser.add_argument("--out", type=Path, default=DEFAULT_OUT)
    return parser.parse_args()


def main() -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")

    args = parse_args()
    src = pd.read_excel(args.q1)
    if "Journal" not in src.columns:
        raise SystemExit(f"Q1 表缺少 Journal 列: {args.q1}")

    rows = []
    for rec in src.to_dict("records"):
        journal = str(rec.get("Journal", "")).strip()
        domain, reason = classify_journal(journal)
        rows.append(
            {
                "Journal": journal,
                "分区": rec.get("分区"),
                "Top": rec.get("Top"),
                "Open Access": rec.get("Open Access"),
                "出现次数": rec.get("出现次数"),
                "领域": domain,
                "匹配说明": reason,
            }
        )
    df = pd.DataFrame(rows)

    summary = []
    for domain in DOMAINS:
        part = df[df["领域"] == domain]
        summary.append(
            {
                "领域": domain,
                "期刊数": int(len(part)),
                "出现次数合计": pd.to_numeric(part["出现次数"], errors="coerce").sum(min_count=1),
            }
        )
    unused_mask = (df["领域"] == "") & (~df["匹配说明"].astype(str).str.startswith("交叉"))
    cross_mask = df["匹配说明"].astype(str).str.startswith("交叉")
    excl_mask = df["匹配说明"] == "综合/排除"
    miss_mask = (df["领域"] == "") & (df["匹配说明"] == "")
    summary.append({"领域": "交叉未分", "期刊数": int(cross_mask.sum()), "出现次数合计": pd.to_numeric(df.loc[cross_mask, "出现次数"], errors="coerce").sum(min_count=1)})
    summary.append({"领域": "综合/排除", "期刊数": int(excl_mask.sum()), "出现次数合计": pd.to_numeric(df.loc[excl_mask, "出现次数"], errors="coerce").sum(min_count=1)})
    summary.append({"领域": "未匹配", "期刊数": int(miss_mask.sum()), "出现次数合计": pd.to_numeric(df.loc[miss_mask, "出现次数"], errors="coerce").sum(min_count=1)})
    summary_df = pd.DataFrame(summary)

    args.out.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(args.out, engine="xlsxwriter") as writer:
        summary_df.to_excel(writer, sheet_name="汇总", index=False)
        for domain in DOMAINS:
            part = df[df["领域"] == domain].drop(columns=["领域"]).reset_index(drop=True)
            if "出现次数" in part.columns:
                part = part.sort_values("出现次数", ascending=False, na_position="last").reset_index(drop=True)
            drop_cols = [c for c in ("分区", "Top", "Open Access") if c in part.columns and part[c].isna().all()]
            if drop_cols:
                part = part.drop(columns=drop_cols)
            part.insert(0, "序号", range(1, len(part) + 1))
            part.to_excel(writer, sheet_name=domain[:31], index=False)
        cross = df[df["匹配说明"].astype(str).str.startswith("交叉")].copy()
        unused = df[(df["领域"] == "") & (~df["匹配说明"].astype(str).str.startswith("交叉"))].copy()
        cross.to_excel(writer, sheet_name="交叉未分", index=False)
        unused.to_excel(writer, sheet_name="未纳入", index=False)

    print(f"Q1 期刊: {len(df)}")
    for rec in summary:
        print(f"  {rec['领域']}: {rec['期刊数']}")
    print(f"已保存: {args.out.resolve()}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
