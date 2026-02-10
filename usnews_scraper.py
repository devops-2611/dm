import requests
from bs4 import BeautifulSoup
from urllib.parse import unquote
import re
import csv
import time
import pandas as pd
from datetime import datetime

headers = {
    "User-Agent": "Mozilla/5.0"
}

# -----------------------------------
# READ QUERIES FROM EXCEL
# -----------------------------------
excel_file = "only_school_names.xlsx"
df = pd.read_excel(excel_file, header=None)
queries = df[0].dropna().tolist()

print(f"✅ Loaded {len(queries)} queries")

results = []

# -----------------------------------
# FIND US NEWS LINK (K-12 + HIGH SCHOOL)
# -----------------------------------


def find_usnews_link(query):
    search_url = "https://html.duckduckgo.com/html/"
    resp = requests.get(search_url, params={"q": query}, headers={
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/117.0.0.0 Safari/537.36"
        ),
        "Accept-Language": "en-US,en;q=0.9",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
        "Referer": "https://duckduckgo.com/",
    })
    soup = BeautifulSoup(resp.text, "html.parser")

    for a in soup.find_all("a", href=True):
        href = a["href"]

        # Check direct usnews.com links (some links might be direct)
        if "usnews.com/education/k12/" in href or "usnews.com/education/best-high-schools/" in href:
            clean_url = href.split("&uddg=")[0].split("?uddg=")[0]
            if clean_url.startswith("https://www.usnews.com/education/"):
                return clean_url

        # Check if link contains uddg= parameter (redirect link)
        if "uddg=" in href:
            try:
                decoded = unquote(href.split("uddg=")[1])
            except IndexError:
                continue

            if (
                "usnews.com/education/k12/" in decoded or
                "usnews.com/education/best-high-schools/" in decoded
            ):
                clean_url = re.split(r"[&?]", decoded)[0]
                if clean_url.startswith("https://www.usnews.com/education/"):
                    return clean_url

    return None




# -----------------------------------
# EXTRACT DATA
# -----------------------------------
def extract_data(url, school_name):
    page = requests.get(url, headers=headers)
    soup = BeautifulSoup(page.text, "html.parser")
    text = soup.get_text("\n")

    # -----------------------------------
    # HIGH SCHOOL
    # -----------------------------------
    if "best-high-schools" in url:
        # ---- Overview ----
        match = re.search(
            r"Overview of .*?\n(.*?)\nAll Rankings",
            text,
            re.S
        )
        overview = match.group(1) if match else ""

        # remove rankings text inside overview
        overview = re.split(r"\n.*?Rankings", overview, 1)[0]

        overview = " ".join(
            line.strip()
            for line in overview.splitlines()
            if line.strip()
        )

        # ---- Metrics ----
        math = re.search(r"Mathematics Proficiency\s*\n\s*(\d+%)", text)
        reading = re.search(r"Reading Proficiency\s*\n\s*(\d+%)", text)
        science = re.search(r"Science Proficiency\s*\n\s*(\d+%)", text)
        grad = re.search(r"Graduation Rate\s*\n\s*(\d+%)", text)

        return (
            overview,
            "",  # student-teacher ratio (usually not present)
            math.group(1) if math else "",
            reading.group(1) if reading else "",
            science.group(1) if science else "",
            grad.group(1) if grad else "",
            "High School"
        )

    # -----------------------------------
    # K-12 (Elementary / Middle)
    # -----------------------------------
    match = re.search(
        r"Overview of .*?\n(.*?)\nAt a Glance",
        text,
        re.S
    )
    overview = match.group(1) if match else ""

    lines = [l.strip() for l in overview.splitlines() if l.strip()]
    if lines and lines[0].lower() == school_name.lower():
        lines.pop(0)

    overview = " ".join(lines)

    ratio = re.search(r"Student/Teacher Ratio\s*\n\s*([\d:]+)", text)
    math = re.search(r"Math Proficiency\s*\n\s*(\d+%)", text)
    reading = re.search(r"Reading Proficiency\s*\n\s*(\d+%)", text)

    return (
        overview,
        ratio.group(1) if ratio else "",
        math.group(1) if math else "",
        reading.group(1) if reading else "",
        "",     # science (not available)
        "",     # graduation (not available)
        "K-12"
    )


# -----------------------------------
# MAIN LOOP
# -----------------------------------
for query in queries:
    print(f"\n🔍 Processing: {query}")

    if "US news for" not in query:
        print("❌ Invalid query format")
        continue

    school_name = query.split("US news for ")[1].split(" located")[0]
    school_name = school_name.replace("High School:", "").strip()

    link = find_usnews_link(query)
    if not link:
        print("❌ US News link not found")
        continue

    print("✅ Link:", link)

    try:
        overview, ratio, math, reading, science, graduation, school_type = extract_data(link, school_name)
        print("\n===== OVERVIEW (Extracted) =====\n")
        print(overview)
        print("\n===============================\n")        

        results.append([
            school_name,
            school_type,
            overview,
            ratio,
            math,
            reading
        ])

    except Exception as e:
        print("❌ Error:", e)

    time.sleep(2)  # polite delay


# -----------------------------------
# SAVE TO CSV WITH DATE-TIME
# -----------------------------------
timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
output_file = f"usnews_output_{timestamp}.csv"

with open(output_file, "w", newline="", encoding="utf-8") as f:
    writer = csv.writer(f)
    writer.writerow([
        "School Name",
        "School Type",
        "Overview",
        "Student/Teacher Ratio",
        "Math Proficiency",
        "Reading Proficiency"
    ])
    writer.writerows(results)

print(f"\n✅ CSV file created: {output_file}")
