import requests
from bs4 import BeautifulSoup
from datetime import datetime
import sys
import xlsxwriter
import os
import re
from urllib.parse import unquote

BASE_URL = "https://www.rpg.b-zone.ro"

FACTIONS_LIST = {
    "1": {"name": "Green Street Bloods", "slug": "greenstreet", "color": "#33AA33"},
    "2": {"name": "Verdant Family", "slug": "verdant", "color": "#656565"},
    "3": {"name": "Vietnamese Boys", "slug": "vietnamese", "color": "#8aa09d"},
    "4": {"name": "The Tsar Bratva", "slug": "tsarbratva", "color": "#946141"},
    "5": {"name": "Red Dragon Triad", "slug": "reddragon", "color": "#d0000f"},
    "6": {"name": "Southern Pimps", "slug": "southernpimps", "color": "#B32CF6"},
    "7": {"name": "Avispa Rifa", "slug": "avispa", "color": "#3a460c"},
    "8": {"name": "69 Pier Mobs", "slug": "69pier", "color": "#33CCFF"},
    "9": {"name": "El Loco Cartel", "slug": "elloco", "color": "#FF9900"}
}

HEADERS = { "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36" }

def check_session_validity(cookie_jar):
    test_url = f"{BASE_URL}/account/general"
    try:
        response = requests.get(test_url, headers=HEADERS, cookies=cookie_jar)
        if "/account/login" in response.url.lower(): return False
        if "Login" in response.text[:2000] and "password" in response.text[:2000]: return False
        return True
    except:
        return False

def get_my_username(cookie_jar):
    url = f"{BASE_URL}/account/general"
    try:
        response = requests.get(url, headers=HEADERS, cookies=cookie_jar)
        soup = BeautifulSoup(response.text, 'html.parser')
        profile_li = soup.find('li', id='profileName')
        if profile_li:
            link = profile_li.find('a')
            if link:
                raw_name = link.get_text(strip=True)
                clean_name = raw_name.replace("keyboard_arrow_down", "").strip()
                return clean_name
    except: pass
    return "Unknown"

def get_soup(url, cookie_jar):
    try:
        response = requests.get(url, headers=HEADERS, cookies=cookie_jar)
        if response.status_code == 200:
            return BeautifulSoup(response.text, 'html.parser')
        return None
    except:
        return None

def get_staff_members(cookie_jar):
    print("[*] Scanez listele de Admini si Lideri...")
    staff_ids = set()
    urls = [BASE_URL + "/staff/admins", BASE_URL + "/staff/leaders"]
    
    for url in urls:
        soup = get_soup(url, cookie_jar)
        if not soup: continue
        links = soup.find_all('a', href=True)
        for link in links:
            href = link['href']
            if "/players/general/" in href:
                try:
                    raw_id = href.split('/players/general/')[1]
                    clean_id = unquote(raw_id).lower().strip()
                    if clean_id:
                        staff_ids.add(clean_id)
                except: continue
    print(f"[INFO] Am incarcat lista de Admini si Lideri pentru a-i exclude din evidenta.")
    return staff_ids

def get_wars(cookie_jar, faction_slug, check_date, my_faction_name):
    target_url = f"{BASE_URL}/wars/viewall/gang/{faction_slug}"
    print(f"[*] Caut waruri pe data de {check_date}...")
    
    soup = get_soup(target_url, cookie_jar)
    if not soup: return []

    war_data_list = [] 
    all_rows = soup.find_all('tr')
    
    for row in all_rows:
        row_text = row.get_text()
        if check_date in row_text:
            links = row.find_all('a', href=True)
            cols = row.find_all('td')
            table_index = 0 
            is_lost = False 
            
            if len(cols) >= 5:
                attacker_name = cols[1].get_text(strip=True)
                defender_name = cols[4].get_text(strip=True)
                try:
                    attacker_score = int(cols[2].get_text(strip=True))
                    defender_score = int(cols[3].get_text(strip=True))
                except:
                    attacker_score = 0; defender_score = 0

                if my_faction_name.lower() in attacker_name.lower():
                    table_index = 0
                    if attacker_score < defender_score: is_lost = True
                elif my_faction_name.lower() in defender_name.lower():
                    table_index = 1 
                    if defender_score < attacker_score: is_lost = True

            for link in links:
                href = link['href'].lower()
                if '/wars/view/' in href and 'viewall' not in href:
                    full_link = BASE_URL + link['href'] if not link['href'].startswith('http') else link['href']
                    is_duplicate = False
                    for existing_link, _, _ in war_data_list:
                        if existing_link == full_link:
                            is_duplicate = True; break
                    if not is_duplicate:
                        war_data_list.append((full_link, table_index, is_lost))
                    break 
    
    return war_data_list[::-1]

def detect_status(cell):
    is_invoit = False; is_inactiv = False; is_banat = False
    spans = cell.find_all('span')
    for span in spans:
        data_title = span.get('data-original-title', '').lower()
        title = span.get('title', '').lower()
        icon_text = span.get_text(strip=True).lower()
        full_text = f"{data_title} {title} {icon_text}"
        if "absence" in full_text or icon_text == "done": is_invoit = True
        if "inactivity" in full_text or icon_text == "notifications_paused": is_inactiv = True
        if "banned" in full_text or icon_text == "lock": is_banat = True
    return is_invoit, is_inactiv, is_banat

def format_score(kills, deaths):
    diff = kills - deaths
    str_val = f"+{diff}" if diff > 0 else f"{diff}"
    if diff == 0: str_val = "0"
    return str_val, diff

def format_signed(value):
    if value > 0: return f"+{value}"
    if value == 0: return "0"
    return str(value)

def find_kd_column_index(table):
    # Looks at the header row of the stats table and returns the index of a
    # column labeled K/D (or a close variant). Returns None if no such column exists,
    # so callers can fall back to computing it from separate Kills/Deaths columns.
    rows = table.find_all('tr')
    if not rows: return None
    header_cells = rows[0].find_all(['th', 'td'])
    kd_labels = {'k/d', 'kd', 'k-d', 'scor', 'score'}
    for idx, cell in enumerate(header_cells):
        label = cell.get_text(strip=True).lower().replace(' ', '')
        if label in kd_labels:
            return idx
    return None

def parse_kd_value(text):
    # Handles a K/D cell already shown as a signed difference ("+5", "-3", "0", "5"),
    # or as a kills/deaths pair ("12/8" or "12-8"), in which case it's reduced to the difference.
    if not text: return None
    text = text.strip()
    pair_match = re.match(r'^(-?\d+)\s*[/\-]\s*(-?\d+)$', text)
    if pair_match:
        try:
            k = int(pair_match.group(1)); d = int(pair_match.group(2))
            return k - d
        except: return None
    signed_match = re.match(r'^([+-]?\d+)$', text)
    if signed_match:
        try: return int(signed_match.group(1))
        except: return None
    return None

def process_war_details(war_url, war_index, target_table_index, is_lost, members_db, cookie_jar, staff_list_ids):
    print(f"[*] Descarc datele War #{war_index}...")
    soup = get_soup(war_url, cookie_jar)
    if not soup: return

    all_tables = soup.find_all('table')
    stat_tables = []
    for table in all_tables:
        txt = table.get_text()
        if "Name" in txt and ("Seconds" in txt or "Secunde" in txt): stat_tables.append(table)
    
    if not stat_tables: return
    target_table = stat_tables[target_table_index] if target_table_index < len(stat_tables) else stat_tables[0]
    kd_col_idx = find_kd_column_index(target_table)

    rows = target_table.find_all('tr')[1:] 
    for row in rows:
        cols = row.find_all('td')
        if len(cols) < 4: continue
        name_cell = cols[0]
        
        links = name_cell.find_all('a')
        is_civilian = False
        for link in links:
            if 'c4c4c4' in link.get('style', '').lower():
                is_civilian = True; break
        if is_civilian: continue 

        if links: clean_name = "".join([l.get_text(strip=True) for l in links])
        else:
            clean_name = name_cell.get_text(strip=True)
            for garbage in ["done", "notifications_paused", "lock"]: clean_name = clean_name.replace(garbage, "")
        
        is_staff = False
        for link in links:
            href = link.get('href', '')
            if "/players/general/" in href:
                try:
                    p_id = href.split('/players/general/')[1]
                    p_id_clean = unquote(p_id).lower().strip()
                    if p_id_clean in staff_list_ids:
                        is_staff = True
                        break
                except: pass

        invoit, inactiv, banat = False, False, False
        for j in range(min(4, len(cols))):
            inv, ina, ban = detect_status(cols[j])
            if inv: invoit = True
            if ina: inactiv = True
            if ban: banat = True
        
        try: seconds = int(cols[-1].get_text(strip=True))
        except: seconds = 0
        
        score_val = 0
        score_str = "0"
        kills = 0
        deaths = 0
        kd_found = False

        if kd_col_idx is not None and kd_col_idx < len(cols):
            parsed = parse_kd_value(cols[kd_col_idx].get_text(strip=True))
            if parsed is not None:
                score_val = parsed
                score_str = format_signed(score_val)
                kd_found = True

        if not kd_found:
            try:
                if len(cols) > 5:
                    kills = int(cols[4].get_text(strip=True))
                    deaths = int(cols[5].get_text(strip=True))
                else:
                    kills = int(cols[1].get_text(strip=True))
                    deaths = int(cols[2].get_text(strip=True))
                score_str, score_val = format_score(kills, deaths)
            except: pass

        if clean_name not in members_db: members_db[clean_name] = {"wars_data": {}}
        
        members_db[clean_name]["wars_data"][war_index] = {
            "seconds": seconds, 
            "score_str": score_str,
            "score_val": score_val,
            "kills": kills,
            "deaths": deaths,
            "invoit": invoit, "inactiv": inactiv, "banat": banat, 
            "is_staff": is_staff, "war_lost": is_lost
        }

def calculate_activity_sanction(total_wars, absences):
    sanctions = []
    if total_wars == 6:
        if absences == 1: sanctions.append("Amenda $25000")
        if absences == 2: sanctions.append("Amenda $50000")
        if absences == 3: sanctions.append("AV")
        if absences == 4: sanctions.append("AV"); sanctions.append("Amenda $25000")
        if absences == 5: sanctions.append("AV"); sanctions.append("Amenda $50000")
        if absences == 6: sanctions.append("FW")
    elif total_wars == 5:
        if absences == 1: sanctions.append("Amenda $25000")
        if absences == 2: sanctions.append("Amenda $50000")
        if absences == 3: sanctions.append("AV")
        if absences == 4: sanctions.append("AV"); sanctions.append("Amenda $50000")
        if absences == 5: sanctions.append("FW")
    elif total_wars == 4:
        if absences == 1: sanctions.append("Amenda $50000")
        if absences == 2: sanctions.append("AV")
        if absences == 3: sanctions.append("AV"); sanctions.append("Amenda $50000")
        if absences == 4: sanctions.append("FW")
    elif total_wars == 3:
        if absences == 1: sanctions.append("AV")
        if absences == 2: sanctions.append("AV"); sanctions.append("Amenda $50000")
        if absences == 3: sanctions.append("FW")
    elif total_wars == 2:
        if absences == 1: sanctions.append("AV"); sanctions.append("Amenda $50000")
        if absences == 2: sanctions.append("FW")
    elif total_wars == 1:
        if absences == 1: sanctions.append("FW")
    return sanctions

def calculate_worst_sanction(score):
    if score <= -15: return "FW"
    if score <= -10: return "Amenda $30000"
    if score <= -5: return "Amenda $25000"
    return None

def parse_and_sum_sanctions(sanction_list):
    total_money = 0
    av_count = 0
    fw_count = 0
    for s in sanction_list:
        if "Amenda" in s:
            nums = re.findall(r'\d+', s.replace('.', ''))
            if nums: total_money += int(nums[0])
        elif "AV" in s: av_count += 1
        elif "FW" in s or "Faction Warn" in s: fw_count += 1
    final_parts = []
    if fw_count > 0: final_parts.append(f"{fw_count}x FW" if fw_count > 1 else "FW")
    if av_count > 0: final_parts.append(f"{av_count}x AV" if av_count > 1 else "AV")
    if total_money > 0:
        formatted_money = "{:,}".format(total_money).replace(",", ".")
        final_parts.append(f"Amenda ${formatted_money}")
    if not final_parts: return "-"
    return " + ".join(final_parts)

def save_html_report(filename, members_db, total_wars, min_seconds_req, use_worst_score, my_name, faction_name, date_str, faction_color):
    entries = []
    
    acronym_style = f"color: {faction_color};" 
    white_text = "color: #FFFFFF;" 

    for member_name, data in members_db.items():
        is_staff_global = False
        for i in range(1, total_wars + 1):
             if data["wars_data"].get(i, {}).get("is_staff"):
                 is_staff_global = True; break
        if is_staff_global: continue
        
        absences = 0
        sanctions_list = []
        worst_score_reasons = [] 
        
        for i in range(1, total_wars + 1):
            war_info = data["wars_data"].get(i, {})
            min_req = min_seconds_req.get(i, 0)
            
            is_present = False

            if war_info.get('seconds', 0) >= min_req:
                is_present = True
            
            if not is_present:
                if not (war_info.get('invoit') or war_info.get('inactiv') or war_info.get('banat')):
                    absences += 1
            
            if use_worst_score and war_info.get('war_lost'):
                 if not (war_info.get('invoit') or war_info.get('inactiv') or war_info.get('banat')):
                    w_sanc = calculate_worst_sanction(war_info.get('score_val', 0))
                    if w_sanc: 
                        sanctions_list.append(w_sanc)
                        worst_score_reasons.append(f"Scor {war_info['score_val']} (War {i})")

        if absences > 0:
            act_sancs = calculate_activity_sanction(total_wars, absences)
            sanctions_list.extend(act_sancs)
        
        if sanctions_list:
            raw_sanc_str = parse_and_sum_sanctions(sanctions_list)
            
            colored_sanc = raw_sanc_str.replace("FW", '<span style="color: #FF0000;">FW</span>')
            colored_sanc = colored_sanc.replace("Faction Warn", '<span style="color: #FF0000;">Faction Warn</span>')
            colored_sanc = colored_sanc.replace("AV", '<span style="color: #FFA500;">AV</span>')
            colored_sanc = re.sub(r'(Amenda \$[\d\.]+)', r'<span style="color: #FFFF00;">\1</span>', colored_sanc)

            reasons = []
            if absences > 0: reasons.append(f"Absent la {absences}/{total_wars} Wars")
            if worst_score_reasons: reasons.extend(worst_score_reasons)
            reason_str = ", ".join(reasons)
            
            entry_html = f"""
            <div style="font-family: Arial, sans-serif; font-size: 14px; margin-bottom: 0px; line-height: 1.5;">
                <b>
                <span style="{acronym_style}">N</span><span style="{white_text}">ume: {my_name}</span><br>
                <span style="{acronym_style}">J</span><span style="{white_text}">ucator: {member_name}</span><br>
                <span style="{acronym_style}">M</span><span style="{white_text}">otiv: {reason_str}</span><br>
                <span style="{acronym_style}">S</span><span style="{white_text}">anctionat cu: {colored_sanc}</span><br>
                <span style="{acronym_style}">A</span><span style="{white_text}">lte precizari: Evidenta Wars {date_str}</span>
                </b>
            </div>
            <br>
            """
            entries.append(entry_html)

    with open(filename, "w", encoding="utf-8") as f:
        f.write(f"""<!DOCTYPE html>
        <html>
        <head>
            <meta charset="UTF-8">
            <title>Sanctiuni {faction_name}</title>
        </head>
        <body style="background-color: #262626; color: #ffffff; font-family: Arial, sans-serif; padding: 20px;">
        <br>
        """)
        
        if not entries:
            f.write('<p style="color: #00FF00; font-weight: bold;">Nicio sanctiune gasita.</p>')
        else:
            for entry in entries:
                f.write(entry)
                
        f.write("</body></html>")
        
    print(f"[GATA] Fisier HTML sanctiuni generat: {filename}")

def save_styled_excel(filename, members_db, total_wars, min_seconds_req, use_worst_score, faction_color):
    try:
        workbook = xlsxwriter.Workbook(filename)
        ws = workbook.add_worksheet('Evidenta')
        
        header_fmt = workbook.add_format({
            'bold': True, 'align': 'center', 'valign': 'vcenter',
            'fg_color': faction_color, 'font_color': '#FFFFFF', 
            'border': 1, 'font_size': 11, 'font_name': 'Calibri'
        })
        
        subheader_fmt = workbook.add_format({
            'bold': True, 'align': 'center', 'valign': 'vcenter',
            'fg_color': '#37474F', 'font_color': '#FFFFFF', 
            'border': 1, 'font_size': 10, 'font_name': 'Calibri'
        })
        
        cell_fmt = workbook.add_format({
            'align': 'center', 'valign': 'vcenter', 'border': 1, 'font_name': 'Calibri', 'font_size': 10
        })
        
        name_fmt = workbook.add_format({
            'bold': True, 'align': 'center', 'valign': 'vcenter', 
            'border': 1, 'bg_color': '#F2F2F2', 'font_name': 'Calibri', 'font_size': 11
        })
        
        bold_fmt = workbook.add_format({
            'bold': True, 'align': 'center', 'valign': 'vcenter', 
            'border': 1, 'font_name': 'Calibri', 'font_size': 10
        })

        score_pos_fmt = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'font_color': '#2E7D32', 'bold': True, 'font_name': 'Calibri'})
        score_zero_fmt = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'font_color': '#000000', 'bold': True, 'font_name': 'Calibri'})
        score_warn_fmt = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'font_color': '#EF6C00', 'bold': True, 'font_name': 'Calibri'})
        score_err1_fmt = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'font_color': '#C62828', 'bold': True, 'font_name': 'Calibri'})
        score_err2_fmt = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'font_color': '#B71C1C', 'bold': True, 'font_name': 'Calibri'})
        score_crit_fmt = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'font_color': '#4A148C', 'bold': True, 'font_name': 'Calibri'})

        absent_fmt = workbook.add_format({
            'bold': True, 'bg_color': '#FFEBEE', 'font_color': '#C62828', 
            'align': 'center', 'valign': 'vcenter', 'border': 1, 'font_name': 'Calibri'
        })
        special_fmt = workbook.add_format({
            'bold': True, 'bg_color': '#ECEFF1', 'font_color': '#455A64', 
            'align': 'center', 'valign': 'vcenter', 'border': 1, 'font_name': 'Calibri'
        })
        sec_ok_fmt = workbook.add_format({
            'bold': True, 'bg_color': '#E8F5E9', 'font_color': '#1B5E20', 
            'align': 'center', 'valign': 'vcenter', 'border': 1, 'font_name': 'Calibri'
        })
        sec_bad_fmt = workbook.add_format({
            'bold': True, 'bg_color': '#FFEBEE', 'font_color': '#B71C1C', 
            'align': 'center', 'valign': 'vcenter', 'border': 1, 'font_name': 'Calibri'
        })
        
        clean_fmt = workbook.add_format({'bold': True, 'bg_color': '#E8F5E9', 'font_color': '#1B5E20', 'align': 'center', 'valign': 'vcenter', 'border': 1, 'font_name': 'Calibri'})
        fw_fmt = workbook.add_format({'bold': True, 'bg_color': '#FFCDD2', 'font_color': '#B71C1C', 'align': 'center', 'valign': 'vcenter', 'border': 1, 'font_name': 'Calibri'})
        av_fmt = workbook.add_format({'bold': True, 'bg_color': '#FFE0B2', 'font_color': '#E65100', 'align': 'center', 'valign': 'vcenter', 'border': 1, 'font_name': 'Calibri'})
        av_money_fmt = workbook.add_format({'bold': True, 'bg_color': '#FFCCBC', 'font_color': '#BF360C', 'align': 'center', 'valign': 'vcenter', 'border': 1, 'font_name': 'Calibri'})
        amenda_fmt = workbook.add_format({'bold': True, 'bg_color': '#FFF9C4', 'font_color': '#F57F17', 'align': 'center', 'valign': 'vcenter', 'border': 1, 'font_name': 'Calibri'})

        def get_sanction_style(text_sanc):
            if text_sanc == "-" or text_sanc in ["INVOIT", "INACTIV", "BANAT", ""]: return clean_fmt
            if "FW" in text_sanc: return fw_fmt
            if "AV" in text_sanc and "Amenda" in text_sanc: return av_money_fmt
            if "AV" in text_sanc: return av_fmt
            if "Amenda" in text_sanc: return amenda_fmt
            return cell_fmt

        ws.merge_range('A1:A2', 'Nume', header_fmt)
        
        col_idx = 1
        for i in range(1, total_wars + 1):
            sec_req = min_seconds_req.get(i, 0)
            ws.merge_range(0, col_idx, 0, col_idx+1, f'War {i} ({sec_req}s)', header_fmt)
            ws.write(1, col_idx, 'Secunde', subheader_fmt)
            ws.write(1, col_idx+1, 'K/D', subheader_fmt)
            col_idx += 2
        
        start_result_col = col_idx
        total_result_cols = 2 if not use_worst_score else 3
        
        ws.merge_range(0, start_result_col, 0, start_result_col + total_result_cols - 1, 'Rezultat', header_fmt)
        ws.write(1, start_result_col, 'Abs', subheader_fmt)
        ws.write(1, start_result_col + 1, 'Sanctiuni (secunde)', subheader_fmt)
        if use_worst_score:
            ws.write(1, start_result_col + 2, 'Sanctiuni (scor)', subheader_fmt)

        ws.set_column(0, 0, 37)
        c = 1
        for _ in range(total_wars):
            ws.set_column(c, c, 10); ws.set_column(c+1, c+1, 8); c += 2
        ws.set_column(c, c, 5)
        ws.set_column(c+1, c+1, 25)
        if use_worst_score: ws.set_column(c+2, c+2, 25)

        row_idx = 2
        for member_name, data in members_db.items():
            is_staff_global = False
            for i in range(1, total_wars + 1):
                if data["wars_data"].get(i, {}).get("is_staff"): is_staff_global = True; break
            if is_staff_global: continue

            ws.write(row_idx, 0, member_name, name_fmt)
            
            absences = 0
            act_sanctions_list = []
            score_sanctions_list = []
            
            has_invoire_any = False; has_inactiv_any = False; has_ban_any = False

            curr = 1
            for i in range(1, total_wars + 1):
                war_info = data["wars_data"].get(i, {
                    "seconds": 0, "score_str": "ABSENT", "score_val": 0, "kills":0, "deaths":0, "invoit": False, "inactiv": False, "banat": False, "is_staff": False, "war_lost": False
                })
                if war_info['invoit']: has_invoire_any = True
                if war_info['inactiv']: has_inactiv_any = True
                if war_info['banat']: has_ban_any = True

                sec_display = war_info['seconds']
                sec_style = cell_fmt

                min_req = min_seconds_req.get(i, 0)
                is_present = False

                if war_info['seconds'] >= min_req:
                    is_present = True

                if war_info['invoit']: sec_display = "INVOIT"; sec_style = special_fmt
                elif war_info['inactiv']: sec_display = "INACTIV"; sec_style = special_fmt
                elif war_info['banat']: sec_display = "BANAT"; sec_style = special_fmt
                elif war_info['seconds'] < 30 and not is_present: 
                    sec_display = "ABSENT"; sec_style = absent_fmt
                else:
                    if is_present: sec_style = sec_ok_fmt 
                    else: sec_style = sec_bad_fmt

                ws.write(row_idx, curr, sec_display, sec_style)

                sc_val = war_info['score_val']
                sc_str = war_info['score_str']
                sc_style = cell_fmt 
                
                if sc_val > 0: sc_style = score_pos_fmt
                elif sc_val == 0: sc_style = score_zero_fmt
                else:
                    if sc_val > -5: sc_style = score_warn_fmt      
                    elif sc_val >= -9: sc_style = score_err1_fmt   
                    elif sc_val >= -14: sc_style = score_err2_fmt  
                    else: sc_style = score_crit_fmt                

                ws.write(row_idx, curr+1, sc_str, sc_style)
                curr += 2

                if not is_present:
                    if not (war_info['invoit'] or war_info['inactiv'] or war_info['banat']): absences += 1
                
                if use_worst_score and war_info['war_lost']:
                    if not (war_info['invoit'] or war_info['inactiv'] or war_info['banat']):
                        w_sanc = calculate_worst_sanction(war_info['score_val'])
                        if w_sanc: score_sanctions_list.append(w_sanc)
            
            ws.write(row_idx, curr, absences, bold_fmt)
            
            act_str = "-"
            if absences > 0:
                act_sanctions_list = calculate_activity_sanction(total_wars, absences)
                act_str = parse_and_sum_sanctions(act_sanctions_list)
            else:
                if has_ban_any: act_str = "BANAT"
                elif has_inactiv_any: act_str = "INACTIV"
                elif has_invoire_any: act_str = "INVOIT"
            ws.write(row_idx, curr+1, act_str, get_sanction_style(act_str))

            if use_worst_score:
                score_str = "-"
                if score_sanctions_list: score_str = parse_and_sum_sanctions(score_sanctions_list)
                ws.write(row_idx, curr+2, score_str, get_sanction_style(score_str))
            row_idx += 1

        workbook.close()
        print(f"[GATA] Fisier EXCEL generat: {filename}")
        
    except Exception as e:
        print(f"[EROARE] Nu am putut salva Excel-ul: {e}")

def main():
    print("Alege mafia pentru care faci evidenta:")
    for k, v in FACTIONS_LIST.items():
        print(f"{k}. {v['name']}")
    
    selected_idx = input("Numar mafie (1-9): ").strip()
    if selected_idx not in FACTIONS_LIST: selected_idx = "6"
    selected_faction = FACTIONS_LIST[selected_idx]
    
    use_worst = input("Calculez sanctiunile pe scor (-5/-10/-15)? (da/nu): ").strip().lower()
    use_worst_score = (use_worst == 'da' or use_worst == 'y')

    cookie_jar = None
    while True:
        print("\n[LOGIN] Introdu cookie-ul 'bzonerpg' din browser (F12 -> Application -> Cookies):")
        manual_cookie = input("Cookie: ").strip()
        if not manual_cookie: continue
        
        cj = requests.cookies.RequestsCookieJar()
        cj.set('bzonerpg', manual_cookie, domain='rpg.b-zone.ro', path='/')
        cj.set('PHPSESSID', manual_cookie, domain='rpg.b-zone.ro', path='/')
        
        if check_session_validity(cj):
            print("[SUCCES] Logat cu succes!")
            cookie_jar = cj
            break
        else:
            print("[EROARE] Cookie invalid sau expirat.")

    my_username = get_my_username(cookie_jar)
    print(f"[INFO] Bine ai venit, {my_username}!")

    staff_list_ids = get_staff_members(cookie_jar)
    
    azi_str = datetime.now().strftime("%d.%m.%Y")
    data_input = input(f"\nData evidentei [Enter pt azi ({azi_str})]: ").strip()
    target_date = data_input if data_input else azi_str
    
    war_data = get_wars(cookie_jar, selected_faction['slug'], target_date, selected_faction['name'])
    
    if not war_data: 
        print(f"[INFO] 0 waruri gasite pentru {selected_faction['name']} pe {target_date}.")
        input("Apasa ENTER..."); sys.exit()
    
    total_wars = len(war_data)
    print(f"\n[OK] {total_wars} waruri identificate.")
    
    min_seconds_req = {}
    for i in range(1, total_wars + 1):
        is_lost = war_data[i-1][2]
        status_war = "PIERDUT" if is_lost else "CASTIGAT"
        while True:
            try:
                sec = int(input(f"Secunde minime War {i} ({status_war}): "))
                min_seconds_req[i] = sec
                break
            except: pass

    members_db = {}
    for i, (link, table_idx, is_lost) in enumerate(war_data, 1):
        process_war_details(link, i, table_idx, is_lost, members_db, cookie_jar, staff_list_ids)

    safe_faction = selected_faction['name'].replace(" ", "_")
    safe_date = target_date.replace('.', '_')
    
    file_excel = f"Evidenta_{safe_faction}_{safe_date}.xlsx"
    save_styled_excel(file_excel, members_db, total_wars, min_seconds_req, use_worst_score, selected_faction['color'])
    
    file_html = f"Sanctiuni_{safe_faction}_{safe_date}.html"
    save_html_report(file_html, members_db, total_wars, min_seconds_req, use_worst_score, my_username, selected_faction['name'], target_date, selected_faction['color'])

if __name__ == "__main__":
    try: main()
    except Exception as e: print(f"\n[CRASH] {e}")
    input("\nApasa ENTER pentru a iesi...")