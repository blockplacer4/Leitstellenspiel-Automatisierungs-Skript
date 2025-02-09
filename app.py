import logging
from time import sleep
from difflib import get_close_matches
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.alert import Alert
import re
import os
import dotenv

dotenv.load_dotenv()

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(message)s')

VEHICLE_MAPPINGS = {
    "Dekon-P": "Dekon-P",
    "Drehleiter": "DLK",
    "ELW 1": "ELW",
    "ELW 2": "ELW 2",
    "Flughafen-Löschfahrzeug": "Flgh-FLF",
    "Rettungstreppen": "Flgh-RTF",
    "Feuerwehrkran": "FwK",
    "Feuerwehrkräne (FwK)": "FwK",
    "GW-A": "GW-A",
    "GW-Atemschutz": "GW-A",
    "Gerätewagen-Atemschutz": "GW-A",
    "GW-Gefahrgut": "GW-G",
    "GW-Höhenrettung": "GW-H",
    "GW-Messtechnik": "GW-M",
    "GW-Öl": "GW-ÖL",
    "GW-S": "GW-S",
    "GW-Mess": "GW-M",
    "GW-L2": "GW-L2",
    "SW 1000": "GW-L2",
    "GW-L2-Wasser": "GW-L2",
    "Löschfahrzeug": "LF",
    "MTW": "MTW",
    "Rüstwagen": "RW",
    "TLF": "TLF",
    "WF-GW": "WF-GW",
    "WF-TL": "WF-TL",
    "WF-TM": "WF-TM",
    "WF-ULF": "WF-ULF",
    "First Responder": "First Responder",
    "G-RTW": "G-RTW",
    "KdoW LNA": "KdoW LNA",
    "KdoW OrgL": "KdoW OrgL",
    "KTW": "KTW",
    "NAW": "NAW",
    "NEF": "NEF",
    "RTH": "RTH",
    "RTW": "RTW",
    "FüKw": "FüKw",
    "FuStW": "FuStW",
    "GefKw": "GefKw",
    "GruKw": "GruKw",
    "IeBeKw": "IeBeKw",
    "MEK-MTF": "MEK-MTF",
    "MEK-ZF": "MEK-ZF",
    "Pol-Hub": "Pol-Hub",
    "SEK-MTF": "SEK-MTF",
    "SEK-ZF": "SEK-ZF",
    "Wasserwerfer": "WaWe",
    "ELW 1 (SEG)": "ELW 1 (SEG)",
    "GW-San": "GW-San",
    "KTW Typ B": "KTW Typ B",
    "MANV 10": "MANV 10",
    "MANV 5": "MANV 5",
    "WR-GW": "WR-GW",
    "WR-GW-Taucher": "WR-GW-Taucher",
    "WR-MZB": "WR-MZB",
    "AH DLE": "AH DLE",
    "AH-MzAB": "AH-MzAB",
    "AH-MzB": "AH-MzB",
    "AH-SchlB": "AH-SchlB",
    "BRmG R": "BRmG R",
    "GKW": "GKW",
    "LKW K 9": "LKW K 9",
    "LKW Lgr 19tm": "LKW Lgr 19tm",
    "MLW 5": "MLW 5",
    "MTW-TZ": "MTW-TZ",
    "Mehrzweckkraftwagen": "MzKW",
    "MzGW (FGr N)": "MzKW",
    "Tankwagen": "TKW",
    "NEA50": "NEA50",
    "LNA": "Kdow LNA"
}

PARTIAL_MATCHES = {
    "atemschutz": "GW-A",
    "gefahrgut": "GW-G",
    "höhenrettung": "GW-H",
    "messtechnik": "GW-M",
    "öl": "GW-ÖL",
    "sanität": "GW-S",
    "drehleiter": "DLK",
    "rüst": "RW",
    "tlf": "LF",
    "lf": "LF",
    "dlk": "DLK",
    "hlf": "RW",
    "hlf 20": "RW",
    "rw": "RW",
    "streifenwagen": "FuStW",
    "GW-Mess": "GW-M",
}

def smart_vehicle_match(vehicle_name):
    if vehicle_name in VEHICLE_MAPPINGS:
        return VEHICLE_MAPPINGS[vehicle_name]
    vehicle_lower = vehicle_name.lower()
    for partial, abbrev in PARTIAL_MATCHES.items():
        if partial in vehicle_lower:
            return abbrev
    if "-" in vehicle_name:
        abbrev_parts = vehicle_name.split("-")
        possible_matches = [v for v in VEHICLE_MAPPINGS.values() if v.startswith(abbrev_parts[0])]
        if possible_matches:
            return possible_matches[0]
    possible_matches = get_close_matches(vehicle_name, VEHICLE_MAPPINGS.keys(), n=1, cutoff=0.6)
    if possible_matches:
        return VEHICLE_MAPPINGS[possible_matches[0]]
    return vehicle_name

def extract_current_vehicles(driver):
    current_vehicles = {}
    enroute_personnel = 0
    sleep(0.125)
    try:
        driving_table = driver.find_element(By.ID, 'mission_vehicle_driving')
        rows = driving_table.find_elements(By.TAG_NAME, 'tr')
        for row in rows[1:]:
            cells = row.find_elements(By.TAG_NAME, 'td')
            if len(cells) >= 3:
                vehicle_cell = cells[1].text
                personnel_text = cells[2].text.strip()
                if '(' in vehicle_cell and ')' in vehicle_cell:
                    raw_type = vehicle_cell.split('(')[1].split(')')[0].strip()
                    matched_type = smart_vehicle_match(raw_type)
                    current_vehicles[matched_type] = current_vehicles.get(matched_type, 0) + 1
                if personnel_text.isdigit():
                    enroute_personnel += int(personnel_text)
    except Exception as e:
        logging.error(f"Error reading driving vehicles: {str(e)}")
    sleep(0.25)
    try:
        at_mission_table = driver.find_element(By.ID, 'mission_vehicle_at_mission')
        rows = at_mission_table.find_elements(By.TAG_NAME, 'tr')
        for row in rows[1:]:
            cells = row.find_elements(By.TAG_NAME, 'td')
            if len(cells) >= 2:
                vehicle_cell = cells[1].text
                if '(' in vehicle_cell and ')' in vehicle_cell:
                    raw_type = vehicle_cell.split('(')[1].split(')')[0].strip()
                    matched_type = smart_vehicle_match(raw_type)
                    current_vehicles[matched_type] = current_vehicles.get(matched_type, 0) + 1
    except Exception as e:
        logging.error(f"Error reading vehicles at mission: {str(e)}")
    sleep(0.25)
    return current_vehicles, enroute_personnel

def extract_vehicle_requirements(table):
    requirements = {}
    rows = table.find_elements(By.TAG_NAME, 'tr')
    for row in rows:
        try:
            cells = row.find_elements(By.TAG_NAME, 'td')
            if len(cells) == 2:
                vehicle_text = cells[0].text.strip()
                value_text = cells[1].text.strip()
                if "anforderungswahrscheinlichkeit" in vehicle_text.lower():
                    continue
                if "feuerwehrleute" in vehicle_text.lower() or "feuerwehrmann" in vehicle_text.lower():
                    total_firefighters = int(value_text)
                    lf_count = (total_firefighters + 2) // 3
                    requirements["LF"] = requirements.get("LF", 0) + lf_count
                elif vehicle_text.startswith('Benötigte '):
                    base_vehicle = vehicle_text.replace('Benötigte ', '').strip()
                    count = int(value_text)
                    if 'schlauchwagen' in base_vehicle.lower():
                        matched_vehicle = 'GW-L2'
                    else:
                        matched_vehicle = smart_vehicle_match(base_vehicle)
                    requirements[matched_vehicle] = count
                else:
                    matched_vehicle = smart_vehicle_match(vehicle_text)
                    try:
                        count = int(value_text)
                        if matched_vehicle in VEHICLE_MAPPINGS.values():
                            requirements[matched_vehicle] = max(requirements.get(matched_vehicle, 0), count)
                    except ValueError:
                        continue
        except Exception as e:
            logging.error(f"Error extracting requirements: {str(e)}")
    return requirements

def extract_actual_patients(driver):
    wait = WebDriverWait(driver, 10)
    try:
        container = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#col_left, .col-lg-6")))
        patient_divs = container.find_elements(By.CSS_SELECTOR, ".mission_patient")
    except Exception as e:
        print("Exception in extract_actual_patients:", e)
        patient_divs = driver.find_elements(By.CSS_SELECTOR, ".mission_patient")
    actual_count = len(patient_divs)
    nef_needed = 0
    for div in patient_divs:
        try:
            alert_div = div.find_element(By.CSS_SELECTOR, ".alert-danger")
            if "NEF" in alert_div.text:
                nef_needed += 1
        except:
            pass
    logging.info(f"Actual patients: {actual_count} - NEF needed: {nef_needed}")
    return actual_count, nef_needed

def extract_patient_requirements(col_md_4_divs):
    min_patients = 0
    nef_probability = 0
    if len(col_md_4_divs) >= 3:
        try:
            info_div = col_md_4_divs[2]
            info_table = info_div.find_element(By.TAG_NAME, 'table')
            rows = info_table.find_elements(By.TAG_NAME, 'tr')
            for row in rows:
                cells = row.find_elements(By.TAG_NAME, 'td')
                if len(cells) == 2:
                    label = cells[0].text.strip().lower()
                    value = cells[1].text.strip()
                    if "mindest patientenanzahl" in label:
                        min_patients = int(value)
                    elif "nef anforderungswahrscheinlichkeit" in label:
                        nef_probability = int(value)
        except Exception as e:
            logging.error(f"Error extracting patient requirements: {str(e)}")
    return min_patients, nef_probability

def extract_missing_personnel(driver):
    try:
        alert_div = driver.find_element(By.CLASS_NAME, 'alert-missing-vehicles')
        personnel_div = alert_div.find_element(By.CSS_SELECTOR, 'div[data-requirement-type="personnel"]')
        text = personnel_div.text.strip()
        match = re.search(r'(\d+)\s+Feuerwehrleute', text)
        if match:
            number = int(match.group(1))
            return number
    except:
        return 0

def extract_missing_water(driver):
    try:
        missing_div = driver.find_element(By.CSS_SELECTOR, ".progress-bar-missing.progress-bar-mission-window-water")
        text = missing_div.text  # z.B. "Fehlen: 3.500 l."
        logging.info(f"Found missing water div text: '{text}'")
        match = re.search(r'Fehlen:\s*([\-\d\.]+)\s*l\.', text)
        if match:
            missing_value = float(match.group(1).replace('.', ''))
            logging.info(f"Parsed missing water: {missing_value}")
            return missing_value
    except Exception as e:
        logging.warning(f"No missing water div found or there is an error")
        # logging.error(f"Error extracting missing water: {str(e)}")
    return 0  # Kein Wasser benötigt lol

def handle_patients_and_nef(driver, required_vehicles, current_vehicles, enroute_personnel, min_patients, nef_probability):
    if enroute_personnel is None:
        enroute_personnel = 0
    actual_count, nef_in_divs = extract_actual_patients(driver)
    final_patient_count = max(min_patients, actual_count)
    current_nef = current_vehicles.get("NEF", 0) + current_vehicles.get("NAW", 0)
    current_rtw = current_vehicles.get("RTW", 0)
    needed_rtw = max(0, final_patient_count - current_rtw)
    needed_nef_total = max(0, nef_in_divs - current_nef)
    needed_nef = 1 if needed_nef_total > 0 else 0

    if needed_rtw > 0 and needed_nef > 0:
        current_naw = current_vehicles.get("NAW", 0)
        if current_naw < 1:
            required_vehicles["NAW"] = required_vehicles.get("NAW", 0) + 1
            logging.info("Alarming one NAW for RTW + NEF requirements (only 1 NEF needed per pass).")
            needed_rtw = 0
            needed_nef = 0
        else:
            required_vehicles["RTW"] = required_vehicles.get("RTW", 0) + needed_rtw
            required_vehicles["NEF"] = required_vehicles.get("NEF", 0) + needed_nef
            logging.info(f"Alarming {needed_rtw} RTW + {needed_nef} NEF (only 1 NEF needed per pass).")
            needed_rtw = 0
            needed_nef = 0
    elif needed_rtw > 0:
        required_vehicles["RTW"] = required_vehicles.get("RTW", 0) + needed_rtw
    elif needed_nef > 0:
        required_vehicles["NEF"] = required_vehicles.get("NEF", 0) + needed_nef

    if needed_nef > 0 and current_nef > 0:
        logging.info("NEF requirement already met by NEF/NAW on route.")

    lna_needed = False
    try:
        patient_divs = driver.find_elements(By.CSS_SELECTOR, ".mission_patient")
        for patient in patient_divs:
            try:
                alert_div = patient.find_element(By.CSS_SELECTOR, ".alert.alert-danger")
                if "LNA" in alert_div.text:
                    lna_needed = True
                    break
            except:
                continue
    except Exception as e:
        logging.error(f"Error checking LNA requirement: {str(e)}")
    if lna_needed:
        if current_vehicles.get("KdoW LNA", 0) < 1:
            logging.info("LNA requirement found. Dispatching 1 KdoW LNA.")
            required_vehicles["KdoW LNA"] = 1

    try:
        missing_personnel = extract_missing_personnel(driver)
        if missing_personnel is None:
            missing_personnel = 0
        missing_personnel = max(0, missing_personnel - enroute_personnel)
    except Exception as e:
        logging.error(f"Error extracting missing personnel: {str(e)}")
        missing_personnel = 0
    if missing_personnel > 0:
        lf_needed = (missing_personnel + 2) // 3
        required_vehicles["LF"] = required_vehicles.get("LF", 0) + lf_needed
        logging.info(f"Missing personnel: {missing_personnel}, dispatching {lf_needed} LF")

    try:
        alert = driver.find_element(By.CSS_SELECTOR, ".alert-missing-vehicles")
        alert_text = alert.text.strip()
        training_match = re.search(r'(\d+)\s+Person(?:en)?\s+mit\s+([\w\s\-\(\)]+)-Ausbildung', alert_text, re.IGNORECASE)
        if training_match:
            count_training = int(training_match.group(1))
            training_type = training_match.group(2).strip()
            mapped_type = smart_vehicle_match(training_type)
            current_training = current_vehicles.get(mapped_type, 0)
            if current_training < count_training:
                required_vehicles[mapped_type] = max(required_vehicles.get(mapped_type, 0), count_training)
                logging.info(f"Dispatching {mapped_type} for missing personnel with training: {training_type}")
    except Exception as e:
        logging.error(f"Error checking personnel training requirement: {str(e)}")

    return required_vehicles

def calculate_missing_vehicles(required_vehicles, current_vehicles):
    missing_vehicles = {}
    for vehicle_type, required_count in required_vehicles.items():
        current_count = current_vehicles.get(vehicle_type, 0)
        if required_count > current_count:
            needed = required_count - current_count
            missing_vehicles[vehicle_type] = needed
    return missing_vehicles

def set_mission_speed(driver, desired_speed):
    current_speed = None
    try:
        pause_element = driver.find_element(By.ID, 'mission_speed_pause')
        if pause_element.value_of_css_property("display") == "flex":
            current_speed = "pause"
        else:
            current_speed = "other"
    except:
        current_speed = "other"
    if desired_speed == "pause" and current_speed != "pause":
        driver.execute_script("window.open('/missionSpeed?speed=6','_blank')")
        driver.switch_to.window(driver.window_handles[-1])
        sleep(0.125)
        driver.close()
        driver.switch_to.window(driver.window_handles[0])
    elif desired_speed == "7" and current_speed == "pause":
        driver.execute_script("window.open('/missionSpeed?speed=7','_blank')")
        driver.switch_to.window(driver.window_handles[-1])
        sleep(0.125)
        driver.close()
        driver.switch_to.window(driver.window_handles[0])

def select_vehicles(driver, required_vehicles, alarm_after_selection=True):
    selected_any = False
    for vehicle_type, count in required_vehicles.items():
        logging.info(f"Preparing to select {count} vehicle(s) of type '{vehicle_type}'")
        vehicle_links = driver.find_elements(By.CLASS_NAME, 'aao_searchable')
        matched_link = None
        sleep(0.037)
        for link in vehicle_links:
            if link.get_attribute('search_attribute') == vehicle_type:
                matched_link = link
                break
        if not matched_link:
            search_attributes = [link.get_attribute('search_attribute') for link in vehicle_links]
            closest_match = get_close_matches(vehicle_type, search_attributes, n=1, cutoff=0.5)
            if closest_match:
                for link in vehicle_links:
                    if link.get_attribute('search_attribute') == closest_match[0]:
                        matched_link = link
                        break
        sleep(0.0625)
        if matched_link:
            for _ in range(count):
                try:
                    availability_span = matched_link.find_element(By.XPATH, ".//span[starts-with(@id,'available_aao_')]")
                    if "label-success" in availability_span.get_attribute("class"):
                        matched_link.click()
                        selected_any = True
                        logging.info(f"Selected vehicle: {vehicle_type}")
                    else:
                        logging.info(f"No vehicles available for {vehicle_type}. Skipping.")
                        break
                except Exception as e:
                    logging.warning(f"No availability info for {vehicle_type}: {e}")
                sleep(0.0625)
        else:
            logging.warning(f"Could not find vehicle: {vehicle_type}")

    if alarm_after_selection:
        sleep(0.125)
        alarm_button = driver.find_element(By.ID, 'mission_alarm_btn')
        alarm_button.click()
        logging.info("Alarm button clicked")
        sleep(6)
    return selected_any

def handle_water_and_dispatch(driver, missing_vehicles):
    logging.info("Dispatching all required vehicles first, then recalculate missing water.")
    if missing_vehicles:
        select_vehicles(driver, missing_vehicles, alarm_after_selection=False)
    else:
        select_vehicles(driver, {}, alarm_after_selection=False)

    missing_water = extract_missing_water(driver)
    logging.info(f"Missing water after dispatching main vehicles: {missing_water}")
    iteration = 0
    while missing_water > 0:
        iteration += 1
        logging.info(f"Iteration {iteration} - Missing water: {missing_water} - dispatching LF")
        selected = select_vehicles(driver, {"LF": 1}, alarm_after_selection=False)
        if not selected:
            logging.warning("No more LFs available. Closing this mission tab and returning.")
            driver.close()
            if driver.window_handles:
                driver.switch_to.window(driver.window_handles[0])
            return True
        sleep(0.2)
        missing_water = extract_missing_water(driver)
        logging.info(f"Missing water after LF dispatch: {missing_water}")

    logging.info("No further water needed, alarming now.")
    select_vehicles(driver, {}, alarm_after_selection=True)
    return False

def check_for_sprechwunsch(driver, wait):
    try:
        sprechwunsch_divs = driver.find_elements(By.CSS_SELECTOR, ".alert.alert-danger")
        for sw_div in sprechwunsch_divs:
            txt = sw_div.text.lower()
            if "sprechwunsch" in txt:
                link = sw_div.find_element(By.TAG_NAME, 'a').get_attribute("href")
                driver.execute_script(f"window.open('{link}','_blank')")
                driver.switch_to.window(driver.window_handles[-1])
                sleep(0.125)
                try:
                    wait.until(EC.presence_of_element_located((By.ID, 'own-hospitals')))
                    hospitals_table = driver.find_element(By.ID, 'own-hospitals')
                    rows = hospitals_table.find_elements(By.TAG_NAME, 'tr')
                    if len(rows) > 1:
                        first_row = rows[1]
                        approach_link = first_row.find_element(By.CSS_SELECTOR, "a.btn.btn-success").get_attribute("href")
                        driver.execute_script(f"window.open('{approach_link}','_blank')")
                        driver.switch_to.window(driver.window_handles[-1])
                        sleep(0.125)
                        driver.close()
                        driver.switch_to.window(driver.window_handles[-1])
                except Exception as e:
                    logging.error(f"Error handling sprechwunsch/hospitals: {str(e)}")
                driver.close()
                driver.switch_to.window(driver.window_handles[-1])
                sleep(0.125)
                break
    except Exception as e:
        logging.error(f"Error checking sprechwunsch: {str(e)}")

def check_mission_completed(driver):
    try:
        success_div = driver.find_element(By.CLASS_NAME, 'mission-success')
        success_image = driver.find_element(By.CLASS_NAME, 'mission-success-image')
        checkmark = success_image.find_element(By.XPATH, ".//img[@alt='Checkmark_mission_complete']")
        if success_div and success_image and checkmark:
            logging.info("Mission already completed, skipping...")
            return True
    except:
        pass
    try:
        success_element = driver.find_element(By.CLASS_NAME, 'alert-success')
        if "Einsatz abgeschlossen" in success_element.text:
            logging.info("Mission already completed (alert), skipping...")
            return True
    except:
        pass
    return False

def check_and_click_easter_egg(driver):
    try:
        heart_link = driver.find_element(By.ID, "easter-egg-link")
        heart_link.click()
        sleep(1)
        logging.info("Easter egg found and clicked, waiting 1 second.")
    except Exception as e:
        logging.info(f"No Easter egg found or unable to click")

def extract_additional_requirements(driver):
    additional = {}
    try:
        mission_info = driver.find_element(By.ID, 'mission_info')
        info_text = mission_info.text
        import re
        match = re.search(r'\d+\s+Person(?:en)?\s+mit\s+([\w\-\(\) ]+)-Ausbildung', info_text, re.IGNORECASE)
        if match:
            vehicle_raw = match.group(1).strip()
            matched_vehicle = smart_vehicle_match(vehicle_raw)
            additional[matched_vehicle] = 1  # Unabhängig der Zahl wird nur 1 alarmiert
            logging.info(f"Additional requirement extracted: 1 x {matched_vehicle} (aus '{vehicle_raw}-Ausbildung')")
    except Exception as e:
        logging.info("No additional requirements found on mission page.")
    return additional

def select_prisoner_vehicle(driver):
    def open_link_in_new_tab(link):
        href = link.get_attribute("href")
        if not href:
            logging.warning("No href found on prisoner link.")
            return False
        try:
            # Open the link in a new tab.
            driver.execute_script("window.open(arguments[0], '_blank');", href)
            driver.switch_to.window(driver.window_handles[-1])
            # Wait until the new tab is fully loaded.
            WebDriverWait(driver, 10).until(
                lambda d: d.execute_script("return document.readyState") == "complete"
            )
            sleep(0.125)
            driver.close()
            driver.switch_to.window(driver.window_handles[0])
            logging.info(f"Prisoner link opened and closed: {href}")
            return True
        except Exception as e:
            logging.error(f"Error handling prisoner link tab: {e}")
            try:
                driver.close()
            except:
                pass
            driver.switch_to.window(driver.window_handles[0])
            return False

    try:
        prisoner_container = driver.find_element(By.CLASS_NAME, "vehicle_prisoner_select")
    except Exception as e:
        logging.info("No prisoner selection container found on mission page.")
        return

    try:
        prisoner_links = prisoner_container.find_elements(By.TAG_NAME, "a")
        for link in prisoner_links:
            classes = link.get_attribute("class")
            if "btn-success" in classes:
                if open_link_in_new_tab(link):
                    logging.info(f"Opened normal prisoner vehicle with text: '{link.text}'")
                    return
                else:
                    logging.error("Failed to open normal prisoner vehicle link.")
        try:
            verbands_header = prisoner_container.find_element(By.XPATH, ".//h5[contains(text(),'Verbandszellen')]")
            verbands_links = prisoner_container.find_elements(By.XPATH, ".//h5[contains(text(),'Verbandszellen')]/following-sibling::a")
        except Exception as e:
            logging.info("No Verbandszellen links found in prisoner selection.")
            return

        if not verbands_links:
            logging.info("No Verbandszellen links available.")
            return

        selection_candidates = []
        for link in verbands_links:
            text = link.text
            perc_match = re.search(r'Abgabe an Besitzer:\s*(\d+)%', text)
            if perc_match:
                try:
                    percentage = int(perc_match.group(1))
                except:
                    percentage = 100
            else:
                percentage = 100
            selection_candidates.append((link, percentage))
        zero_candidates = [tpl for tpl in selection_candidates if tpl[1] == 0]
        if zero_candidates:
            chosen = zero_candidates[0][0]
        else:
            selection_candidates.sort(key=lambda tpl: tpl[1])
            chosen = selection_candidates[0][0]
        if open_link_in_new_tab(chosen):
            logging.info(f"Opened Verbandszellen prisoner vehicle with text: '{chosen.text}'")
        else:
            logging.error("Failed to open Verbandszellen prisoner vehicle link.")
    except Exception as e:
        logging.error(f"Error in prisoner vehicle selection: {e}")

def main():
    while True:
        try:
            email = os.getenv("EMAIL")
            password = os.getenv("PASSWORD")
            service = Service('./chromedriver.exe')
            driver = webdriver.Chrome(service=service)
            sleep(0.25)
            driver.get('https://www.leitstellenspiel.de/users/sign_in')
            sleep(0.25)
            email_login = driver.find_element(By.ID, 'user_email')
            email_login.send_keys(email)
            sleep(0.125)
            password_login = driver.find_element(By.ID, 'user_password')
            password_login.send_keys(password)
            sleep(0.125)
            user_remember_me = driver.find_element(By.ID, 'user_remember_me')
            user_remember_me.click()
            sleep(0.125)
            login = driver.find_element(By.NAME, 'commit')
            login.click()
            sleep(0.25)
            while True:
                try:
                    driver.get('https://www.leitstellenspiel.de/')
                    sleep(0.25)
                    try:
                        finishing_button = driver.find_element(By.ID, 'mission_select_finishing')
                        finishing_button.click()
                        logging.info("Finishing filter button clicked")
                    except Exception as e:
                        logging.warning(f"Could not click finishing filter button: {str(e)}")
                    sleep(0.25)
                    wait = WebDriverWait(driver, 10)
                    mission_list = wait.until(EC.presence_of_element_located((By.ID, 'mission_list')))
                    mission_entries = driver.find_elements(By.CLASS_NAME, 'missionSideBarEntry')
                    non_finishing_missions = [
                        m for m in mission_entries
                        if (
                            m.get_attribute("data-mission-state-filter") != "finishing" and 
                            "[Verband]" not in m.find_element(By.CLASS_NAME, 'map_position_mover').text
                        )
                    ]
                    if len(non_finishing_missions) > 4:
                        set_mission_speed(driver, "pause")
                        logging.info("Mission speed set to pause")
                    if len(non_finishing_missions) < 6:
                        set_mission_speed(driver, "7")
                        logging.info("Mission speed set to 7")
                    logging.info(f"Found {len(non_finishing_missions)} non-Verband missions")
                    sleep(0.225)
                    for mission in non_finishing_missions:
                        try:
                            sleep(0.15)
                            mission_caption = mission.find_element(By.CLASS_NAME, 'map_position_mover').text
                            if "Verband" in mission_caption:
                                logging.info("Verband mission detected, skipping.")
                                continue
                            sleep(0.05)
                            alarm_button = mission.find_element(By.CLASS_NAME, 'mission-alarm-button')
                            mission_url = alarm_button.get_attribute('href')
                            sleep(0.03125)
                            if mission_url:
                                driver.execute_script(f"window.open('{mission_url}','_blank')")
                                driver.switch_to.window(driver.window_handles[-1])
                                sleep(0.05)
                                if len(driver.window_handles) > 5:
                                    logging.warning("Too many tabs open, restarting...")
                                    driver.execute_script("window.open('https://www.leitstellenspiel.de','_blank')")
                                    for handle in driver.window_handles[:-1]:
                                        driver.switch_to.window(handle)
                                        driver.close()
                                    driver.switch_to.window(driver.window_handles[-1])
                                    continue
                                if check_mission_completed(driver):
                                    driver.close()
                                    driver.switch_to.window(driver.window_handles[0])
                                    continue

                                check_and_click_easter_egg(driver)
                                check_for_sprechwunsch(driver, wait)
                                select_prisoner_vehicle(driver)
                                mission_title = wait.until(EC.presence_of_element_located((By.ID, 'missionH1')))
                                if "Verband" in mission_title.text:
                                    logging.info("Verband mission detected in title, closing...")
                                    driver.close()
                                    driver.switch_to.window(driver.window_handles[-1])
                                    continue
                                current_vehicles, enroute_personnel = extract_current_vehicles(driver)
                                sleep(0.05)
                                help_button = driver.find_element(By.ID, 'mission_help')
                                help_url = help_button.get_attribute('href')
                                sleep(0.03125)
                                driver.execute_script(f"window.open('{help_url}','_blank')")
                                driver.switch_to.window(driver.window_handles[-1])
                                sleep(0.05)
                                wait.until(EC.presence_of_element_located((By.CLASS_NAME, 'col-md-4')))
                                col_md_4_divs = driver.find_elements(By.CLASS_NAME, 'col-md-4')
                                sleep(0.0625)
                                if len(col_md_4_divs) >= 2:
                                    vehicle_div = col_md_4_divs[1]
                                    requirements_table = vehicle_div.find_element(By.TAG_NAME, 'table')
                                    required_vehicles = extract_vehicle_requirements(requirements_table)
                                    min_patients, nef_probability = extract_patient_requirements(col_md_4_divs)
                                    sleep(0.05)
                                    driver.close()
                                    driver.switch_to.window(driver.window_handles[-1])
                                    sleep(0.05)
                                    required_vehicles = handle_patients_and_nef(
                                        driver, required_vehicles, current_vehicles, enroute_personnel, min_patients, nef_probability
                                    )
                                    missing_vehicles = calculate_missing_vehicles(required_vehicles, current_vehicles)
                                    sleep(0.05)

                                    # Immer handle_water_and_dispatch aufrufen
                                    closed_tab = handle_water_and_dispatch(driver, missing_vehicles)
                                    if closed_tab:
                                        logging.info("Tab closed due to insufficient LFs. Skipping further steps.")
                                        continue

                                    if missing_vehicles:
                                        logging.info("Vehicles dispatched")
                                        if "NEF" in missing_vehicles:
                                            required_vehicles["NAW"] = required_vehicles.get("NAW", 0) + 1
                                            logging.info("NEF not dispatched, dispatching NAW as backup")
                                            select_vehicles(driver, {"NAW": 1})
                                    else:
                                        logging.info("No additional vehicles needed")
                                else:
                                    driver.close()
                                    driver.switch_to.window(driver.window_handles[-1])
                                    sleep(0.05)
                                driver.close()
                                driver.switch_to.window(driver.window_handles[0])
                                sleep(0.125)
                        except Exception as e:
                            if ("unexpected alert open" in str(e).lower()
                                or "no such element" in str(e).lower()
                                or "element not found" in str(e).lower()):
                                try:
                                    Alert(driver).accept()
                                    Alert(driver).send_keys("Enter")
                                    sleep(0.05)
                                except:
                                    pass
                                continue
                            logging.error(f"Error processing mission: {str(e)}")
                            continue
                    logging.info("Mission processing completed, waiting 20 seconds")
                    sleep(20)
                except Exception as e:
                    logging.error(f"Error in mission loop: {str(e)}")
                    try:
                        driver.get('https://www.leitstellenspiel.de/')
                    except:
                        pass
                    sleep(30)
                    continue
        except Exception as e:
            logging.error(f"Critical error occurred: {str(e)}")
            try:
                driver.execute_script("window.open('https://www.leitstellenspiel.de','_blank')")
                for handle in driver.window_handles[:-1]:
                    driver.switch_to.window(handle)
                    driver.close()
                driver.switch_to.window(driver.window_handles[-1])
            except:
                pass
            sleep(25)
            continue

if __name__ == "__main__":
    while True:
        try:
            main()
        except:
            sleep(25)
            continue