from datetime import datetime
from time import sleep
import logging

import pandas as pd
import gspread
from google.oauth2.service_account import Credentials



# ── Constants ────────────────────────────────────────────────────────────────

EXCEL_FILE         = r"C:\Users\himan\Documents\UiPath\CrickInfo\Data\IPLPointsData.xlsx"
SERVICE_ACCOUNT    = r"C:\Users\himan\Documents\UiPath\CrickInfo\Data\service_account.json"
SHEET_NAME         = "IPLTeam_S2"
TEAMS_FILE         = "teams.txt"
LOG_FILE_PATH      = r"C:\Users\himan\Documents\UiPath\CrickInfo\Data\logs\ipl_updater_101.log"   # ← replace this

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

SKIP_NAMES         = {'Player Name', 'Last Match', 'Team', 'Sr. No', '', 'Name'}



UPDATE_DELAY_SEC   = 3
BIG_PAUSE_SEC   = 5

TODAYS_DATE_LABEL    = datetime.now().strftime("%d/%m")  # e.g. "15/04"
#TODAYS_DATE_LABEL    = "19/04"

PLAYER_NOT_UPDATED = {}  # To track players that were not updated for any reason


# ── Logging setup ─────────────────────────────────────────────────────────────
def setup_logging(log_file: str) -> logging.Logger:
    """
    Configure and return a logger that writes to a log file only.
    Print statements handle console output; logger handles file output.
    """
    logger = logging.getLogger("IPLUpdater")
    logger.setLevel(logging.DEBUG)

    # Remove all existing handlers to prevent duplicate logs
    if logger.hasHandlers():
        logger.handlers.clear()

    formatter = logging.Formatter(
        fmt="[%(asctime)s] %(levelname)-8s - %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    # File handler only — console output is handled by existing print statements
    file_handler = logging.FileHandler(log_file, mode='a', encoding='utf-8')
    file_handler.setLevel(logging.DEBUG)
    file_handler.setFormatter(formatter)

    logger.addHandler(file_handler)

    return logger


# ── Helpers ───────────────────────────────────────────────────────────────────


# ── Load Excel sheets ─────────────────────────────────────────────────────────
def load_excel_sheets(filepath: str, logger: logging.Logger) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Load player points (Sheet1) and bowler dot-ball data (Sheet2) from Excel."""
    df_players = pd.read_excel(filepath, sheet_name='Sheet1')
    df_bowlers = pd.read_excel(filepath, sheet_name='Sheet2')
    logger.info(f"Excel loaded: {len(df_players)} players, {len(df_bowlers)} bowlers from '{filepath}'")
    return df_players, df_bowlers

# ── Google Sheets authentication ─────────────────────────────────────────
def authenticate_gsheet(service_account_file: str, logger: logging.Logger) -> gspread.Client:
    """Authenticate with Google Sheets using a service account and return a client."""
    creds = Credentials.from_service_account_file(service_account_file, scopes=SCOPES)
    client = gspread.authorize(creds)
    logger.info("Google Sheets authentication successful.")
    return client


# ── Open worksheets ─────────────────────────────────────────────────────────\
def open_worksheets_d(
    client: gspread.Client,
    sheet_names: list[str],
    logger: logging.Logger,
) -> dict[str, gspread.Worksheet]:
    """Open and return all required worksheets keyed by their logical name."""
    workbook = client.open(SHEET_NAME)
    sheets = {}
    for name in sheet_names:
        try:
            sheets[name] = workbook.worksheet(name)
        except gspread.exceptions.WorksheetNotFound:
            logger.error(f"Worksheet '{name}' not found in '{SHEET_NAME}'. Please check the sheet name and try again.")
            raise
    if len(sheets) == 0:
        logger.error(f"No worksheets were opened. Please check the sheet names and try again.")
        raise Exception("No worksheets opened.")
    
    logger.info(f"Worksheets opened: {list(sheets.keys())}")
    return sheets


# ── Update player row in sheet ───────────────────────────────────────────────
def update_player_row_data(
    sheet: gspread.Worksheet,
    sheet_row: int,
    new_pts: int,
    dot_balls: int,
    prev_total_pts: int,
    player_name: str,
    sheet_label: str,
    logger: logging.Logger,
) -> None:
    """
    Write previous and new point values to the correct columns in a sheet row.

    Args:
        sheet:        The worksheet to update.
        sheet_row:    1-based row index.
        prev_pts:     Previous points to write into COL_PREV_PTS_WRITE.
        new_pts:      New points to write into COL_NEW_PTS_WRITE.
        player_name:  Used only for the success log message.
        sheet_label:  Name of the sheet/team for the log message.
        logger:       Logger instance.
        label:        Unit label shown in the log message (e.g. 'points' or 'Dot Balls').
    """
    sheet.update_cell(sheet_row, 5, new_pts)
    sheet.update_cell(sheet_row, 4, dot_balls)
    print(f"Points gained : {new_pts + dot_balls - prev_total_pts} (Points: {new_pts}, Dot Balls: {dot_balls}, Previous Total: {prev_total_pts})")
    logger.info(f"Points gained for {player_name} in {sheet_label}: {new_pts + dot_balls - prev_total_pts} (Points: {new_pts}, Dot Balls: {dot_balls}, Previous Total: {prev_total_pts})")
    msg = f"Data updated successfully for {player_name} in {sheet_label} - {new_pts} points and {dot_balls} dot balls!"
    print(msg)
    logger.info(msg)


# ── Find and update player points in a sheet ─────────────────────────────────
def find_and_update_player_points_dotballs(
    sheet: gspread.Worksheet,
    google_data: list[list],
    player_name: str,
    new_pts: int,
    dot_balls: int,
    sheet_label: str,
    logger: logging.Logger,
    update_flag: bool = True,
) -> bool:
    """
    Search for a player in a sheet's data and update their points if found.

    Returns True if the player was found (regardless of whether an update was needed).
    """
    for row_idx, row in enumerate(google_data):
        cell_name = row[1]

        if cell_name.strip() in SKIP_NAMES:
            continue

        if cell_name == player_name:
            msg = f"Match found for {player_name} in {sheet_label} at row {row_idx + 1}"
            print(msg)
            logger.debug(msg)

            total_pts = new_pts + dot_balls

            prev_total_pts = int(row[7]) if row[7] != '' else 0

            old_total_pts = int(row[5]) if row[5] != '' else 0

            sheet_row = row_idx + 1

            if prev_total_pts == total_pts:             
                if update_flag:
                    PLAYER_NOT_UPDATED[player_name] = "Points unchanged"
                msg = f"No update needed for {player_name} in {sheet_label} as points are unchanged."
                print(msg)
                logger.info(msg)
                return True

            
            update_player_row_data(sheet, sheet_row, new_pts, dot_balls,prev_total_pts, player_name, sheet_label, logger)
            print(f"Waiting for {UPDATE_DELAY_SEC} seconds before next updates...")
            logger.debug(f"Sleeping {UPDATE_DELAY_SEC}s before next update.")
            sleep(UPDATE_DELAY_SEC)
            return True

    return False


# ── Process players and update points in team sheets ─────────────────────────
def process_players_pointsandDotballs(
    df_players_points: pd.DataFrame,
    sheet: gspread.Worksheet,
    points_data: list[list],
    logger: logging.Logger,
) -> None:
    """
    Iterate over all players from the Excel sheet and update their points
    across the team worksheets.
    """
    Team_Dict = {}
    proccessing_Count = 0
    print("\n*******************************************************************")
    #print(f"--------------- Processing players for team: {team1} ---------------")
    logger.debug("===============================================================")
    #logger.info(f"Processing players for team: {team1}")
    for _, row in df_players_points.iterrows():
        
        player_name = row['Name']
        new_pts     = row['Points']
        team_name   = row['Team Name']
        dot_balls   = row['Dot Balls']
        
        
        print("\n***************************************************************")
        logger.debug("===============================================================")
        proccessing_Count += 1
        print(f"Processing count= {proccessing_Count}")
        logger.info(f"Processing count= {proccessing_Count}")
        if new_pts == 0 or new_pts == '0':
            PLAYER_NOT_UPDATED[player_name] = "Points is zero"
            msg = f"\nSkipping {player_name} as points is zero."
            print(msg)
            logger.info(msg)
            continue

        
        print(f"\nChecking Player of {team_name} = {player_name}  - {new_pts}\n")
        logger.info(f"Processing player of {team_name} : '{player_name}' | {new_pts} points | team: {team_name}")

        
        found = False
        
        
        found = find_and_update_player_points_dotballs(
            sheet=sheet,
            google_data=points_data,
            player_name=player_name,
            new_pts=int(new_pts),
            dot_balls=int(dot_balls),
            sheet_label="Points",
            logger=logger,
            update_flag=True,
        )
            

        if not found:
            PLAYER_NOT_UPDATED[player_name] = "Player not found in sheet"
            msg = f"{player_name} not found in (Points)."
            print(msg)
            logger.warning(msg)
        Team_Dict[team_name] = Team_Dict.get(team_name, 0) + 1
    for team, count in Team_Dict.items():
        logger.info(f"Total players processed for team '{team}': {count}")
        print(f"Total players processed for team '{team}': {count}")


# ── Update 'Points Gained Today' in Data sheet ─────────────────────────────
def update_points_gained(Home_sheet, Data_sheet, logger) -> None:
    """
    Fetch 'Points Gained Today' for each team from Home sheet (O2:P6)
    and write them into the Data sheet under today's date column.

    Home sheet layout (O2:P6):
        Col O: Team abbreviation (HW, MC, OL, AS, PA)
        Col P: Points gained today (e.g. +81, +182)

    Data sheet layout:
        Col A: Team abbreviation
        Row 1: Date headers formatted as "dd/MM"
    """

    # ── 1. Read K2:L6 from Home sheet ────────────────────────────────────────
    try:
        home_data = Home_sheet.get("K2:L6")  # [[abbr, points], ...]
    except Exception as e:
        logger.error(f"Failed to read Home sheet K2:L6: {e}")
        return

    # ── 2. Parse points into {abbr: points} ──────────────────────────────────
    points: dict[str, int] = {}
    for row in home_data:
        if len(row) < 2:
            continue
        abbr = row[0].strip()                          # e.g. "HW"
        pts_str = str(row[1]).replace("+", "").replace(",", "").strip()  # "+81" → "81"
        try:
            points[abbr] = int(pts_str)
        except ValueError:
            logger.warning(f"Could not parse points for '{abbr}': '{row[1]}' — skipping.")

    if not points:
        logger.warning("No points extracted from O2:P6 — aborting update.")
        return

    logger.info(f"Points Gained Today fetched: {points}")

    # ── 3. Find today's date column in Data sheet ─────────────────────────────
    #today_label = datetime.today().strftime("%d/%m")   # e.g. "15/04"
    today_label = TODAYS_DATE_LABEL

    header_row = Data_sheet.row_values(1)              # ["Team", "Points", "15/04", ...]
    try:
        col_index = header_row.index(today_label) + 1  # 1-based column number
    except ValueError:
        logger.error(f"Date column '{today_label}' not found in Data sheet headers: {header_row}")
        return

    logger.info(f"Writing to column {col_index} ({today_label})")

    # ── 4. Match team rows and batch write ────────────────────────────────────
    team_col = Data_sheet.col_values(1)                # Col A: team abbreviations

    updates = []
    for abbr, pts in points.items():
        try:
            row_index = team_col.index(abbr) + 1       # 1-based row number
        except ValueError:
            logger.warning(f"Team '{abbr}' not found in Data sheet col A — skipping.")
            continue

        cell = gspread.utils.rowcol_to_a1(row_index, col_index)
        updates.append({"range": cell, "values": [[pts]]})
        logger.debug(f"  {abbr} → {cell} = {pts}")

    if updates:
        Data_sheet.batch_update(updates)
        logger.info(f"Successfully updated {len(updates)} team(s) for {today_label}.")
    else:
        logger.warning("No cells were updated.")

    pass


# ── Update TOP 12 points in Data sheet ─────────────────────────────────
def update_top12_points(Home_sheet, Data_sheet, logger) -> None:
    """
    Fetch TOP 12 points for each team from Home sheet (J2:K6)
    and write them into today's date row in the Data table (starting row 9) of Data sheet.

    Home sheet layout (J2:K6):
        Col J: Full team name (e.g. "Pranav Astra")
        Col K: TOP 12 points (e.g. 2605)

    Data sheet new table layout (starts row 9):
        Row 9 : Headers → Date | HW | MC | OL | AS | PA
        Col A : Dates in "dd/MM" format (e.g. "15/04")
        Col B : HW points
        Col C : MC points
        Col D : OL points
        Col E : AS points
        Col F : PA points
    """

    NAME_TO_ABBR = {
        "Himanshu Warriors": "HW",
        "Mehul Challengers":  "MC",
        "Onkar Legends":      "OL",
        "Abhishek Strikers":  "AS",
        "Pranav Astra":       "PA",
    }


    # ── 1. Read TOP 12 points from Home sheet (F2:G6) ────────────────────────
    try:
        home_data = Home_sheet.get("F2:G6")  # [[team_name, top12_pts], ...]
    except Exception as e:
        logger.error(f"Failed to read Home sheet F2:G6: {e}")
        return

    # Parse into {abbr: top12_points}
    points: dict[str, int] = {}
    for row in home_data:
        if len(row) < 2:
            continue
        team_name = row[0].strip()
        pts_str = str(row[1]).replace(",", "").strip()
        abbr = NAME_TO_ABBR.get(team_name)
        if abbr:
            try:
                points[abbr] = int(pts_str)
            except ValueError:
                logger.warning(f"Could not parse TOP 12 points for '{team_name}': '{row[1]}'")

    if not points:
        logger.warning("No TOP 12 points extracted from Home sheet — aborting.")
        return

    logger.info(f"TOP 12 points fetched: {points}")

    # ── 2. Read the Data table headers (row 9) to get column mapping ──────────
    # Row 9: ["Date", "HW", "MC", "OL", "AS", "PA"]
    try:
        header_row = Data_sheet.row_values(9)  # 1-based row index
    except Exception as e:
        logger.error(f"Failed to read Data sheet row 9: {e}")
        return

    # Build {abbr: col_index (1-based)} from header row
    abbr_to_col: dict[str, int] = {}
    for col_idx, header in enumerate(header_row, start=1):
        if header.strip() in points:
            abbr_to_col[header.strip()] = col_idx

    if not abbr_to_col:
        logger.error(f"No matching team headers found in row 9: {header_row}")
        return

    logger.info(f"Column mapping: {abbr_to_col}")

    # ── 3. Find today's date row in Col A (starting from row 10) ─────────────
    #today_label = datetime.today().strftime("%d/%m")  # e.g. "15/04"
    today_label = TODAYS_DATE_LABEL

    try:
        date_col = Data_sheet.col_values(1)  # All values in Col A (1-based)
    except Exception as e:
        logger.error(f"Failed to read Data sheet Col A: {e}")
        return

    # date_col is 0-indexed list; row index = list index + 1
    try:
        row_index = date_col.index(today_label) + 1  # 1-based row number
    except ValueError:
        logger.error(f"Today's date '{today_label}' not found in Data sheet Col A.")
        return

    logger.info(f"Today '{today_label}' found at row {row_index}")

    # ── 4. Batch write points into today's row ────────────────────────────────
    updates = []
    for abbr, pts in points.items():
        col_index = abbr_to_col.get(abbr)
        if col_index is None:
            logger.warning(f"No column found for team '{abbr}' — skipping.")
            continue

        cell = gspread.utils.rowcol_to_a1(row_index, col_index)
        updates.append({"range": cell, "values": [[pts]]})
        logger.debug(f"  {abbr} → {cell} = {pts}")

    if updates:
        Data_sheet.batch_update(updates)
        logger.info(f"Successfully updated {len(updates)} team(s) for '{today_label}' at row {row_index}.")
    else:
        logger.warning("No cells were updated.")

    days_count = int(Data_sheet.get("G8")[0][0]) + 1

    Data_sheet.update_cell(8, 7, str(days_count))

# ── Update past TOP 12 points in Data sheet from Home sheet ─────────────────────────
def update_past_top12_points(Home_sheet, Data_sheet, logger) -> None:
    """
    (Optional) Fetch TOP 12 points for each team from Home sheet (F2:G6)
    and write them into the correct date row in the Data table of Data sheet.
    This can be used to backfill or correct past TOP 12 points if needed.
    """
     # ── 1. Read TOP 12 points from Home sheet (F2:G6) ────────────────────────
    try:
        home_data = Home_sheet.get("F2:G6")  # [[team_name, top12_pts], ...]
    except Exception as e:
        logger.error(f"Failed to read Home sheet F2:G6: {e}")
        return
    
    for idx, row in enumerate(home_data):
        if len(row) < 2:
            continue
        
        pts_str = str(row[1]).replace(",", "").strip()
        team_name = row[0].strip()
        switcher = {
            "Himanshu Warriors": 2,
            "Mehul Challengers": 3,
            "Onkar Legends": 4,
            "Abhishek Strikers": 5,
            "Pranav Astra": 6,
        }

        Data_sheet.update_cell(switcher.get(team_name), 2, pts_str)
        if idx == 4:  # Assuming you want to update the 5th row (0-indexed)
            break
    
# ── Update previous points for all players in team sheets ─────────────────────────
def update_players_prev_points(sheet: gspread.Worksheet, google_data: list[list], sheet_label: str, logger: logging.Logger) -> None:
    """
    Iterate through the sheet's data and copy current points from COL_PREV_PTS to COL_PREV_PTS_WRITE
    for all players. This ensures that the "previous points" column is up-to-date before processing new updates.
    """
    
    logger.info("Updating previous points for all players before processing.")
    print("Updating previous points for all players before processing...")
    logger.debug(f"======================================================")
    print(f"======================================================")
    for row_idx, row in enumerate(google_data):
        cell_name = row[1]

        if cell_name.strip() in SKIP_NAMES:
            continue
        
        # If points are unchanged, we still want to ensure COL_PREV_PTS_WRITE is updated to reflect the current points for accurate future comparisons.
        
        pts = int(row[7]) if row[7] != '' else 0
        prev_pts = int(row[5]) if row[5] != '' else 0
        if prev_pts != pts:
            sheet_row = row_idx + 1
            sheet.update_cell(sheet_row, 6, pts)  # Assuming COL_PREV_PTS_WRITE is the 6th column (index 5)
            sleep(UPDATE_DELAY_SEC)  # Sleep to avoid hitting rate limits if there are many updates
            logger.debug(f"Updated previous points for '{cell_name}' at row {row_idx + 1}: {prev_pts} → {pts}")
            print(f"Updated previous points for '{cell_name}' at row {row_idx + 1}: {prev_pts} → {pts}")
            logger.debug(f"Sleeping {UPDATE_DELAY_SEC}s after updating previous points.")
            print(f"Sleeping {UPDATE_DELAY_SEC}s after updating previous points.")
        
    
    logger.debug(f"======================================================")
    print(f"======================================================")
    return




#////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    import os
    base, ext = os.path.splitext(LOG_FILE_PATH)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_file_with_timestamp = f"{base}_{timestamp}{ext}"
    logger = setup_logging(log_file_with_timestamp)
#/ Main orchestrator: loads data, authenticates, then runs player and bowler updates.
def main() -> None:
    """Main orchestrator: loads data, authenticates, then runs player and bowler updates."""
    logger = setup_logging(LOG_FILE_PATH.replace('101', datetime.now().strftime("%Y%m%d_%H%M%S")))

    print("Process started...")
    logger.info("═════════════════════ Process started ═════════════════════")

    try:
        #initialize teams, load excel data, authenticate and open sheets
        df_sheet1, df_sheet2 = load_excel_sheets(EXCEL_FILE, logger)

        client = authenticate_gsheet(SERVICE_ACCOUNT, logger)
        ws      = open_worksheets_d(client, ["Points","Home","Data"], logger)  

        
        points_data = ws["Points"].get_all_values()
        
        #update previous points for all players in team sheets before processing to ensure we have the latest points for comparison with excel data//////////
        flag = input("Do you want to update previous points for all players before processing in team sheets? (y/n): ").strip().lower()
        if flag == 'y':
            update_players_prev_points(ws["Points"], points_data, "Points", logger)
        
        #update date in teams sheets before processing/////////////////
        #update_date([ws[t1],ws[t2]], logger)

        #update past top 12 points in data sheet before processing players to ensure we have the latest points for comparison////////////
        flag = input("Do you want to update past top 12 points in Data sheet before processing players? (y/n): ").strip().lower()
        if flag == 'y':
            update_past_top12_points(ws["Home"], ws["Data"], logger)

    
        print("\n*******************************************************************")
        logger.debug("===============================================================")

        print(f"\nWaiting for {BIG_PAUSE_SEC} seconds before next steps...")
        logger.info(f"Pausing {BIG_PAUSE_SEC}s before before next steps.")

        #sleep(BIG_PAUSE_SEC)
        print("\n*******************************************************************")
        logger.debug("===============================================================")

        flag = input("Do you want to process player points and dot balls updates now? (y/n): ").strip().lower()
        if flag == 'y':
            process_players_pointsandDotballs(df_sheet2, ws["Points"], points_data, logger)

        

        print("\n*******************************************************************")
        logger.debug("===============================================================")

        print(f"\nPlayers not updated: {len(PLAYER_NOT_UPDATED)}")
        
        for player, reason in PLAYER_NOT_UPDATED.items():
            logger.warning(f"Player not updated: {player} | Reason: {reason}")
            print(f"Player not updated: {player} | Reason: {reason}")

        
        print("\n*******************************************************************")
        logger.debug("===============================================================")

        flag = input("Do you want to update points gained today and top 12 points in Data sheet now? (y/n): ").strip().lower()
        if flag == 'y':
            flag = input("Do you want to change todays date lable?(y/n): ").strip().lower()
            if flag == 'y':
                global TODAYS_DATE_LABEL
                TODAYS_DATE_LABEL = input("Please enter today's date label(DD): ")
                TODAYS_DATE_LABEL = TODAYS_DATE_LABEL.zfill(2) + "/" + datetime.now().strftime("%m")  # e.g. "15/04"
                logger.info(f"TODAYS_DATE_LABEL updated to: {TODAYS_DATE_LABEL}")
            # Re-open worksheets to get the latest data for points gained and top 12 updates
            ws2     = open_worksheets_d(client, ["Home", "Data"], logger)

            #update points gained today in data sheet////////////////////
            update_points_gained(ws2["Home"], ws2["Data"], logger)


            #update top 12 points in data sheet//////////////////////////
            update_top12_points(ws2["Home"], ws2["Data"], logger)


        print("\nAll updates completed!")
        logger.info("═════════════════════ All updates completed ═════════════════════")

    except Exception as e:
        print(f"Fatal error: {e}")
        logger.exception(f"Fatal error — process aborted: {e}")
        raise


# ── Entry point ───────────────────────────────────────────────────────────────
if __name__ == "__main__":
    main()