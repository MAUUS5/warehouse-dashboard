# SIMPLE LOCAL UPDATE (No GitHub Actions)

If you prefer to update the PowerPoint yourself without GitHub automation:

## Requirements
- Python 3 installed on your computer
- Excel files in the same folder as the script

## Step 1: Install Required Libraries

Open Command Prompt or Terminal and run:
```bash
pip install openpyxl python-pptx
```

## Step 2: Run the Update Script

1. Put these files in the same folder:
   - `update_powerpoint.py`
   - `Health_Tracker_2026_xlsx.xlsx`
   - `Picker_Efficiency_2026_xlsx.xlsx`
   - `Putaway_Efficiency_2026_xlsx.xlsx`

2. Open Command Prompt in that folder

3. Run:
   ```bash
   python update_powerpoint.py
   ```

4. The script will create: `US5_Warehouse_KPI_Dashboard.pptx`

5. Upload the PowerPoint to GitHub manually

## That's it!

Every time you want to update:
1. Update your Excel files
2. Run `python update_powerpoint.py`
3. Upload new PowerPoint to GitHub
4. Done!

---

## Even Simpler: Just Send Me Your Excel

Don't want to deal with Python?

1. Upload your updated Excel files to this chat
2. I'll generate the new PowerPoint
3. Download and upload to GitHub
4. Takes 2 minutes!
