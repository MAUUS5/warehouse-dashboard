# PowerPoint Dashboard - Easy Update Guide

## ✅ SETUP (One-Time)

### Step 1: Upload Files to GitHub

Upload these files to your repository:
1. `update_powerpoint.py` - The update script
2. `.github/workflows/update-powerpoint.yml` - Automation workflow
3. Your 3 Excel files (if not already uploaded)

### Step 2: Enable GitHub Actions Permissions

1. Go to: Settings → Actions → General
2. Under "Workflow permissions":
   - Select ✅ "Read and write permissions"
   - Check ✅ "Allow GitHub Actions to create and approve pull requests"
3. Click "Save"

---

## 🔄 HOW TO UPDATE (Three Easy Ways)

### **OPTION A: Upload Excel Files (Automatic)**

**Simplest! PowerPoint auto-updates when you upload Excel.**

1. Update your Excel files on your computer
2. Go to GitHub repository
3. Upload updated Excel files (replace old ones)
4. Commit changes
5. **Wait 2 minutes**
6. PowerPoint automatically regenerates!
7. Download updated PowerPoint from GitHub

**That's it!** No manual work needed.

---

### **OPTION B: Manual Trigger (On Demand)**

**Run the update manually anytime.**

1. Go to: Actions → "Update PowerPoint Dashboard"
2. Click "Run workflow"
3. Click green "Run workflow" button
4. **Wait 2 minutes**
5. Download updated PowerPoint

---

### **OPTION C: Scheduled (Automatic Daily)**

**Set it and forget it!**

- Runs automatically every day at 6:00 AM EST
- Pulls latest data from Excel files
- Regenerates PowerPoint
- No action needed from you!

---

## 📥 HOW TO GET UPDATED POWERPOINT

### Method 1: Download from GitHub
1. Go to repository
2. Click on `US5_Warehouse_KPI_Dashboard.pptx`
3. Click "Download" button
4. Open and display!

### Method 2: Direct Download Link
Use this link (replace YOUR_USERNAME):
```
https://github.com/YOUR_USERNAME/warehouse-dashboard/raw/main/US5_Warehouse_KPI_Dashboard.pptx
```

Anyone with access to the repository can download the latest version!

---

## 📺 DISPLAY ON TV

1. Download PowerPoint from GitHub
2. Open on computer connected to TV
3. Press **F5** (start slideshow)
4. **Slideshow → Set Up Show:**
   - ✅ Loop continuously until 'Esc'
5. **Transitions → Timing:**
   - ✅ After: 15 seconds (or your preferred time)
   - Apply to all slides
6. Done! Slides will auto-advance forever

---

## 🔧 TROUBLESHOOTING

**Q: PowerPoint didn't update after uploading Excel**
A: Check Actions tab for errors. Make sure permissions are set correctly.

**Q: How do I change employee names?**
A: Edit `update_powerpoint.py` line 200+ (Employee Recognition section)

**Q: How do I change the schedule time?**
A: Edit `.github/workflows/update-powerpoint.yml` line 14
   - Current: `'0 11 * * *'` = 6 AM EST
   - Change to: `'0 13 * * *'` = 8 AM EST
   - Reference: https://crontab.guru

**Q: Can I customize the design?**
A: Yes! Edit `update_powerpoint.py` - change colors, fonts, layout, etc.

---

## ✅ BENEFITS

✅ **No caching issues** - Always correct numbers  
✅ **Easy updates** - Just upload Excel files  
✅ **Automatic** - Runs daily on schedule  
✅ **Shareable** - Anyone can download from GitHub  
✅ **Offline ready** - Works without internet  
✅ **Professional** - PowerPoint quality  

---

## 📊 CURRENT DATA SOURCES

The PowerPoint pulls from:
- `Health_Tracker_2026_xlsx.xlsx` - Row 34 (totals), Row 6-33 (daily data)
- `Picker_Efficiency_2026_xlsx.xlsx` - Formulas sheet
- `Putaway_Efficiency_2026_xlsx.xlsx` - Formulas sheet

Safety days auto-calculated from: November 29, 2023 (start date)

---

## 🎯 SUMMARY

**To update dashboard:**
1. Update Excel files
2. Upload to GitHub
3. Done! PowerPoint auto-regenerates

**No HTML, no caching, no problems!** 🎉
