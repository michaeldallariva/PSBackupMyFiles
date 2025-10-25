# PowerShell Backup Script - User Guide

## Quick Start Guide

This script automatically backs up your files from Desktop, Documents, and Downloads folders to either a local drive or network share. It organizes files by type (PDF, Word, Excel, etc.) and keeps a log of everything it copies.
It will also backup any folder you specify in the variables.

---

## Step 1: Configure the Script (First Time Setup)

Open the script file `PSBackupMyFiles_v0.4.ps1` in Notepad or any text editor.

### What You MUST Change:

**1. Set Your Backup Destination** (around line 48)

Find this section:
```powershell
# 👇 SET YOUR BACKUP DESTINATION HERE 👇
$BackupRoot = "C:\Backup"                    # Default: Local C: drive
# $BackupRoot = "D:\Backup"                  # Alternative: Local D: drive
# $BackupRoot = "\\192.168.1.10\mybackup"    # Alternative: Network share by IP
```

**Choose ONE option:**
- For **local backup**: Leave as is or change to `D:\Backup` or `E:\Backup`
- For **network backup**: Remove the `#` from the network line and add `#` to the local line

Example for network backup:
```powershell
# $BackupRoot = "C:\Backup"                    # Commented out
$BackupRoot = "\\192.168.1.10\mybackup"        # Active now
```

**2. Update Specific Folders to Backup** (around line 59 - OPTIONAL)

If you want to backup entire folders (like "My Projects" or "Tax Documents"), uncomment and change the folder paths:

```powershell
$FoldersToBackup = @(
  "C:\Users\YourName\Desktop\Important Folder"
  "C:\Users\YourName\Documents\Tax Documents"
)
```

⚠️ **Replace "YourName" with your actual Windows username!**

**That's it!** Save the file.

---

## Step 2: Test the Script Manually

1. Right-click on the script file
2. Select **"Run with PowerShell"**
3. Watch it backup your files
4. Check the backup location to confirm files were copied

---

## Step 3: Schedule Daily Automatic Backups (Optional)

### Setting Up Windows Task Scheduler

**Step-by-Step Instructions:**

1. **Open Task Scheduler**
   - Press `Windows Key + R`
   - Type: `taskschd.msc`
   - Press Enter

2. **Create a New Task**
   - Click **"Create Basic Task..."** in the right panel
   - Name it: `Daily File Backup`
   - Description: `Automatic backup of my important files`
   - Click **Next**

3. **Set the Trigger (When to Run)**
   - Select: **Daily**
   - Click **Next**
   - Set your preferred time (e.g., 9:00 PM when you're done working)
   - Click **Next**

4. **Set the Action (What to Run)**
   - Select: **Start a program**
   - Click **Next**
   - In **Program/script** field, type:
     ```
     pwsh.exe
     ```
     (Or use `powershell.exe` if you don't have PowerShell 7 installed)
   
   - In **Add arguments** field, type:
     ```
     -ExecutionPolicy Bypass -File "C:\Path\To\Your\Script\PSBackupMyFiles_v0_4_NetworkSupport.ps1"
     ```
     ⚠️ **Change the path to where you actually saved the script!**

   - Click **Next**

5. **Finish Setup**
   - Check **"Open the Properties dialog"** checkbox
   - Click **Finish**

6. **Advanced Settings (Important!)**
   - In the Properties window that opens:
   - Go to the **General** tab:
     - Check: **"Run whether user is logged on or not"**
     - Check: **"Run with highest privileges"**
   - Go to the **Conditions** tab:
     - Uncheck: **"Start the task only if the computer is on AC power"** (if laptop)
   - Go to the **Settings** tab:
     - Check: **"Run task as soon as possible after a scheduled start is missed"**
   - Click **OK**
   - Enter your Windows password when prompted

7. **Test It!**
   - Right-click on your new task in Task Scheduler
   - Select **"Run"**
   - Check if the backup runs successfully

---

## Common Questions

**Q: Where are my backed-up files?**
- They're in the location you set in `$BackupRoot`, organized by date and file type
- Example: `C:\Backup\26-10-2025\pdf\` (for PDF files backed up today)

**Q: Will the script delete my original files?**
- No! It only **copies** files. Your originals stay untouched.

**Q: What if I run it multiple times in one day?**
- It's smart! It skips files that are already backed up and identical.

**Q: The script says "cannot be loaded because running scripts is disabled"**
- You need to enable PowerShell scripts:
  1. Right-click PowerShell and run as Administrator
  2. Type: `Set-ExecutionPolicy RemoteSigned`
  3. Press Enter and type `Y` to confirm

**Q: Network backup isn't working?**
- Make sure the network share is accessible (try opening it in File Explorer)
- Verify you have write permissions to the network folder
- Check that the network path starts with `\\` (double backslash)

**Q: How do I stop automatic backups?**
- Open Task Scheduler
- Find "Daily File Backup"
- Right-click and select **Disable** or **Delete**

---

## What the Script Backs Up

By default, it backs up these file types:
- **Documents**: .pdf, .doc, .docx, .xls, .xlsx, .ppt, .pptx, .csv, .txt, .pub
- **eBooks**: .epub, .mobi, .azw, .azw3

You can add more file types by editing the `$Extensions` array in the script (around line 57).

---

## Tips

✅ **Best Practice**: Test the script manually a few times before setting up Task Scheduler

✅ **Network Backups**: Make sure your network storage is always on when the task runs

✅ **Check the Logs**: The script creates CSV files showing what was backed up - review them occasionally

✅ **Storage Space**: Monitor your backup location to ensure you don't run out of space

---

Need help? Check the CSV log files in your backup folder - they show exactly what happened during each backup!
