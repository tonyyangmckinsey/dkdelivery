import shutil
from datetime import datetime
import os
today_str = datetime.today().strftime("%Y%m%d")  # Format: YYYYMMDD


# Config
source_file = r"C:\Users\tonyan\OneDrive - GlobalConnect A S\Hubble H1 2025\DK backlog tracker\DK backlog tracker.xlsx"
destination_folder = r"C:\Users\tonyan\OneDrive - GlobalConnect A S\Documents\Python\dkdelivery\3. DKTrackerSnapshots"
os.makedirs(destination_folder, exist_ok=True)

timestamp = datetime.now().strftime("%Y%m%d_%H%M")
destination_file = os.path.join(destination_folder, f"{today_str}_DK_backlog_snapshot.xlsx")

# Copy
shutil.copy2(source_file, destination_file)
print(f"✅ Snapshot saved: {destination_file}")
