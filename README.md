# ILA OADM Report - Advanced Data Filter and Merger 📊🚀

An advanced, GUI-based standalone Windows application designed to aggregate, parse, clean, and format critical telecom infrastructure network alarm logs. It features a clean, sleek Dark Mode desktop interface built with Tkinter, processing high-volume Excel and CSV dump sheets within seconds into a single executable file (`.exe`).

No Python installation required for end-users! Just download and run.

---

## 🔥 Key Features

* **Standalone Windows Executable (.exe):** Runs directly on any Windows machine without needing Python, libraries, or command-line setups.
* **Multi-Source Data Aggregation:** Processes SOC (MFL), HT (High Temperature), and DG Manual dumps simultaneously.
* **Intelligent Owner Mapping:** Auto-extracts and classifies site ownership (`P1`, `RP1`, `P1 colo`) using dynamic logic parsing based on SAP IDs.
* **State & Circle Normalization:** Multi-layered Regex engine extracts standard telecom circle prefixes and accurately maps them to full State names using custom dictionary lookups.
* **Fully Formatted Outputs:** Generates a multi-sheet unified `.xlsx` workbook featuring automated autowidth structural grid borders, clear typography settings, and highlighted headers using `openpyxl`.

---

## 🚀 How to Run the App (For End-Users)

Aapko is app ko chalane ke liye apne system mein Python install karne ki koi zarurat nahi hai. Simply niche diye gaye steps follow karein:

1. **Download the App:** Go to the [Releases](https://github.com/your-username/ila-oadm-data-merger/releases) page and download the `ILA_OADM_Merger.exe` file.
2. **Launch:** Double-click on `ILA_OADM_Merger.exe` to open the clean Dark Mode GUI dashboard.
3. **Select Folders:**
   * **MFL Folder:** Click **Browse** and select the folder containing your `MFL*.xlsx` trackers.
   * **HT Folder:** Click **Browse** and select the directory containing the `Export-*.xlsx` ticketing dumps.
   * **DG Manual Folder:** Click **Browse** and select the root directory hosting the recursive CSV audit reports (`lfl`/`sfl` structures).
   * **Output Save Folder:** Pick the destination directory where the merged dashboard should drop.
4. **Execute:** Click **START MERGE PROCESS** to execute the pipeline! 🎉 Once successful, a pop-up alert will show you the exact location of your final file.

---

## 🛠️ How to Build the .exe File (For Developers)

Agar aap code mein koi change karte hain aur dubara se new `.exe` file generate karna chahte hain, toh in steps ko follow karein:

### Step 1: Clone & Setup Environment
```bash
git clone [https://github.com/your-username/ila-oadm-data-merger.git](https://github.com/your-username/ila-oadm-data-merger.git)
cd ila-oadm-data-merger
python -m venv venv
# Windows activate:
venv\Scripts\activate
```
Step 2: Install Dependencies & PyInstaller
Bash
```
pip install pandas openpyxl toolz pyinstaller
```
Step 3: Compile Code into Standalone .exe

App ko ek single windows executable file mein convert karne ke liye niche di gayi terminal command run karein:
Bash
```
pyinstaller --noconsole --onefile --name="ILA_OADM_Merger" main.py
```
Command Breakdown:

    --onefile: Saare scripts aur dependencies ko ek single .exe file mein pack karta hai.

    --noconsole: App chalte waat piche black color ka CMD/Console window open nahi hoga (sirf clean GUI dikhegi).

    --name: Aapki executable file ka customized naam set karta hai.

Compilation khatam hone ke baad, aapko aapki desktop ready executable file dist/ folder ke andar ILA_OADM_Merger.exe naam se mil jayegi.

## 📂 Expected Input Formats

To ensure flawless pipeline executions, please verify that your dump files follow this structure:
Tracker Type	File Name Wildcard	Target Worksheets/Extensions	Required Core Columns
SOC Tracker	MFL*.xlsx	Sheet Name: Worksheet	SAP ID, Circle, Alarm Type, Facility
HT Logs	Export-*.xlsx	Sheet Name: Data	NUMBER, ASSIGNMENT, TITLE
DG Manuals	*.csv (Recursive)	Format: .csv	sapid or alarm_name, sitetype (for SFL)
📄 License

Distributed under the MIT License. See LICENSE for more information.


---

### 💡 Pro-Tip for your PyInstaller build:
Jab aap `--onefile` use karte hain, toh app launch hone mein 2-3 seconds ka thoda sa time le sakti hai kyunki backend par saari libraries temp directory mein extract hoti hain. Yeh normal behavior hai, isse ghabrane ki baat nahi hai!
