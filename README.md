# SA-MP Automatic War Evidence

A simple, automated Python tool that generates faction war evidence (Excel and HTML reports) for SA-MP RPG servers. 

Instead of manually checking every player's seconds and score, this script does it for you. It reads the war logs directly from the server's website, checks if players were excused or inactive, excludes staff members, and automatically calculates the fines or faction warns based on the results.

---

## How to Download & Use

If you are not a developer and just want to use the program, you don't need to install Python or touch any code!

1. Go to the [Releases](https://github.com/Madalin02/sa-mp-automatic-war-evidence/releases) page of this repository.
2. Download the latest `.exe` file.
3. Double-click the file to run it.
4. Follow the instructions on the screen (enter your faction, date, and browser cookie).

---

## For Developers

If you want to run the script from the source code or contribute to the project, follow these steps:

### Prerequisites
- Python 3.8 or newer.

### Setup
1. Clone this repository:
   ```bash
   git clone [https://github.com/Madalin02/sa-mp-automatic-war-evidence.git](https://github.com/Madalin02/sa-mp-automatic-war-evidence.git)
   cd sa-mp-automatic-war-evidence
   ```
2. Create and activate a virtual environment:
   ```bash
   python -m venv .venv
   
   # Windows:
   .\.venv\Scripts\activate
   
   # Linux/macOS:
   source .venv/bin/activate
   ```
3. Install the required dependencies:
   ```bash
   pip install -r requirements.txt
   ```
4. Run the script:
   ```bash
   python AutoWarEvidence.py
   ```

---

## How it works

When you run the app, it will ask you for:
- **Faction:** Choose your faction from the list.
- **Score Penalties:** Choose if you want to sanction players for bad scores (e.g., -5, -10, -15) on lost wars.
- **Cookie:** Your active browser session cookie to access the logs securely.
- **Date:** The date of the wars.
- **Minimum Seconds:** The required seconds for each war.

The tool will output two files in the same folder: an Excel file (`.xlsx`) for your records, and an HTML file (`.html`) ready to be copy-pasted directly onto the forum.

---

## Contributing
Contributions are always welcome! If you want to improve this tool, please check the `CONTRIBUTING.md` file for details on how to submit a Pull Request.

## License
This project is licensed under the MIT License - see the `LICENSE` file for details.